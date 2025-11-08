import os
import sys
import io
import tempfile
import shutil
import traceback
import re
from pathlib import Path
from PIL import Image
import img2pdf
from PyQt6.QtGui import QIcon
from ebooklib import epub

def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QListWidget, QFileDialog, QMessageBox, QTextEdit, QCheckBox, QRadioButton, QButtonGroup
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal

# 支持扩展名
IMAGE_EXTS = ('.jpg', '.jpeg', '.png', '.webp')

def safe_filename(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|]', '_', name)

def numeric_sort_key(name: str):
    nums = re.findall(r"\d+", name)
    if nums:
        return [int(n) for n in nums] + [name]
    return [999999, name]

def gather_images(folder: Path):
    folder = Path(folder)
    if not folder.exists() or not folder.is_dir():
        return []
    files = [p for p in folder.rglob('*') if p.is_file() and p.suffix.lower() in IMAGE_EXTS]
    files.sort(key=lambda p: (str(p.relative_to(folder.parent if folder.parent else folder)).lower(), numeric_sort_key(p.name)))
    return files

def images_to_temp_jpegs(image_paths):
    tmp_dir = tempfile.mkdtemp(prefix="manga_tmp_")
    tmp_files = []
    try:
        for p in image_paths:
            p = Path(p)
            if not p.exists():
                continue
            if p.suffix.lower() in ('.jpg', '.jpeg', '.png'):
                tmp_files.append(str(p))
            else:
                img = Image.open(p).convert('RGB')
                out = os.path.join(tmp_dir, p.stem + '.jpg')
                img.save(out, format='JPEG', quality=95)
                tmp_files.append(out)
        return tmp_dir, tmp_files
    except Exception:
        shutil.rmtree(tmp_dir, ignore_errors=True)
        raise

def convert_images_to_pdf(image_paths, out_pdf_path):
    tmp_dir = None
    try:
        paths = [str(p) for p in image_paths]
        tmp_dir, tmp_files = images_to_temp_jpegs(paths)
        with open(out_pdf_path, 'wb') as f:
            f.write(img2pdf.convert(tmp_files))
        return True, None
    except Exception as e:
        return False, traceback.format_exc()
    finally:
        if tmp_dir:
            shutil.rmtree(tmp_dir, ignore_errors=True)

def convert_folder_to_epub(image_paths, out_epub_path, book_title=None, volume_title=None):
    try:
        paths = [Path(p) for p in image_paths if Path(p).exists()]
        if not paths:
            return False, "没有图片"
        paths.sort(key=lambda p: numeric_sort_key(p.name))

        book = epub.EpubBook()
        book.set_identifier(safe_filename(book_title or "manga"))
        book.set_title(volume_title or book_title or "Manga")
        book.set_language("zh")

        first = paths[0]
        try:
            with open(first, 'rb') as f:
                raw = f.read()
            book.set_cover("cover.jpg", raw)
        except Exception:
            img = Image.open(first).convert("RGB")
            buf = io.BytesIO()
            img.save(buf, format="JPEG")
            book.set_cover("cover.jpg", buf.getvalue())

        for i, p in enumerate(paths, start=1):
            p = Path(p)
            img = Image.open(p).convert("RGB")
            buf = io.BytesIO()
            img.save(buf, format="JPEG")
            img_bytes = buf.getvalue()

            img_item = epub.EpubItem(
                uid=f"image_{i}",
                file_name=f"images/{i:03d}.jpg",
                media_type="image/jpeg",
                content=img_bytes
            )
            book.add_item(img_item)

            html = f'<html><body style="margin:0;padding:0;background:#000;"><img src="images/{i:03d}.jpg" style="width:100%;height:auto;display:block;margin:0 auto;"/></body></html>'
            chap = epub.EpubHtml(title=f"Page {i}", file_name=f"page_{i:03d}.xhtml", content=html)
            book.add_item(chap)

        pages = [item for item in book.get_items() if isinstance(item, epub.EpubHtml)]
        book.toc = tuple(pages)
        book.spine = ['nav'] + pages
        book.add_item(epub.EpubNcx())
        book.add_item(epub.EpubNav())
        epub.write_epub(out_epub_path, book)
        return True, None
    except Exception:
        return False, traceback.format_exc()

class Worker(QThread):
    log = pyqtSignal(str)
    finished = pyqtSignal()

    def __init__(self, folders, out_dir: str, export_pdf: bool, do_merge: bool):
        super().__init__()
        self.folders = folders
        self.out_dir = Path(out_dir)
        self.export_pdf = export_pdf
        self.do_merge = do_merge
        self._stop = False

    def stop(self):
        self._stop = True

    def run(self):
        try:
            if not self.folders:
                self.log.emit("没有要处理的卷。")
                self.finished.emit()
                return

            groups = {}
            for f in self.folders:
                parent = f.parent
                if parent and parent.name.lower() in ["单行本", "單行本", "默認", "其他汉化版"]:
                    manga_name = parent.parent.name if parent.parent else "Unknown"
                else:
                    manga_name = parent.name if parent else "Unknown"
                groups.setdefault(manga_name, []).append(f)

            for manga_name, vols in groups.items():
                manga_out = self.out_dir / safe_filename(manga_name)
                manga_out.mkdir(parents=True, exist_ok=True)

                self.log.emit(f"开始处理漫画：{manga_name}，共 {len(vols)} 卷")
                for vol in vols:
                    if self._stop:
                        self.log.emit("已取消操作。")
                        self.finished.emit()
                        return
                    vol_name = vol.name
                    vol_name = re.sub(r'^\d+\s*', '', vol_name)
                    self.log.emit(f"  处理卷：{vol_name}")
                    imgs = gather_images(vol)
                    if not imgs:
                        self.log.emit(f"    ⚠️ 未找到图片：{vol} （检查子目录与扩展名）")
                        continue
                    out_path = manga_out / f"{vol_name}{'.pdf' if self.export_pdf else '.epub'}"
                    if self.export_pdf:
                        ok, err = convert_images_to_pdf(imgs, str(out_path))
                    else:
                        ok, err = convert_folder_to_epub(imgs, str(out_path), book_title=manga_name, volume_title=vol_name)
                    if ok and out_path.exists():
                        self.log.emit(f"    ✅ 输出成功：{out_path}")
                    else:
                        self.log.emit(f"    ❌ 输出失败：{out_path} 错误: {err}")

                if self.do_merge:
                    self.log.emit(f"  开始合并 {manga_name} 的所有卷...")
                    all_imgs = []
                    for vol in vols:
                        imgs = gather_images(vol)
                        all_imgs.extend(imgs)
                    if not all_imgs:
                        self.log.emit("    ⚠️ 合并失败：没有图片")
                        continue
                    merged_name = manga_out / f"总集{'.pdf' if self.export_pdf else '.epub'}"
                    if self.export_pdf:
                        ok, err = convert_images_to_pdf(all_imgs, str(merged_name))
                    else:
                        ok, err = convert_folder_to_epub(all_imgs, str(merged_name), book_title=manga_name, volume_title=f"{manga_name} 总集")
                    if ok and merged_name.exists():
                        self.log.emit(f"    ✅ 合并输出成功：{merged_name}")
                    else:
                        self.log.emit(f"    ❌ 合并输出失败：{merged_name} 错误: {err}")

            self.log.emit("全部处理完成。")
        except Exception as e:
            self.log.emit(f"线程异常：{e}\n{traceback.format_exc()}")
        finally:
            self.finished.emit()

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Manga Converter")
        self.setWindowIcon(QIcon(resource_path("aa965-5qyrc.png")))
        self.resize(920, 640)
        self.setAcceptDrops(True)

        v = QVBoxLayout(self)
        h_top = QHBoxLayout()
        self.btn_add_parent = QPushButton("添加父文件夹（扫描子文件夹）")
        self.btn_add_volume = QPushButton("添加卷文件夹（单卷）")
        h_top.addWidget(self.btn_add_parent)
        h_top.addWidget(self.btn_add_volume)
        v.addLayout(h_top)

        middle = QHBoxLayout()
        left_v = QVBoxLayout()
        left_v.addWidget(QLabel("已添加的卷："))
        self.list_widget = QListWidget()
        left_v.addWidget(self.list_widget)
        h_left_btn = QHBoxLayout()
        self.btn_remove = QPushButton("移除选中")
        self.btn_clear = QPushButton("清空列表")
        h_left_btn.addWidget(self.btn_remove)
        h_left_btn.addWidget(self.btn_clear)
        left_v.addLayout(h_left_btn)
        middle.addLayout(left_v, 45)

        right_v = QVBoxLayout()
        right_v.addWidget(QLabel("日志："))
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        right_v.addWidget(self.log)
        middle.addLayout(right_v, 55)
        v.addLayout(middle)

        h_opts = QHBoxLayout()
        h_format = QHBoxLayout()
        h_format.addWidget(QLabel("导出格式："))
        self.radio_pdf = QRadioButton("PDF")
        self.radio_epub = QRadioButton("EPUB")
        self.radio_pdf.setChecked(True)
        self.btn_group = QButtonGroup()
        self.btn_group.addButton(self.radio_pdf)
        self.btn_group.addButton(self.radio_epub)
        h_format.addWidget(self.radio_pdf)
        h_format.addWidget(self.radio_epub)
        h_opts.addLayout(h_format, 40)

        self.chk_merge = QCheckBox("同时生成总集（总集.pdf / 总集.epub）")
        h_opts.addWidget(self.chk_merge, 40)

        out_v = QVBoxLayout()
        self.lbl_out = QLabel("输出目录： 未选择（默认优先 S:\\Manga）")
        out_v.addWidget(self.lbl_out)
        self.btn_choose_out = QPushButton("选择输出目录")
        out_v.addWidget(self.btn_choose_out)
        h_opts.addLayout(out_v, 60)
        v.addLayout(h_opts)

        h_run = QHBoxLayout()
        self.btn_start_stop = QPushButton("开始转换")
        h_run.addWidget(self.btn_start_stop)
        v.addLayout(h_run)

        self.setStyleSheet("""
        QWidget { 
            background: #f5f5f7; 
            font-family: "Helvetica Neue", "Microsoft YaHei"; 
            color: #333; 
            font-size: 13px; 
        }

        /* 按钮 */
        QPushButton { 
            background: #ffb6c1; 
            color: white; 
            border-radius: 10px; 
            padding: 8px 12px; 
        }
        QPushButton:hover { 
            background: #ff99b2; 
        }

        /* 列表和文本框 */
        QListWidget, QTextEdit { 
            background: white; 
            border: 1px solid #dcdcdc; 
            border-radius: 10px; 
            padding: 6px; 
        }

        /* 滚动条  */
        QScrollBar:vertical {
            width: 12px;
            background: #f0f0f0;
            margin: 0px 0px 0px 0px;
            border-radius: 6px;
        }
        QScrollBar::handle:vertical {
            background: #ffb6c1;
            min-height: 30px;
            border-radius: 6px;
        }
        QScrollBar::handle:vertical:hover {
            background: #ff99b2;
        }
        QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
            height: 0px;
        }
        QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
            background: none;
        }

        /* 单选按钮 - PDF/EPUB */
        QRadioButton::indicator {
            width: 18px;
            height: 18px;
            border-radius: 9px;
            border: 2px solid #ffb6c1;
            background: white;
        }
        QRadioButton::indicator:checked {
            background: #ffb6c1;
            border: 2px solid #ff99b2;
        }

        /* 复选框 - 同时生成总集 */
        QCheckBox::indicator {
            width: 18px;
            height: 18px;
            border-radius: 5px;
            border: 2px solid #ffb6c1;
            background: white;
        }
        QCheckBox::indicator:checked {
            background: #ffb6c1;
            border: 2px solid #ff99b2;
        }
        """)

        self.folders = []
        default_out = Path("S:/Manga") if Path("S:/").exists() else Path.cwd() / "output"
        self.out_dir = str(default_out)
        self.lbl_out.setText(f"输出目录： {self.out_dir}")
        self.worker = None

        self.btn_add_parent.clicked.connect(self.on_add_parent)
        self.btn_add_volume.clicked.connect(self.on_add_volume)
        self.btn_remove.clicked.connect(self.on_remove)
        self.btn_clear.clicked.connect(self.on_clear)
        self.btn_choose_out.clicked.connect(self.on_choose_out)
        self.btn_start_stop.clicked.connect(self.on_start_stop)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        added = 0
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if os.path.isdir(path):
                p = Path(path)
                if any(p.rglob('*' + ext) for ext in IMAGE_EXTS):
                    if str(path) not in [str(pth) for pth in self.folders]:
                        self.folders.append(Path(path))
                        self.list_widget.addItem(str(path))
                        self.log_msg(f"已拖入：{path}")
                        added += 1
                else:
                    self.log_msg(f"拖入目录不包含图片：{path}")
        if added == 0:
            self.log_msg("拖入的文件夹已存在或没有有效子文件。")

    def on_add_parent(self):
        folder = QFileDialog.getExistingDirectory(self, "选择父文件夹（会添加其所有子文件夹）")
        if not folder:
            return
        added = 0
        base = Path(folder)
        for sub in sorted(set([p.parent for p in base.rglob('*') if p.is_file()])):
            sub = Path(sub)
            imgs = [p for p in sub.rglob('*') if p.is_file() and p.suffix.lower() in IMAGE_EXTS]
            if not imgs:
                continue
            if str(sub) not in [str(p) for p in self.folders]:
                self.folders.append(sub)
                self.list_widget.addItem(str(sub))
                self.log_msg(f"添加子文件夹：{sub}")
                added += 1
        self.log_msg(f"从父文件夹添加完成，共添加 {added} 个卷。")

    def on_add_volume(self):
        folder = QFileDialog.getExistingDirectory(self, "选择卷文件夹（单卷）")
        if not folder:
            return
        p = Path(folder)
        imgs = [q for q in p.rglob('*') if q.is_file() and q.suffix.lower() in IMAGE_EXTS]
        if not imgs:
            self.log_msg("该目录未检测到支持的图片（jpg/png/webp）。请检查是否为正确的卷目录。")
            return
        if str(folder) in [str(p) for p in self.folders]:
            self.log_msg("该卷已在列表中。")
            return
        self.folders.append(Path(folder))
        self.list_widget.addItem(str(folder))
        self.log_msg(f"已添加卷：{folder}")

    def on_remove(self):
        sel = self.list_widget.selectedItems()
        for item in sel:
            text = item.text()
            self.folders = [p for p in self.folders if str(p) != text]
            self.list_widget.takeItem(self.list_widget.row(item))
            self.log_msg(f"已移除：{text}")

    def on_clear(self):
        self.folders.clear()
        self.list_widget.clear()
        self.log_msg("已清空列表。")

    def on_choose_out(self):
        d = QFileDialog.getExistingDirectory(self, "选择输出目录", self.out_dir)
        if d:
            self.out_dir = d
            self.lbl_out.setText(f"输出目录： {self.out_dir}")
            self.log_msg(f"输出目录设置为：{self.out_dir}")

    def on_start_stop(self):
        if self.worker and self.worker.isRunning():
            # 停止任务
            self.worker.stop()
            self.btn_start_stop.setText("开始转换")
            self.log_msg("停止请求已发送，当前任务完成后将停止。")
        else:
            # 开始任务
            if not self.folders:
                QMessageBox.warning(self, "提示", "请先添加至少一个卷文件夹。")
                return
            if not self.out_dir:
                QMessageBox.warning(self, "提示", "请选择输出目录。")
                return
            export_pdf = self.radio_pdf.isChecked()
            do_merge = self.chk_merge.isChecked()
            self.log_msg("开始转换任务...")
            folders_copy = [Path(p) for p in self.folders]
            self.worker = Worker(folders_copy, self.out_dir, export_pdf, do_merge)
            self.worker.log.connect(self.log_msg)
            self.worker.finished.connect(self.on_worker_finished)
            self.worker.start()
            self.btn_start_stop.setText("停止")

    def on_worker_finished(self):
        self.log_msg("所有后台任务已结束。")
        self.btn_start_stop.setText("开始转换")

    def log_msg(self, text: str):
        self.log.append(text)
        self.log.verticalScrollBar().setValue(self.log.verticalScrollBar().maximum())

def main():
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
