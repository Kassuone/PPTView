import sys
import os
import shutil
import tempfile
import pythoncom
import win32com.client
try:
    from PIL import Image
    HAS_PIL = True
except Exception:
    HAS_PIL = False
from PyQt5.QtWidgets import (QApplication, QMainWindow, QFileDialog, QListWidget, 
                             QListWidgetItem, QMenu, QMessageBox, QVBoxLayout, 
                             QWidget, QLabel, QProgressBar, QAbstractItemView,
                             QDialog, QPushButton, QScrollArea, QHBoxLayout, QInputDialog)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QSize
from PyQt5.QtGui import QIcon, QPixmap, QFont, QCursor, QPainter, QColor

class PPTConverterThread(QThread):
    progress_signal = pyqtSignal(int)
    finished_signal = pyqtSignal(list)
    error_signal = pyqtSignal(str)

    def __init__(self, ppt_path, temp_dir, width=1920, height=1080):
        super().__init__()
        self.ppt_path = ppt_path
        self.temp_dir = temp_dir
        self.width = width
        self.height = height

    def run(self):
        try:
            # 在线程中初始化COM库
            pythoncom.CoInitialize()
            
            # 启动 PowerPoint 进程（不可见模式）
            powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            # 某些版本可能需要设置为可见才能正常导出，视情况而定，这里设为最小化以防干扰
            # powerpoint.Visible = 1 
            
            # 以只读方式打开，不显示窗口
            presentation = powerpoint.Presentations.Open(self.ppt_path, WithWindow=False, ReadOnly=True)
            
            image_paths = []
            total_slides = presentation.Slides.Count
            
            for i, slide in enumerate(presentation.Slides):
                # 构造输出文件名 (slide_1.png, slide_2.png...)
                image_name = f"slide_{i+1}.png"
                full_path = os.path.join(self.temp_dir, image_name)
                
                # 导出该页为PNG，指定宽高可提高清晰度，这里用默认比例，放大倍数以保证清晰度
                # 1920x1080 约等于 ScaleWidth=1920
                # 使用传入的分辨率
                slide.Export(full_path, "PNG", int(self.width), int(self.height))
                image_paths.append(full_path)
                
                # 发送进度
                self.progress_signal.emit(int((i + 1) / total_slides * 100))

            presentation.Close()
            # 不关闭 Application，因为打开关闭太慢，或者根据需求关闭
            # powerpoint.Quit() 
            
            self.finished_signal.emit(image_paths)
            
        except Exception as e:
            self.error_signal.emit(str(e))
        finally:
            pythoncom.CoUninitialize()


class ZoomDialog(QDialog):
    def __init__(self, image_path, parent=None):
        super().__init__(parent)
        self.setWindowTitle("放大预览")
        self.image_path = image_path
        self.zoom = 1.0

        self.scroll = QScrollArea()
        self.label = QLabel()
        self.label.setAlignment(Qt.AlignCenter)
        self.scroll.setWidget(self.label)
        self.scroll.setWidgetResizable(True)

        btn_zoom_in = QPushButton("放大")
        btn_zoom_out = QPushButton("缩小")
        btn_close = QPushButton("关闭")

        btn_zoom_in.clicked.connect(self.zoom_in)
        btn_zoom_out.clicked.connect(self.zoom_out)
        btn_close.clicked.connect(self.close)

        btn_layout = QHBoxLayout()
        btn_layout.addWidget(btn_zoom_in)
        btn_layout.addWidget(btn_zoom_out)
        btn_layout.addWidget(btn_close)

        main_layout = QVBoxLayout(self)
        main_layout.addWidget(self.scroll)
        main_layout.addLayout(btn_layout)

        self.load_image()

    def load_image(self):
        pix = QPixmap(self.image_path)
        self.orig_pix = pix
        self.update_pixmap()

    def update_pixmap(self):
        if self.orig_pix.isNull():
            return
        size = self.orig_pix.size()
        new_size = QSize(int(size.width() * self.zoom), int(size.height() * self.zoom))
        scaled = self.orig_pix.scaled(new_size, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        self.label.setPixmap(scaled)

    def zoom_in(self):
        self.zoom *= 1.25
        self.update_pixmap()

    def zoom_out(self):
        self.zoom /= 1.25
        self.update_pixmap()

class PPTViewer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PPT 高保真预览工具")
        self.resize(1000, 700)
        self.temp_dir = tempfile.mkdtemp() # 创建临时文件夹存放预览图
        self.image_paths = [] # 存储当前的图片路径列表
        # 默认导出分辨率
        self.export_width = 1920
        self.export_height = 1080
        # 单击多选开关状态（False 为默认，需要按 Ctrl 进行多选）
        self.click_multiselect_enabled = False
        
        self.init_ui()
        self.apply_styles()

    def init_ui(self):
        # 主布局
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # 顶部提示区
        self.lbl_status = QLabel("请导入 PPT/PPTX 文件以开始预览...")
        self.lbl_status.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.lbl_status)

        # 进度条 (默认隐藏)
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        # 预览列表区域 (使用 QListWidget 实现缩略图网格)
        self.list_widget = QListWidget()
        self.list_widget.setIconSize(QSize(280, 158)) # 设置缩略图大小 16:9 比例
        self.list_widget.setViewMode(QListWidget.IconMode) # 图标模式
        self.list_widget.setResizeMode(QListWidget.Adjust) # 自适应布局
        self.list_widget.setSpacing(15) # 间距
        self.list_widget.setSelectionMode(QAbstractItemView.ExtendedSelection) # 允许按Ctrl/Shift多选
        self.list_widget.setContextMenuPolicy(Qt.CustomContextMenu) # 启用右键菜单
        self.list_widget.customContextMenuRequested.connect(self.show_context_menu)
        # 支持双击放大查看
        self.list_widget.itemDoubleClicked.connect(self.open_zoom_view)
        # 选择变更用于控制按钮状态
        self.list_widget.itemSelectionChanged.connect(self.on_selection_changed)
        layout.addWidget(self.list_widget)

        # 菜单栏 / 工具栏
        toolbar = self.addToolBar("文件")
        # 保存为实例变量以便根据选择状态启用/禁用
        self.action_import = toolbar.addAction("导入 PPT 文件")
        self.action_import.triggered.connect(self.import_ppt)
        self.action_set_res = toolbar.addAction("设置导出分辨率")
        self.action_set_res.triggered.connect(self.set_export_resolution)
        # 添加单击多选开关：开启后单击即可多选（无需按 Ctrl），关闭恢复为 ExtendedSelection（保留 Ctrl 多选）
        self.action_toggle_click_multiselect = toolbar.addAction("单击多选（无需Ctrl）")
        self.action_toggle_click_multiselect.setCheckable(True)
        self.action_toggle_click_multiselect.toggled.connect(self.toggle_click_multiselect)
        self.action_export_pdf = toolbar.addAction("导出为 PDF")
        self.action_export_pdf.triggered.connect(self.export_all_as_pdf)
        self.action_export_selected = toolbar.addAction("导出选中")
        self.action_export_selected.triggered.connect(self.show_export_menu)

        # 允许在主窗口上拖放文件
        self.setAcceptDrops(True)


    def apply_styles(self):
        # 美化界面 (QSS)
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f0f2f5;
            }
            QListWidget {
                background-color: white;
                border: 1px solid #dcdcdc;
                border-radius: 8px;
                padding: 10px;
            }
            QListWidget::item {
                background-color: #ffffff;
                border: 1px solid #eaeaea;
                border-radius: 5px;
                padding: 5px;
            }
            QListWidget::item:selected {
                background-color: #e6f7ff;
                border: 2px solid #1890ff;
            }
            QListWidget::item:hover {
                background-color: #fafafa;
            }
            QLabel {
                font-family: 'Microsoft YaHei';
                font-size: 14px;
                color: #333;
                padding: 10px;
            }
            QProgressBar {
                border: none;
                background-color: #e0e0e0;
                height: 4px;
                text-align: center;
            }
            QProgressBar::chunk {
                background-color: #1890ff;
            }
        """)

    def import_ppt(self, file_path: str = None):
        """导入 PPT 文件，支持传入路径或弹窗选择。"""
        if not file_path:
            file_path, _ = QFileDialog.getOpenFileName(
                self, "选择 PPT 文件", "", "PowerPoint Files (*.ppt *.pptx)"
            )
            if not file_path:
                return

        # 清空当前显示
        self.list_widget.clear()
        self.image_paths = []
        self.lbl_status.setText(f"正在解析: {os.path.basename(file_path)} (这可能需要几秒钟)...")
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(True)
        self.list_widget.setEnabled(False)

        # 启动后台线程转换，使用当前设置的分辨率
        self.converter_thread = PPTConverterThread(file_path, self.temp_dir, self.export_width, self.export_height)
        self.converter_thread.progress_signal.connect(self.update_progress)
        self.converter_thread.finished_signal.connect(self.load_previews)
        self.converter_thread.error_signal.connect(self.show_error)
        self.converter_thread.start()

    def update_progress(self, val):
        self.progress_bar.setValue(val)

    def show_error(self, err_msg):
        self.progress_bar.setVisible(False)
        self.list_widget.setEnabled(True)
        QMessageBox.critical(self, "错误", f"解析 PPT 失败:\n请确保已安装 Microsoft Office。\n\n错误信息: {err_msg}")
        self.lbl_status.setText("导入失败")

    def load_previews(self, image_paths):
        self.progress_bar.setVisible(False)
        self.list_widget.setEnabled(True)
        self.lbl_status.setText(f"预览加载完成，共 {len(image_paths)} 页。可在图片上【右键】导出。")
        self.image_paths = image_paths

        # 将图片加载到界面
        for idx, img_path in enumerate(image_paths):
            pixmap = QPixmap(img_path)
            
            # 创建图标
            icon = QIcon(pixmap)
            item = QListWidgetItem(icon, f"第 {idx+1} 页")
            # 存储对应的图片路径在 UserRole 数据中，方便后续调用
            item.setData(Qt.UserRole, img_path) 
            self.list_widget.addItem(item)

        # 更新工具栏按钮状态
        self.update_actions_state()

    def show_context_menu(self, position):
        # 如果没有选中项，不显示菜单
        if not self.list_widget.itemAt(position):
            return

        menu = QMenu()
        
        # 获取选中项
        selected_items = self.list_widget.selectedItems()
        count = len(selected_items)
        
        action_save_one = None
        action_save_all = menu.addAction("保存所有幻灯片为图片")
        action_export_pdf = None
        action_export_pdf_all = menu.addAction("导出所有为 PDF")
        
        if count == 1:
            action_save_one = menu.addAction("保存当前页")
            action_export_pdf = menu.addAction("导出当前页为 PDF")
        elif count > 1:
            action_save_one = menu.addAction(f"保存选中的 {count} 页")
            action_export_pdf = menu.addAction(f"导出选中的 {count} 页为 PDF")
            
        action = menu.exec_(self.list_widget.mapToGlobal(position))
        
        if action == action_save_one:
            self.save_images(selected_items)
        elif action == action_save_all:
            # 获取所有项目
            all_items = [self.list_widget.item(i) for i in range(self.list_widget.count())]
            self.save_images(all_items)
        elif action == action_export_pdf:
            # 导出选中或当前为 PDF
            self.save_as_pdf(selected_items)
        elif action == action_export_pdf_all:
            all_items = [self.list_widget.item(i) for i in range(self.list_widget.count())]
            self.save_as_pdf(all_items)

    def save_images(self, items):
        if not items:
            return
            
        # 如果是单张，直接弹窗保存文件
        if len(items) == 1:
            src_path = items[0].data(Qt.UserRole)
            save_path, _ = QFileDialog.getSaveFileName(
                self, "保存图片", f"{items[0].text()}.png", "PNG Images (*.png)"
            )
            if save_path:
                try:
                    shutil.copy(src_path, save_path)
                    QMessageBox.information(self, "成功", "保存成功！")
                except Exception as e:
                    QMessageBox.warning(self, "错误", str(e))
        else:
            # 如果是多张，选择文件夹
            folder_path = QFileDialog.getExistingDirectory(self, "选择保存文件夹")
            if folder_path:
                try:
                    count = 0
                    for item in items:
                        src_path = item.data(Qt.UserRole)
                        file_name = f"{item.text().replace(' ', '_')}.png" # 例如 第_1_页.png
                        dst_path = os.path.join(folder_path, file_name)
                        shutil.copy(src_path, dst_path)
                        count += 1
                    QMessageBox.information(self, "成功", f"成功保存 {count} 张图片！")
                except Exception as e:
                    QMessageBox.warning(self, "错误", str(e))

    def set_export_resolution(self):
        # 让用户输入宽和高
        w, ok1 = QInputDialog.getInt(self, "设置宽度", "PNG 导出宽度 (像素):", value=self.export_width, min=100, max=10000)
        if not ok1:
            return
        h, ok2 = QInputDialog.getInt(self, "设置高度", "PNG 导出高度 (像素):", value=self.export_height, min=100, max=10000)
        if not ok2:
            return
        self.export_width = w
        self.export_height = h
        QMessageBox.information(self, "已设置", f"导出分辨率已设置为 {w} x {h}。下次导入/转换将使用该分辨率。")

    def toggle_click_multiselect(self, checked: bool):
        """切换单击多选模式。

        - checked == True: 允许单击直接多选（无需 Ctrl），使用 MultiSelection
        - checked == False: 恢复 ExtendedSelection（单击单选，Ctrl 可多选）
        """
        self.click_multiselect_enabled = checked
        if checked:
            self.list_widget.setSelectionMode(QAbstractItemView.MultiSelection)
            self.lbl_status.setText("已启用：单击多选（无需按 Ctrl）")
        else:
            self.list_widget.setSelectionMode(QAbstractItemView.ExtendedSelection)
            # 关闭多选时取消所有已选中的项目
            self.list_widget.clearSelection()
            self.lbl_status.setText("已关闭单击多选：使用 Ctrl/Shift 进行多选")

        # 更新工具栏按钮状态（根据是否多选或已选择多项）
        self.update_actions_state()

    def on_selection_changed(self):
        """当列表选择变化时，更新其他按钮的可用状态。

        规则：当 `单击多选` 开启，或选中项数量大于 1 时，除 `导出选中` 和 `单击多选` 外的按钮均禁用。
        """
        selected_count = len(self.list_widget.selectedItems())
        # 如果是通过 Ctrl/Shift 多选且选中超过 1，或单击多选被打开
        self.update_actions_state()

    def update_actions_state(self):
        """根据当前多选开关或选中数量来启用/禁用工具栏按钮。

        保持 `self.action_toggle_click_multiselect`（用于关闭多选）和
        `self.action_export_selected` 始终可用；其它按钮在多选场景下禁用。
        """
        selected_count = len(self.list_widget.selectedItems())
        multi_mode = self.click_multiselect_enabled or (selected_count > 1)

        # 哪些动作需要在多选时被禁用（但保留 toggle 与 导出选中）
        actions_to_control = [getattr(self, 'action_import', None), getattr(self, 'action_set_res', None), getattr(self, 'action_export_pdf', None)]
        for act in actions_to_control:
            if act:
                act.setEnabled(not multi_mode)

    def dragEnterEvent(self, event):
        """接收拖入的文件（支持 .ppt/.pptx）。"""
        mime = event.mimeData()
        if mime.hasUrls():
            for url in mime.urls():
                if url.isLocalFile() and url.toLocalFile().lower().endswith(('.ppt', '.pptx')):
                    event.acceptProposedAction()
                    return
        event.ignore()

    def dropEvent(self, event):
        """处理文件放下，取第一个 PPT 文件并导入。"""
        mime = event.mimeData()
        if mime.hasUrls():
            for url in mime.urls():
                if url.isLocalFile():
                    path = url.toLocalFile()
                    if path.lower().endswith(('.ppt', '.pptx')):
                        # 使用导入函数直接处理此路径
                        self.import_ppt(path)
                        event.acceptProposedAction()
                        return
        event.ignore()

    def open_zoom_view(self, item):
        img = item.data(Qt.UserRole)
        if not img or not os.path.exists(img):
            QMessageBox.warning(self, "错误", "图片文件不存在。")
            return
        dlg = ZoomDialog(img, self)
        dlg.resize(900, 600)
        dlg.exec_()

    def save_as_pdf(self, items):
        if not items:
            return
        if not HAS_PIL:
            QMessageBox.warning(self, "缺少依赖", "PDF 导出需要安装 Pillow 库，请运行: pip install pillow")
            return

        # 保持原始幻灯片顺序：按 list_widget 中的顺序过滤（通过图片路径判断）
        selected_paths = set([it.data(Qt.UserRole) for it in items])
        ordered = [self.list_widget.item(i) for i in range(self.list_widget.count()) if self.list_widget.item(i).data(Qt.UserRole) in selected_paths]

        save_path, _ = QFileDialog.getSaveFileName(self, "保存为 PDF", "slides.pdf", "PDF Files (*.pdf)")
        if not save_path:
            return

        try:
            pil_images = []
            for idx, item in enumerate(ordered):
                img_path = item.data(Qt.UserRole)
                img = Image.open(img_path).convert('RGB')
                pil_images.append(img)

            if not pil_images:
                QMessageBox.warning(self, "错误", "没有图片可以导出。")
                return

            first, rest = pil_images[0], pil_images[1:]
            first.save(save_path, save_all=True, append_images=rest)
            QMessageBox.information(self, "成功", "PDF 导出成功！")
        except Exception as e:
            QMessageBox.warning(self, "错误", str(e))

    def show_export_menu(self):
        # 弹出导出选项菜单（图片 / 合并 PDF / 分开的 PDF）
        selected_items = self.list_widget.selectedItems()
        if not selected_items:
            QMessageBox.information(self, "提示", "请先选择要导出的幻灯片（可多选）。")
            return

        menu = QMenu(self)
        act_img = menu.addAction("导出为图片")
        act_pdf_merge = menu.addAction("导出为合并 PDF")
        act_pdf_separate = menu.addAction("导出为分开的 PDF（每页一个）")

        action = menu.exec_(QCursor.pos())
        if action == act_img:
            self.export_selected_as_images(selected_items)
        elif action == act_pdf_merge:
            self.save_as_pdf(selected_items)
        elif action == act_pdf_separate:
            self.export_selected_as_separate_pdfs(selected_items)

    def export_selected_as_images(self, items):
        # 将选中的项按当前列表顺序导出为 PNG（单个文件夹，按 slide_1.png 命名）
        if not items:
            return

        folder_path = QFileDialog.getExistingDirectory(self, "选择保存文件夹")
        if not folder_path:
            return

        try:
            selected_paths = set([it.data(Qt.UserRole) for it in items])
            ordered = [self.list_widget.item(i) for i in range(self.list_widget.count()) if self.list_widget.item(i).data(Qt.UserRole) in selected_paths]
            count = 0
            for idx, item in enumerate(ordered, start=1):
                src_path = item.data(Qt.UserRole)
                file_name = f"slide_{idx}.png"
                dst_path = os.path.join(folder_path, file_name)
                shutil.copy(src_path, dst_path)
                count += 1
            QMessageBox.information(self, "成功", f"成功保存 {count} 张图片！")
        except Exception as e:
            QMessageBox.warning(self, "错误", str(e))

    def export_selected_as_separate_pdfs(self, items):
        # 将选中的每个幻灯片分别导出为独立 PDF（需要 Pillow）
        if not items:
            return
        if not HAS_PIL:
            QMessageBox.warning(self, "缺少依赖", "单页 PDF 导出需要安装 Pillow 库，请运行: pip install pillow")
            return

        folder_path = QFileDialog.getExistingDirectory(self, "选择保存文件夹")
        if not folder_path:
            return

        try:
            selected_paths = set([it.data(Qt.UserRole) for it in items])
            ordered = [self.list_widget.item(i) for i in range(self.list_widget.count()) if self.list_widget.item(i).data(Qt.UserRole) in selected_paths]
            count = 0
            for idx, item in enumerate(ordered, start=1):
                img_path = item.data(Qt.UserRole)
                img = Image.open(img_path).convert('RGB')
                dst = os.path.join(folder_path, f"slide_{idx}.pdf")
                img.save(dst)
                count += 1
            QMessageBox.information(self, "成功", f"成功生成 {count} 个 PDF 文件！")
        except Exception as e:
            QMessageBox.warning(self, "错误", str(e))

    def export_all_as_pdf(self):
        # 导出全部为一个 PDF
        all_items = [self.list_widget.item(i) for i in range(self.list_widget.count())]
        self.save_as_pdf(all_items)

    def closeEvent(self, event):
        # 程序关闭时清理临时文件
        try:
            shutil.rmtree(self.temp_dir)
        except:
            pass
        event.accept()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # 设置高分屏支持
    app.setAttribute(Qt.AA_EnableHighDpiScaling)
    app.setAttribute(Qt.AA_UseHighDpiPixmaps)
    
    # 设置全局字体
    font = QFont("Microsoft YaHei", 10)
    app.setFont(font)
    # 生成并保存桌面图标（如果可能），并设置为应用图标
    try:
        desktop_dir = os.path.join(os.path.expanduser('~'), 'Desktop')
        if not os.path.isdir(desktop_dir):
            desktop_dir = os.path.expanduser('~')
        desktop_icon_path = os.path.join(desktop_dir, 'ppt_viewer_icon.ico')

        # 使用 QPixmap + QPainter 绘制一个简单的圆形带字母的图标
        pix = QPixmap(64, 64)
        pix.fill(Qt.transparent)
        painter = QPainter(pix)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setBrush(QColor('#1890ff'))
        painter.setPen(Qt.NoPen)
        painter.drawEllipse(4, 4, 56, 56)
        painter.setPen(QColor('#ffffff'))
        txt_font = QFont('Arial', 28, QFont.Bold)
        painter.setFont(txt_font)
        painter.drawText(pix.rect(), Qt.AlignCenter, 'P')
        painter.end()

        # 保存为 ICO（Qt 支持保存为 ICO 格式）
        pix.save(desktop_icon_path, 'ICO')

        # 将生成的图标设置为应用图标
        app.setWindowIcon(QIcon(desktop_icon_path))
    except Exception:
        # 生成或设置图标失败时静默继续，不影响主程序运行
        pass

    viewer = PPTViewer()
    viewer.show()
    sys.exit(app.exec_())