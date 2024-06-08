import csv
import glob
import os
import shutil
import time
import cv2
import numpy as np
import pandas as pd
import pyautogui
import markdown
from PyQt5 import uic
from PyQt5.QtCore import QRegExp, Qt
from PyQt5.QtGui import QRegExpValidator, QColor, QPixmap, QIcon
from PyQt5.QtWidgets import QMessageBox, QHeaderView, QTreeWidgetItem, QTableWidgetItem, QGraphicsScene, QHBoxLayout
from pynput import mouse, keyboard
from pynput.mouse import Controller
from skimage.metrics import structural_similarity as compare_ssim
from zoom_graphics_view import CZoomGraphicsView


class CMainWindow:
    def __init__(self):
        # 从文件中加载UI定义
        self.main_window = uic.loadUi("mainwindow.ui")
        self.instruction = uic.loadUi("instruction.ui")
        self.image_comparison = uic.loadUi("image.ui")

        # 初始化
        self.passed = 0  # 0未回放，1通过，2失败
        pyautogui.FAILSAFE = False
        self.screenshot_list = []
        self.list_index = None
        self.df = None
        self.last_time = None
        self.root_item = None
        self.child_item = None
        self.grandchild_item = None
        self.software_name = None
        self.module_name = None
        self.case_number = None
        self.case_title = None
        self.mouse_listener = None
        self.keyboard_listener = None

        # 获取当前用户的主目录路径
        self.home_dir = os.path.expanduser("~")

        # 在用户主目录下创建"测试数据"目录
        self.test_data_dir = os.path.join(self.home_dir, "测试数据")
        if not os.path.exists(self.test_data_dir):
            os.makedirs(self.test_data_dir)

        # 在"测试数据"路径下创建两个子目录
        self.record_dir = os.path.join(self.test_data_dir, "录制数据")
        if not os.path.exists(self.record_dir):
            os.makedirs(self.record_dir)
        self.playback_dir = os.path.join(self.test_data_dir, "回放数据")
        if not os.path.exists(self.playback_dir):
            os.makedirs(self.playback_dir)

        # QLineEdit输入规则
        self.main_window.softwareNameEdit.setValidator(
            QRegExpValidator(QRegExp("[\u4e00-\u9fa5A-Za-z0-9-——_. （）()]+"), self.main_window))
        self.main_window.softwareNameSearch.lineEdit().setValidator(
            QRegExpValidator(QRegExp("[\u4e00-\u9fa5A-Za-z0-9-——_. （）()]+"), self.main_window))
        self.main_window.moduleNameEdit.setValidator(
            QRegExpValidator(QRegExp("[\u4e00-\u9fa5A-Za-z0-9-——_. ]+"), self.main_window))
        self.main_window.caseNumberEdit.setValidator(
            QRegExpValidator(QRegExp("[A-Za-z0-9._]+"), self.main_window))

        # ModuleNameEdit命名规范
        with open("../resource/configure/edit_information.txt", 'r', encoding='utf-8') as f:
            self.main_window.moduleNameEdit.setToolTip(f.read())

        # RecordButton录制注意事项
        with open("../resource/configure/button_information.txt", 'r', encoding='utf-8') as f:
            self.main_window.recordBtn.setToolTip(f.read())

        # 主界面初始化
        self.main_window.setWindowIcon(QIcon("../resource/icon/client.ico"))
        self.instruction.setWindowIcon(QIcon("../resource/icon/instruction.ico"))
        self.image_comparison.setWindowIcon(QIcon("../resource/icon/image.ico"))

        # QTreeWidget初始化
        self.treeInitialization(self.main_window.tree, self.record_dir)

        # QTableWidget初始化
        self.main_window.table.horizontalHeader().setSectionsClickable(False)
        self.main_window.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

        # QCombobox初始化
        self.main_window.softwareNameSearch.lineEdit().setPlaceholderText("请输入软件名称查询")
        font = self.main_window.softwareNameSearch.lineEdit().font()
        font.setFamily("Microsoft YaHei UI")
        font.setPointSize(11)
        self.main_window.softwareNameSearch.lineEdit().setFont(font)
        self.main_window.softwareNameSearch.setCurrentIndex(-1)

        # QGraphicsView初始化
        self.image_comparison.expectScene = QGraphicsScene(self.image_comparison)
        self.image_comparison.runtimeScene = QGraphicsScene(self.image_comparison)
        self.image_comparison.expectView = CZoomGraphicsView(self.image_comparison)
        self.image_comparison.runtimeView = CZoomGraphicsView(self.image_comparison)
        self.image_comparison.imageLayout = QHBoxLayout(self.image_comparison.graphicsFrame)
        self.image_comparison.imageLayout.addWidget(self.image_comparison.expectView)
        self.image_comparison.imageLayout.addWidget(self.image_comparison.runtimeView)

        # 连接槽函数
        self.main_window.instruction.triggered.connect(self.showInstruction)
        self.main_window.recordBtn.clicked.connect(self.recordButtonClicked)
        self.main_window.recordScreenshotBtn.clicked.connect(self.screenshotButtonClicked)
        self.main_window.playbackBtn.clicked.connect(self.playbackButtonClicked)
        self.main_window.deleteBtn.clicked.connect(self.deleteButtonClicked)
        self.main_window.softwareNameSearch.currentIndexChanged.connect(self.treeSearch)
        self.main_window.softwareNameEdit.textChanged.connect(self.softwareNameUpdate)
        self.main_window.moduleNameEdit.textChanged.connect(self.moduleNameUpdate)
        self.main_window.caseNumberEdit.textChanged.connect(self.caseNumberUpdate)
        self.main_window.caseTitleEdit.textChanged.connect(self.caseTitleUpdate)
        self.main_window.tree.currentItemChanged.connect(self.treeCurrentItemUpdate)
        self.main_window.table.cellDoubleClicked.connect(self.imageComparisonShow)
        self.image_comparison.preBtn.clicked.connect(self.previousButtonClicked)
        self.image_comparison.nextBtn.clicked.connect(self.nextButtonClicked)

    def showInstruction(self):
        with open("../resource/configure/instruction.md", 'r', encoding='utf-8') as f:
            md_content = f.read()
        html_content = markdown.markdown(md_content)
        self.instruction.instructionBrowser.setHtml(html_content)
        self.instruction.show()

    def softwareNameUpdate(self, text):
        self.software_name = text

    def moduleNameUpdate(self, text):
        self.module_name = text

    def caseNumberUpdate(self, text):
        self.case_number = text

    def caseTitleUpdate(self, text):
        self.case_title = text

    def comboBoxUpDate(self, name):
        self.main_window.softwareNameSearch.addItem(name)
        items = [self.main_window.softwareNameSearch.itemText(i) for i in
                 range(self.main_window.softwareNameSearch.count())]
        items.sort()
        self.main_window.softwareNameSearch.clear()
        for item in items:
            self.main_window.softwareNameSearch.addItem(item)
        self.main_window.softwareNameSearch.setCurrentIndex(-1)

    def treeInitialization(self, parent, path):
        for name in os.listdir(path):
            full_path = os.path.join(path, name)
            if not (os.path.isdir(full_path) and os.path.basename(full_path) != "expect_image"):
                continue
            item = QTreeWidgetItem(parent, [name])
            item.setData(0, Qt.UserRole, full_path)
            if path == self.record_dir:
                self.comboBoxUpDate(name)
                item.setHidden(True)
            if os.path.isdir(full_path) and os.path.basename(full_path) != "expect_image":
                self.treeInitialization(item, full_path)

    def otherNodeHide(self, current_item):
        for i in range(self.main_window.tree.topLevelItemCount()):
            item = self.main_window.tree.topLevelItem(i)
            if current_item.text(0) == item.text(0):
                item.setHidden(False)
            else:
                item.setHidden(True)
        self.main_window.tree.expandAll()

    def treeSearch(self):
        current_text = self.main_window.softwareNameSearch.currentText()
        for i in range(self.main_window.tree.topLevelItemCount()):
            item = self.main_window.tree.topLevelItem(i)
            if item.text(0) == current_text:
                self.otherNodeHide(item)
                self.main_window.softwareNameSearch.setCurrentIndex(-1)
                break

    def levelJudge(self, level, item):
        if level == 1:
            self.grandchild_item = None
            self.child_item = None
            self.root_item = item
        elif level == 2:
            self.grandchild_item = None
            self.child_item = item
            self.root_item = self.child_item.parent()
        elif level == 3:
            self.grandchild_item = item
            self.child_item = self.grandchild_item.parent()
            self.root_item = self.child_item.parent()

    def lineEditUpdate(self):
        if self.grandchild_item:
            if self.getCsvName():
                parts = self.getCsvName().split(".")
                self.main_window.caseTitleEdit.setText(os.path.basename(parts[0]))
            self.main_window.caseNumberEdit.setText(self.grandchild_item.text(0))
        else:
            self.main_window.caseTitleEdit.setText("")
            self.main_window.caseNumberEdit.setText("")
        if self.child_item:
            self.main_window.moduleNameEdit.setText(self.child_item.text(0))
        else:
            self.main_window.moduleNameEdit.setText("")
        if self.root_item:
            self.main_window.softwareNameEdit.setText(self.root_item.text(0))
        else:
            self.main_window.softwareNameEdit.setText("")

    def treeCurrentItemUpdate(self, current_item):
        current_dir = current_item.data(0, Qt.UserRole)
        level = 0
        while current_dir != self.record_dir:
            current_dir = os.path.dirname(current_dir)
            level += 1
        self.levelJudge(level, current_item)
        self.lineEditUpdate()
        self.tableUpdate()

    @staticmethod
    def keyJudge(key):
        try:
            return key.char
        except AttributeError:
            parts = str(key).split(".")
            return parts[1]

    def onMouseMove(self, x, y):
        if time.time() - self.last_time < 0.15:
            return
        csv_list = [int((time.time() - self.last_time) * 1000),
                    "Mouse", "mouse move", [x, y], "", ""]
        self.csvUpdate(csv_list)

    def onMouseClick(self, x, y, button, pressed):
        if button != mouse.Button.left and button != mouse.Button.right and button != mouse.Button.middle:
            return
        if pressed:
            csv_list = [int((time.time() - self.last_time) * 1000),
                        "Mouse", "mouse " + button.name + " down", [x, y], "", ""]
            self.csvUpdate(csv_list)
        else:
            csv_list = [int((time.time() - self.last_time) * 1000),
                        "Mouse", "mouse " + button.name + " up", [x, y], "", ""]
            self.csvUpdate(csv_list)

    def onMouseWheel(self, x, y, dx, dy):
        if dy > 0:
            csv_list = [int((time.time() - self.last_time) * 1000),
                        "Mouse", "mouse wheel up", [x, y], "", ""]
        else:
            csv_list = [int((time.time() - self.last_time) * 1000),
                        "Mouse", "mouse wheel down", [x, y], "", ""]
        self.csvUpdate(csv_list)

    def onKeyboardDown(self, key):
        if key == keyboard.Key.esc:
            self.mouse_listener.stop()
            self.keyboard_listener.stop()
            return
        csv_list = [int((time.time() - self.last_time) * 1000),
                    "Keyboard", "key down", self.keyJudge(key), "", ""]
        self.csvUpdate(csv_list)

    def onKeyboardUp(self, key):
        csv_list = [int((time.time() - self.last_time) * 1000),
                    "Keyboard", "key up", self.keyJudge(key), "", ""]
        self.csvUpdate(csv_list)

    def monitor(self):
        self.mouse_listener = mouse.Listener(
            on_move=self.onMouseMove, on_click=self.onMouseClick, on_scroll=self.onMouseWheel)
        self.keyboard_listener = keyboard.Listener(on_press=self.onKeyboardDown, on_release=self.onKeyboardUp)
        self.mouse_listener.start()
        self.keyboard_listener.start()
        self.last_time = time.time()
        self.mouse_listener.join()
        self.keyboard_listener.join()
        self.tableUpdate()

    def csvUpdate(self, csv_list):
        csv_path = os.path.join(self.grandchild_item.data(0, Qt.UserRole), self.case_title + ".csv")
        with open(csv_path, mode='a', encoding='utf-8', newline='') as csv_file:
            csv.writer(csv_file).writerow(csv_list)
        self.last_time = time.time()

    def dirJudge(self):
        temp_dir = self.grandchild_item.data(0, Qt.UserRole).replace(self.record_dir, self.playback_dir)
        if not os.path.exists(
                self.grandchild_item.data(0, Qt.UserRole).replace(self.record_dir, self.playback_dir)):
            temp_dir = self.grandchild_item.data(0, Qt.UserRole)
        return temp_dir

    def getCsvName(self):
        csv_dir = self.dirJudge()
        files = glob.glob(f"{csv_dir}/*.csv")
        if files:
            return files[0]

    def clickStepColor(self, row):
        step_state = str(self.df.iloc[row, 5])
        for col in range(self.df.shape[1]):
            item = QTableWidgetItem(str(self.df.iloc[row, col]))
            self.main_window.table.setItem(row, col, item)
            if step_state == "nan":
                item.setBackground(QColor(0, 190, 225))
            elif step_state == "通过":
                if self.passed == 0:
                    self.passed = 1
                item.setBackground(QColor(0, 250, 26))
            elif step_state == "失败":
                self.passed = 2
                item.setBackground(QColor(255, 39, 0))

    def createTableItem(self):
        self.main_window.table.setRowCount(self.df.shape[0])
        self.main_window.table.setColumnCount(self.df.shape[1])
        self.main_window.table.setHorizontalHeaderLabels(self.df.columns)
        for row in range(self.df.shape[0]):
            if row < self.df.shape[0] - 1:
                if (str(self.df.iloc[row, 2]) == "mouse left up"
                        and ((str(self.df.iloc[row + 1, 2]) != "mouse left down"
                              and str(self.df.iloc[row + 1, 2]) != "mouse right down")
                             or self.df.iloc[row + 1, 0] > 200)):
                    self.screenshot_list.append(row + 1)
                    self.clickStepColor(row)
                elif (str(self.df.iloc[row, 2]) == "mouse right up"
                      and ((str(self.df.iloc[row + 1, 2]) != "mouse left down"
                            and str(self.df.iloc[row + 1, 2]) != "mouse right down")
                           or self.df.iloc[row + 1, 0] > 200)):
                    self.screenshot_list.append(row + 1)
                    self.clickStepColor(row)
                else:
                    for col in range(self.df.shape[1]):
                        item = QTableWidgetItem(str(self.df.iloc[row, col]))
                        self.main_window.table.setItem(row, col, item)
            else:
                if (str(self.df.iloc[row, 2]) == "mouse left up"
                        or str(self.df.iloc[row, 2]) == "mouse right up"):
                    self.screenshot_list.append(row + 1)
                    self.clickStepColor(row)
                else:
                    for col in range(self.df.shape[1]):
                        item = QTableWidgetItem(str(self.df.iloc[row, col]))
                        self.main_window.table.setItem(row, col, item)

    def statusBarUpdate(self):
        if self.df.shape[0] == 0:
            self.main_window.statusBar.showMessage("当前用例状态：无录制数据！")
        elif self.passed == 0:
            self.main_window.statusBar.showMessage("当前用例状态：未回放！")
        elif self.passed == 1:
            self.main_window.statusBar.showMessage("当前用例状态：回放通过！")
        elif self.passed == 2:
            self.main_window.statusBar.showMessage("当前用例状态：回放失败！")

    def tableUpdate(self):
        if not self.grandchild_item:
            self.main_window.table.setRowCount(0)
            self.main_window.table.setColumnCount(0)
            self.main_window.statusBar.showMessage("")
            return
        if not self.getCsvName():
            self.main_window.statusBar.showMessage("")
            return
        self.passed = 0
        self.screenshot_list.clear()
        self.df = pd.read_csv(self.getCsvName(), encoding='utf8')
        self.createTableItem()
        self.statusBarUpdate()

    def recordFileCatalogUpdate(self):
        # root目录
        root_dir = os.path.join(self.record_dir, self.software_name)
        if not os.path.exists(root_dir):
            os.makedirs(root_dir)
            self.root_item = QTreeWidgetItem(self.main_window.tree, [self.software_name])
            self.root_item.setData(0, Qt.UserRole, root_dir)
            self.comboBoxUpDate(self.software_name)
        else:
            for i in range(self.main_window.tree.topLevelItemCount()):
                item = self.main_window.tree.topLevelItem(i)
                if item.text(0) == self.software_name:
                    self.root_item = item
                    break
        self.otherNodeHide(self.root_item)
        # child目录
        child_dir = os.path.join(root_dir, self.module_name)
        if not os.path.exists(child_dir):
            os.makedirs(child_dir)
            self.child_item = QTreeWidgetItem(self.root_item, [self.module_name])
            self.child_item.setData(0, Qt.UserRole, child_dir)
        else:
            for i in range(self.root_item.childCount()):
                item = self.root_item.child(i)
                if item.text(0) == self.module_name:
                    self.child_item = item
                    break
        # grandchild目录
        grandchild_dir = os.path.join(child_dir, self.case_number)
        if not os.path.exists(grandchild_dir):
            os.makedirs(grandchild_dir)
            self.grandchild_item = QTreeWidgetItem(self.child_item, [self.case_number])
            self.grandchild_item.setData(0, Qt.UserRole, grandchild_dir)
        else:
            for i in range(self.child_item.childCount()):
                item = self.child_item.child(i)
                if item.text(0) == self.case_number:
                    self.grandchild_item = item
                    break
        self.main_window.tree.setCurrentItem(self.grandchild_item)

    def deleteCase(self):
        if os.path.exists(self.grandchild_item.data(0, Qt.UserRole)):
            shutil.rmtree(self.grandchild_item.data(0, Qt.UserRole))
        playback_data_dir = self.grandchild_item.data(0, Qt.UserRole).replace(self.record_dir, self.playback_dir)
        if os.path.exists(playback_data_dir):
            shutil.rmtree(playback_data_dir)
        self.main_window.table.setRowCount(0)
        self.main_window.table.setColumnCount(0)

    def record(self):
        header = ["操作间隔", "按键类别", "事件类型", "输入数据", "相似度", "结果"]
        self.csvUpdate(header)
        self.main_window.showMinimized()
        self.monitor()
        self.main_window.showNormal()
        QMessageBox.information(self.main_window, "提示", "录制结束，若显示数据为空请重新录制！")

    def recordButtonRunnableJudge(self):
        if not self.software_name:
            QMessageBox.critical(self.main_window, "错误", "项目名不能为空！")
        elif not self.module_name:
            QMessageBox.critical(self.main_window, "错误", "测试模块名不能为空！")
        elif not self.case_number:
            QMessageBox.critical(self.main_window, "错误", "用例编号不能为空！")
        elif not self.case_title:
            QMessageBox.critical(self.main_window, "错误", "用例标题不能为空！")
        else:
            self.recordFileCatalogUpdate()
            if os.listdir(self.grandchild_item.data(0, Qt.UserRole)):
                warning_box = QMessageBox.warning(
                    self.main_window, "警告", "已存在录制数据，是否删除当前录制数据？（继续录制请重新点击开始录制按钮）",
                    QMessageBox.Yes | QMessageBox.No)
                if warning_box == QMessageBox.Yes:
                    self.deleteCase()
                    os.makedirs(self.grandchild_item.data(0, Qt.UserRole))
                return
            self.record()

    def recordButtonClicked(self):
        self.recordButtonRunnableJudge()

    def Screenshot(self):
        playback_data_dir = self.grandchild_item.data(0, Qt.UserRole).replace(self.record_dir, self.playback_dir)
        if os.path.exists(playback_data_dir):
            shutil.rmtree(playback_data_dir)
        is_screenshot = False
        m = Controller()
        index = 0
        parts = []
        expect_image_dir = os.path.join(self.grandchild_item.data(0, Qt.UserRole), "expect_image")
        if not os.path.exists(expect_image_dir):
            os.makedirs(expect_image_dir)
        else:
            shutil.rmtree(expect_image_dir)
            os.makedirs(expect_image_dir)
        for row in range(self.df.shape[0]):
            lapse_time = self.df.iloc[row, 0] * 0.001 + 0.01
            event_type = str(self.df.iloc[row, 2])
            input_data = str(self.df.iloc[row, 3])
            parts.clear()
            input_data = input_data.replace("[", "")
            input_data = input_data.replace("]", "")
            parts = input_data.split(", ")
            if event_type == "mouse move":
                time.sleep(lapse_time)
                pyautogui.moveTo(int(parts[0]), int(parts[1]))
            elif event_type == "mouse left down":
                time.sleep(lapse_time)
                pyautogui.moveTo(int(parts[0]), int(parts[1]))
                pyautogui.mouseDown()
            elif event_type == "mouse left up":
                time.sleep(lapse_time)
                pyautogui.mouseUp()
                if row != self.screenshot_list[index] - 1:
                    continue
                index += 1
                time.sleep(2)
                is_screenshot = True
                pyautogui.screenshot().save(os.path.join(expect_image_dir, str(row + 1) + ".png"))
            elif event_type == "mouse right down":
                time.sleep(lapse_time)
                pyautogui.moveTo(int(parts[0]), int(parts[1]))
                pyautogui.mouseDown(button="right")
            elif event_type == "mouse right up":
                time.sleep(lapse_time)
                pyautogui.mouseUp(button="right")
                if row != self.screenshot_list[index] - 1:
                    continue
                index += 1
                time.sleep(2)
                is_screenshot = True
                pyautogui.screenshot().save(os.path.join(expect_image_dir, str(row + 1) + ".png"))
            elif event_type == "mouse middle down":
                time.sleep(lapse_time)
                pyautogui.moveTo(int(parts[0]), int(parts[1]))
                pyautogui.mouseDown(button="middle")
            elif event_type == "mouse middle up":
                time.sleep(lapse_time)
                pyautogui.mouseUp(button="middle")
            elif event_type == "mouse wheel down":
                time.sleep(lapse_time)
                pyautogui.moveTo(int(parts[0]), int(parts[1]))
                m.scroll(0, -1)
            elif event_type == "mouse wheel up":
                time.sleep(lapse_time)
                pyautogui.moveTo(int(parts[0]), int(parts[1]))
                m.scroll(0, 1)
            elif event_type == "key down":
                time.sleep(lapse_time)
                pyautogui.keyDown(parts[0])
            elif event_type == "key up":
                time.sleep(lapse_time)
                pyautogui.keyUp(parts[0])
            elif event_type == "nan":
                time.sleep(lapse_time)
        self.main_window.showNormal()
        if not is_screenshot:
            QMessageBox.critical(self.main_window, "错误", "无有效鼠标输入！")
            shutil.rmtree(expect_image_dir)
        else:
            QMessageBox.information(self.main_window, "提示", "录制截图正常！")

    def screenshotButtonRunnableJudge(self):
        if not self.software_name:
            QMessageBox.critical(self.main_window, "错误", "项目名不能为空！")
        elif not self.module_name:
            QMessageBox.critical(self.main_window, "错误", "测试模块名不能为空！")
        elif not self.case_number:
            QMessageBox.critical(self.main_window, "错误", "用例编号不能为空！")
        elif not self.case_title:
            QMessageBox.critical(self.main_window, "错误", "用例标题不能为空！")
        else:
            if not self.grandchild_item:
                self.recordFileCatalogUpdate()
            if not self.getCsvName():
                QMessageBox.critical(self.main_window, "错误", "项目录制数据不存在！")
                return
            self.main_window.showMinimized()
            self.Screenshot()

    def screenshotButtonClicked(self):
        self.screenshotButtonRunnableJudge()

    def playbackFileCatalogUpdate(self):
        playback_data_dir = self.grandchild_item.data(0, Qt.UserRole).replace(self.record_dir, self.playback_dir)
        if not os.path.exists(playback_data_dir):
            os.makedirs(playback_data_dir)
        else:
            shutil.rmtree(playback_data_dir)
            os.makedirs(playback_data_dir)

    def showImage(self):
        step_name = self.screenshot_list[self.list_index]
        self.image_comparison.setWindowTitle(
            "查看第" + str(step_name) + "步图片，结果：" + str(self.df.iloc[step_name - 1, 5]))
        record_image_dir = os.path.join(self.grandchild_item.data(0, Qt.UserRole), "expect_image")
        playback_image_dir = os.path.join(
            self.grandchild_item.data(0, Qt.UserRole).replace(self.record_dir, self.playback_dir), "runtime_image")
        record_image_path = os.path.join(record_image_dir, str(self.screenshot_list[self.list_index]) + ".png")
        playback_image_path = os.path.join(playback_image_dir, str(self.screenshot_list[self.list_index]) + ".png")
        self.image_comparison.expectScene.addPixmap(QPixmap(record_image_path))
        self.image_comparison.runtimeScene.addPixmap(QPixmap(playback_image_path))
        self.image_comparison.expectView.setScene(self.image_comparison.expectScene)
        self.image_comparison.runtimeView.setScene(self.image_comparison.runtimeScene)

    def previousButtonClicked(self):
        if self.list_index == 0:
            self.image_comparison.preBtn.setEnabled(False)
            return
        self.list_index -= 1
        self.image_comparison.nextBtn.setEnabled(True)
        self.showImage()

    def nextButtonClicked(self):
        if self.list_index == len(self.screenshot_list) - 1:
            self.image_comparison.nextBtn.setEnabled(False)
            return
        self.list_index += 1
        self.image_comparison.preBtn.setEnabled(True)
        self.showImage()

    def imageComparisonShow(self, row):
        if str(self.df.iloc[row, 5]) == "nan":
            return
        self.list_index = self.screenshot_list.index(row + 1)
        self.image_comparison.preBtn.setEnabled(True)
        self.image_comparison.nextBtn.setEnabled(True)
        if self.list_index == 0:
            self.image_comparison.preBtn.setEnabled(False)
        if self.list_index == len(self.screenshot_list) - 1:
            self.image_comparison.nextBtn.setEnabled(False)
        self.showImage()
        self.image_comparison.show()

    def imageComparison(self, row, playback_image_dir, record_image_dir):
        playback_image_path = os.path.join(playback_image_dir, str(row + 1) + ".png")
        record_image_path = os.path.join(record_image_dir, str(row + 1) + ".png")
        playback_image = cv2.imdecode(np.fromfile(playback_image_path, dtype=np.uint8), cv2.IMREAD_UNCHANGED)
        record_image = cv2.imdecode(np.fromfile(record_image_path, dtype=np.uint8), cv2.IMREAD_UNCHANGED)
        playback_gray = cv2.cvtColor(playback_image, cv2.COLOR_BGR2GRAY)
        record_gray = cv2.cvtColor(record_image, cv2.COLOR_BGR2GRAY)
        ssim_value = compare_ssim(playback_gray, record_gray, multichannel=True)
        self.df.iloc[row, 4] = round(ssim_value, 4)
        if ssim_value > 0.985:
            self.df.iloc[row, 5] = "通过"
        else:
            self.df.iloc[row, 5] = "失败"

    def playback(self):
        m = Controller()
        index = 0
        parts = []
        playback_data_dir = self.grandchild_item.data(0, Qt.UserRole).replace(self.record_dir, self.playback_dir)
        runtime_image_dir = os.path.join(playback_data_dir, "runtime_image")
        if not os.path.exists(runtime_image_dir):
            os.makedirs(runtime_image_dir)
        else:
            shutil.rmtree(runtime_image_dir)
            os.makedirs(runtime_image_dir)
        for row in range(self.df.shape[0]):
            lapse_time = self.df.iloc[row, 0] * 0.001 + 0.01
            event_type = str(self.df.iloc[row, 2])
            input_data = str(self.df.iloc[row, 3])
            parts.clear()
            input_data = input_data.replace("[", "")
            input_data = input_data.replace("]", "")
            parts = input_data.split(", ")
            if event_type == "mouse move":
                time.sleep(lapse_time)
                pyautogui.moveTo(int(parts[0]), int(parts[1]))
            elif event_type == "mouse left down":
                time.sleep(lapse_time)
                pyautogui.moveTo(int(parts[0]), int(parts[1]))
                pyautogui.mouseDown()
            elif event_type == "mouse left up":
                time.sleep(lapse_time)
                pyautogui.mouseUp()
                if row != self.screenshot_list[index] - 1:
                    continue
                index += 1
                time.sleep(2)
                pyautogui.screenshot().save(os.path.join(runtime_image_dir, str(row + 1) + ".png"))
                self.imageComparison(row, runtime_image_dir,
                                     os.path.join(self.grandchild_item.data(0, Qt.UserRole), "expect_image"))
            elif event_type == "mouse right down":
                time.sleep(lapse_time)
                pyautogui.moveTo(int(parts[0]), int(parts[1]))
                pyautogui.mouseDown(button="right")
            elif event_type == "mouse right up":
                time.sleep(lapse_time)
                pyautogui.mouseUp(button="right")
                if row != self.screenshot_list[index] - 1:
                    continue
                index += 1
                time.sleep(2)
                pyautogui.screenshot().save(os.path.join(runtime_image_dir, str(row + 1) + ".png"))
                self.imageComparison(
                    row, runtime_image_dir, os.path.join(self.grandchild_item.data(0, Qt.UserRole), "expect_image"))
            elif event_type == "mouse middle down":
                time.sleep(lapse_time)
                pyautogui.moveTo(int(parts[0]), int(parts[1]))
                pyautogui.mouseDown(button="middle")
            elif event_type == "mouse middle up":
                time.sleep(lapse_time)
                pyautogui.mouseUp(button="middle")
            elif event_type == "mouse wheel down":
                time.sleep(lapse_time)
                pyautogui.moveTo(int(parts[0]), int(parts[1]))
                m.scroll(0, -1)
            elif event_type == "mouse wheel up":
                time.sleep(lapse_time)
                pyautogui.moveTo(int(parts[0]), int(parts[1]))
                m.scroll(0, 1)
            elif event_type == "key down":
                time.sleep(lapse_time)
                pyautogui.keyDown(parts[0])
            elif event_type == "key up":
                time.sleep(lapse_time)
                pyautogui.keyUp(parts[0])
            elif event_type == "nan":
                time.sleep(lapse_time)
        self.main_window.showNormal()
        QMessageBox.information(self.main_window, "提示", "回放结束！")
        self.df.to_csv(os.path.join(playback_data_dir, self.case_title + ".csv"), index=False, encoding='utf-8')

    def playbackButtonRunnableJudge(self):
        if not self.software_name:
            QMessageBox.critical(self.main_window, "错误", "项目名不能为空！")
        elif not self.module_name:
            QMessageBox.critical(self.main_window, "错误", "测试模块名不能为空！")
        elif not self.case_number:
            QMessageBox.critical(self.main_window, "错误", "用例编号不能为空！")
        elif not self.case_title:
            QMessageBox.critical(self.main_window, "错误", "用例标题不能为空！")
        else:
            if not self.grandchild_item:
                self.recordFileCatalogUpdate()
            if not os.path.exists(os.path.join(self.grandchild_item.data(0, Qt.UserRole), "expect_image")):
                QMessageBox.critical(self.main_window, "错误", "开始回放前需要执行当前用例录制截图操作！")
                return
            self.playbackFileCatalogUpdate()
            self.main_window.showMinimized()
            self.playback()
            self.tableUpdate()

    def playbackButtonClicked(self):
        self.playbackButtonRunnableJudge()

    def deleteButtonRunnableJudge(self):
        if not self.software_name:
            QMessageBox.critical(self.main_window, "错误", "项目名不能为空！")
        elif not self.module_name:
            QMessageBox.critical(self.main_window, "错误", "测试模块名不能为空！")
        elif not self.case_number:
            QMessageBox.critical(self.main_window, "错误", "用例编号不能为空！")
        elif not self.case_title:
            QMessageBox.critical(self.main_window, "错误", "用例标题不能为空！")
        else:
            if not self.grandchild_item:
                self.recordFileCatalogUpdate()
            warning_box = QMessageBox.warning(
                self.main_window, "警告", "是否删除当前用例？", QMessageBox.Yes | QMessageBox.No)
            if warning_box == QMessageBox.Yes:
                self.deleteCase()
                self.child_item.removeChild(self.grandchild_item)
                self.grandchild_item = None
                QMessageBox.information(self.main_window, "提示", "测试用例已删除！")

    def deleteButtonClicked(self):
        self.deleteButtonRunnableJudge()
