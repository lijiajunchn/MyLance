from PyQt5.QtWidgets import QApplication
from mainwindow import CMainWindow


if __name__ == '__main__':
    app = QApplication([])
    w = CMainWindow()
    w.main_window.show()
    app.exec_()
