import threading
# 从PyQt库导入QtWidget通用窗口类,基本的窗口集在PyQt5.QtWidgets模块里.
from PyQt5.QtWidgets import QWidget, QMessageBox, QApplication, QAction, QSystemTrayIcon, QMenu, QDialog
from PyQt5.QtGui import QStandardItem, QIcon, QStandardItemModel
from PyQt5.QtCore import QCoreApplication, Qt
from mainwindow import *
from path_trans import pp
from reporter import *
from progress_dial import Ui_Dialog


class DownloadDialog(QDialog, Ui_Dialog):

    def __init__(self, auto_close=True, parent=None):
        """
        Constructor

        @download_url:下载地址
        @auto_close:下载完成后时候是否需要自动关闭
        """
        super(DownloadDialog, self).__init__(parent)
        self.setupUi(self)
        self.setFixedSize(self.size())
        self.progressBar.setValue(0)
        self.auto_close = auto_close

    def change_progressbar_value(self, value, spd):
        self.setWindowTitle('程序更新 - ' + spd)
        self.progressBar.setValue(value)
        if self.auto_close and value == 100:
            self.close()


class MainForm(QWidget, Ui_MainForm):

    def __init__(self):
        super(MainForm, self).__init__()
        self.setupUi(self)
        self.setWindowFlag(Qt.WindowType.MSWindowsFixedSizeDialogHint)
        self.progress_dial = DownloadDialog(parent=self)
        self.wlanuser.setEnabled(False)
        self.wlanpass.setEnabled(False)
        self.zjuwlan_cbox.setChecked(False)
        self.manager = Reporter(self)
        if self.manager.first_use:
            self.about()
        self.manager.prompt = self.prompt
        self.model = QStandardItemModel(len(self.manager.profile["profiles"].keys()), 2)
        self.model.setHorizontalHeaderLabels(["统一登录账号", "密码"])
        count = 0
        for key in self.manager.profile["profiles"]:
            self.model.setItem(count, 0, QStandardItem(key))
            self.model.setItem(count, 1, QStandardItem(self.manager.profile["profiles"][key]))
            count += 1
        self.tableView.setModel(self.model)
        self.hourbox.setValue(self.manager.profile["time"][0])
        self.minbox.setValue(self.manager.profile["time"][1])
        self.secbox.setValue(self.manager.profile["time"][2])
        if self.manager.profile["wlaninfo"][0] != "":
            self.zjuwlan_cbox.setChecked(True)
            self.wlanuser.setEnabled(True)
            self.wlanpass.setEnabled(True)
            self.wlanuser.setText(self.manager.profile["wlaninfo"][0])
            self.wlanpass.setText(self.manager.profile["wlaninfo"][1])

        self.add_btn.clicked.connect(self.add)
        self.del_btn.clicked.connect(self.delete)
        self.model.dataChanged.connect(self.data_changed)
        self.wlanuser.editingFinished.connect(self.data_changed)
        self.wlanpass.editingFinished.connect(self.data_changed)
        self.zjuwlan_cbox.clicked.connect(self.data_changed)
        self.hourbox.valueChanged.connect(self.data_changed)
        self.minbox.valueChanged.connect(self.data_changed)
        self.secbox.valueChanged.connect(self.data_changed)
        self.manager.download_start_sig.connect(self.show_download_ui)
        self.manager.download_process_sig.connect(self.progress_dial.change_progressbar_value)

        clock_thread = threading.Thread(name="check_time", target=self.manager.check_clocking)
        status_thread = threading.Thread(name="check_status", target=self.manager.check_status)
        recon_thread = threading.Thread(name="wifi_recon", target=self.manager.internet_reconnect)
        update_thread = threading.Thread(name="check_upgrade", target=self.manager.check_upgrade)

        recon_thread.start()
        clock_thread.start()
        status_thread.start()
        update_thread.start()

    def data_changed(self):
        temp = {}
        for i in range(self.model.rowCount()):
            temp[self.model.item(i,0).text()] = self.model.item(i, 1).text()
        self.manager.profile["profiles"] = temp

        if self.zjuwlan_cbox.isChecked():
            self.wlanpass.setEnabled(True)
            self.wlanuser.setEnabled(True)
            self.wlanuser.clearFocus()
            self.wlanpass.clearFocus()
            self.manager.profile["wlaninfo"][0] = self.wlanuser.text()
            self.manager.profile["wlaninfo"][1] = self.wlanpass.text()
        else:
            self.wlanuser.setEnabled(False)
            self.wlanpass.setEnabled(False)
            self.manager.profile["wlaninfo"] = ['', '']

        self.manager.profile["time"] = [self.hourbox.value(), self.minbox.value(), self.secbox.value()]
        self.manager.save_data()

        print(self.manager.profile)

    def add(self):
        self.model.appendRow([QStandardItem("请输入"), QStandardItem("请输入")])

    def delete(self):
        if self.model.rowCount() > 1:
            indexes = self.tableView.selectedIndexes()
            if len(indexes) > 0:
                self.model.removeRows(indexes[0].row(), 1)
                self.data_changed()

    def first_run(self):
        re = QMessageBox.question(self, "免责声明",  "本程序仅方便长期在校在杭同学的日常打卡，作者对包括但\n不限于利用该程序作为离杭出省的虚拟打卡用途负任何责任！\n您是否接受该声明？", QMessageBox.Yes |
                                   QMessageBox.No)
        if re == QMessageBox.No:
            # 关闭窗体程序
            os._exit(0)
        else:
            log("用户已经接受免责声明")

    def show_download_ui(self):
        self.progress_dial.show()
        download_thread = threading.Thread(name="down", target=self.manager.upgrade)
        download_thread.start()

    def multi_instant_warning(self):
        re = QMessageBox.critical(self, "警告", "已有程序在运行", QMessageBox.Yes )
        if re == QMessageBox.Yes:
            os._exit(0)


if __name__ == '__main__':
    # pyqt窗口必须在QApplication方法中使用 
    # 每一个PyQt5应用都必须创建一个应用对象.sys.argv参数是来自命令行的参数列表.Python脚本可以从shell里运行.这是我们如何控制我们的脚本运行的一种方法.
    app = QApplication(sys.argv)

    # 关闭所有窗口,也不关闭应用程序
    QApplication.setQuitOnLastWindowClosed(False)
    from PyQt5 import QtWidgets

    # QWidget窗口是PyQt5中所有用户界口对象的基本类.我们使用了QWidget默认的构造器.默认的构造器没有父类.一个没有父类的窗口被称为一个window.
    w = MainForm()
    # resize()方法调整了窗口的大小.被调整为250像素宽和250像素高.
    #w.resize(250, 250)
    # move()方法移动了窗口到屏幕坐标x=300, y=300的位置.
    w.move(300, 300)
    # 在这里我们设置了窗口的标题.标题会被显示在标题栏上.
    w.setWindowTitle('每日打卡 v.' + w.manager.parser.GetFileVersion(sys.argv[0]))
    w.setWindowIcon(QIcon(pp(r'.\img\icon.ico')))
    # show()方法将窗口显示在屏幕上.一个窗口是先在内存中被创建,然后显示在屏幕上的.
    w.show()

    # from PyQt5.QtWidgets import QSystemTrayIcon
    # from PyQt5.QtGui import QIcon
    # 在系统托盘处显示图标
    tp = QSystemTrayIcon(w)
    tp.setIcon(QIcon(pp(r'.\img\icon.ico')))
    # 设置系统托盘图标的菜单
    a1 = QAction('&显示(Show)', triggered=w.show)


    def quit_app():
        w.show()  # w.hide() #隐藏
        re = QMessageBox.question(w, "提示", "退出打卡", QMessageBox.Yes |
                                  QMessageBox.No, QMessageBox.No)
        if re == QMessageBox.Yes:
            # 关闭窗体程序
            QCoreApplication.instance().quit()
            w.manager.profile['exe_pid'] = ''
            w.manager.save_data()
            # 在应用程序全部关闭后，TrayIcon其实还不会自动消失，
            # 直到你的鼠标移动到上面去后，才会消失，
            # 这是个问题，（如同你terminate一些带TrayIcon的应用程序时出现的状况），
            # 这种问题的解决我是通过在程序退出前将其setVisible(False)来完成的。 
            tp.setVisible(False)


    a2 = QAction('&退出(Exit)', triggered=quit_app)  # 直接退出可以用qApp.quit

    tpMenu = QMenu()
    tpMenu.addAction(a1)
    tpMenu.addAction(a2)
    tp.setContextMenu(tpMenu)
    # 不调用show不会显示系统托盘
    tp.show()

    # 信息提示
    # 参数1：标题
    # 参数2：内容
    # 参数3：图标（0没有图标 1信息图标 2警告图标 3错误图标），0还是有一个小图标
    #tp.showMessage('每日打卡', '自动打卡已开始工作', icon=0)


    def message():
        print("弹出的信息被点击了")


    tp.messageClicked.connect(message)


    def act(reason):
        # 鼠标点击icon传递的信号会带有一个整形的值，1是表示单击右键，2是双击，3是单击左键，4是用鼠标中键点击
        if reason == 2 or reason == 3:
            w.show()
        # print("系统托盘的图标被点击了")


    tp.activated.connect(act)

    # sys为了调用sys.exit(0)退出程序
    # 最后,我们进入应用的主循环.事件处理从这里开始.主循环从窗口系统接收事件,分派它们到应用窗口.如果我们调用了exit()方法或者主窗口被销毁,则主循环结束.sys.exit()方法确保一个完整的退出.环境变量会被通知应用是如何结束的.
    # exec_()方法是有一个下划线的.这是因为exec在Python中是关键字.因此,用exec_()代替.
    #sys.exit(app.exec_())
    os._exit(app.exec_())

