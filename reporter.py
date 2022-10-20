import os, shutil, requests, pywifi, sys, win32process, win32event, psutil, glob
from urllib import request
from bs4 import BeautifulSoup
from time import sleep
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from win32com.client import Dispatch
from zipfile import ZipFile
from datetime import datetime
from pywifi import const
#from ddddocr import DdddOcr
from ftplib import FTP
from PyQt5.QtCore import pyqtSignal, QObject


class Reporter(QObject):

    profile_path = os.getenv('TEMP')[:len(os.getenv('TEMP')) - 4] + "zju_ahr.txt"
    driver_path = os.getenv('TEMP')[:len(os.getenv('TEMP')) - 4] + "msedgedriver.exe"
    edge_path = 'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe'
    operation_url = 'https://maxzyx.lofter.com/post/7510259a_2b5f60df5'
    upgrade_url = 'https://maxzyx.lofter.com/post/7510259a_2b70769ef'
    driver_backup_path = os.getenv('TEMP')[:len(os.getenv('TEMP')) - 18] + 'Documents\\'
    ocr_img_path = os.getenv('TEMP') + '\\img.png'
    current_date = datetime.today().date()
    prompt = ""
    ready = True
    first_use = False
    parser = Dispatch("Scripting.FileSystemObject")
    download_process_sig = pyqtSignal(int, str)
    download_start_sig = pyqtSignal()
    app_version = parser.GetFileVersion(sys.argv[0])
    exe_path = str(sys.argv[0]).replace('\\', '/')

    def __init__(self, window):
        super(Reporter, self).__init__()
        print("身份信息保存目录：" + self.profile_path)
        backup_pyfiles()

        # 获取云端的最新版本号和下载链接
        self.window = window
        self.is_upgrading = False

        if os.path.exists(os.getcwd() + "/upgrade_helper.exe"):
            os.remove(os.getcwd() + "/upgrade_helper.exe")

        with open(self.profile_path, 'r') as f:
            self.profile = eval(f.read())

        if 'profiles' not in self.profile.keys():
            window.first_run()
            default_data = {"profiles": {"请输入": "请输入"}, "wlaninfo": ['', ''], "time": [0, 10, 0], 'exe_pid': ''}
            with open(self.profile_path, 'w') as f:
                f.write(str(default_data))

        with open(self.profile_path, 'r') as f:
            self.profile = eval(f.read())
        self.profile['exe_path'] = self.exe_path
        self.save_data()

        pids = psutil.pids()
        is_running = False
        self_pid = ''

        for pid in pids:
            p = psutil.Process(pid)
            if self.profile['exe_pid'] == pid:
                is_running = True
            else:
                if p.name() == os.path.basename(sys.argv[0]):
                    self_pid = pid
        if is_running:
            log("重复运行")
            window.multi_instant_warning()
        else:
            log('当前程序PID为：' + str(self_pid))
            self.profile['exe_pid'] = self_pid
            self.save_data()

    @staticmethod
    def move_file(srcfile, dstpath):
        if not os.path.isfile(srcfile):
            print("请完整的解压文件夹后运行本程序")
        else:
            fpath, fname = os.path.split(srcfile)
            if not os.path.exists(dstpath):
                os.makedirs(dstpath)
            shutil.move(srcfile, dstpath+fname)

    @staticmethod
    def internet_connection():
        try:
            html = requests.get("https://maxzyx.lofter.com/post/7510259a_2b5f60df5", timeout=1)
        except:
            return False
        return True

    @staticmethod
    def list_from_lofter(url):
        html = request.urlopen(url).read()
        soup = BeautifulSoup(html, 'html.parser', from_encoding='utf-8')
        resdata = soup.findAll('p')
        linedata = []
        for i in resdata:
            ix = str(i)[3:-4]
            ix = ix.replace("<br/>", "")
            linedata.append(ix)
        return linedata

    def save_data(self):
        with open(self.profile_path, 'w') as f:
            f.write(str(self.profile))

    def upgrade(self):


        '''
        ftp = FTP()
        ftp.set_debuglevel(1)
        ftp.connect(host="zjuahp.6te.net", port=21)
        ftp.login("zjuahp.6te.net", "zyx12040")
        file_list = ftp.nlst()
        print(file_list)
        zipfile_path = os.getenv('TEMP')[:len(os.getenv('TEMP')) - 4] + "upgrade.zip"
        ftp.cwd("/AHR")
        file_size = ftp.size("upgrade.zip")
        with open(zipfile_path, 'wb') as f:
            def write(data):
                f.write(data)
                pct = int(100*os.stat(zipfile_path).st_size / file_size)
                self.download_process_sig.emit(pct)

            ftp.retrbinary('RETR %s' % os.path.basename("upgrade.zip"), write, blocksize=1024)
'''
        temp = self.list_from_lofter(self.upgrade_url)
        latest_app_url = temp[1]

        zipfile_path = os.getenv('TEMP')[:len(os.getenv('TEMP')) - 4] + "upgrade.zip"
        file_size = requests.get(latest_app_url, stream=True).headers['Content-Length']

        with open(zipfile_path, 'wb') as f:
            db = requests.get(latest_app_url, stream=True)
            offset = 0
            last_time = datetime.now()
            last_offset = 0
            speed_str = ''
            for chunk in db.iter_content(chunk_size=1024):
                if not chunk:
                    break
                f.seek(offset)
                f.write(chunk)
                offset = offset + len(chunk)
                process = offset / int(file_size) * 100
                dt_time = datetime.now() - last_time
                if dt_time.seconds == 1:
                    speed = (offset - last_offset)/1024/1
                    if speed < 1000:
                        speed_str = str(round(speed, 1)) + 'KB/s'
                    else:
                        speed_str = str(round(speed/1024, 1)) + 'MB/s'
                    last_time = datetime.now()
                    last_offset = offset

                self.download_process_sig.emit(process, speed_str)
        fz = ZipFile(zipfile_path, 'r')
        for files in fz.namelist():
            fz.extract(files, os.getcwd())

        helper_exe = os.getcwd() + "/upgrade_helper.exe"
        handle = win32process.CreateProcess(helper_exe, '', None, None, 0, win32process.CREATE_NO_WINDOW, None, None,
                                            win32process.STARTUPINFO())
        win32event.WaitForSingleObject(handle[0], 0)
        os._exit(0)

    def check_upgrade(self):
        while self.app_version != '' and not self.is_upgrading:
            temp = self.list_from_lofter(self.upgrade_url)
            latest_app_version = temp[0]
            if latest_app_version != self.app_version:
                log("软件版本：" + self.app_version)
                log("新版本：" + latest_app_version)
                log("启动更新进程")
                self.is_upgrading = True
                self.download_start_sig.emit()
            sleep(30)


    def download_driver(self):
        log("下载驱动中")
        try:
            response = requests.get('https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/')
            response.encoding = "utf-8"
            html = response.text

            soup = BeautifulSoup(html)
            allnode_of_a = soup.findAll("a")
            result = [_.get("href") for _ in allnode_of_a]
            driver_url = result[3]
            driver_header = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; WOW64; rv:68.0) Gecko/20100101 Firefox/68.0"}
            r = requests.get(url=driver_url, headers=driver_header)
            driver_zipfile_path = os.getenv('TEMP') + "\msdriver.zip"
            driver_unzip_path = os.getenv('TEMP')[:len(os.getenv('TEMP')) - 4]
            with open(driver_zipfile_path, 'wb') as f:
                f.write(r.content)
                f.flush()
            fz = ZipFile(driver_zipfile_path, 'r')
            for files in fz.namelist():
                fz.extract(files, driver_unzip_path)
                fz.extract(files, self.driver_backup_path)
        except:
            return False
        else:
            log("驱动下载成功！")
            return True

    def clocking_in(self):
        for key in self.profile["profiles"]:
            try:
                browser = webdriver.Edge(executable_path=self.driver_path)
            except:
                log("浏览器启动失败")
                log(key + ' 打卡失败！')
            else:
                try:
                    browser.get('https://healthreport.zju.edu.cn/ncov/wap/default/index?from=history')
                    element_username = browser.find_element(By.ID, "username")
                    element_username.send_keys(key)
                    element_password = browser.find_element(By.ID, "password")
                    element_password.send_keys(self.profile["profiles"][key])
                    element_password.send_keys(Keys.RETURN)
                except:
                    log("统一账号登录失败")
                    log(key + ' 打卡失败！')
                else:
                    try:
                        operation_list = self.list_from_lofter(self.operation_url)
                        ocr_str = ''
                        for line in operation_list:
                            operation = line.split(',')
                            if operation[1] == 'click':
                                browser.find_element(By.XPATH, operation[0]).click()
                                sleep(int(operation[2]))
                            if operation[1] == 'input':
                                browser.find_element(By.XPATH, operation[0]).send_keys(operation[2])
                            if operation[1] == 'input_ocr':
                                browser.find_element(By.XPATH, operation[0]).send_keys(ocr_str)
                            if operation[1] == 'img_ocr':
                                browser.find_element(By.XPATH, operation[0]).screenshot(self.ocr_img_path)
                                ocr = DdddOcr()
                                with open(self.ocr_img_path, 'rb')as f:
                                    ocr_str = ocr.classification(f.read())
                                log("验证码为：" + ocr_str)
                    except:
                        log("模拟输入出错")
                        log(key + ' 打卡失败！')
                    else:
                        log(key + ' 打卡成功！')
                browser.quit()

    def internet_reconnect(self):
        while True:
            if not self.internet_connection():
                if len(self.profile["wlaninfo"]) != 0:
                    log("尝试重连")
                    if not os.path.exists(self.driver_path):
                        log("驱动不存在，无法重连")
                    else:
                        wifi = pywifi.PyWiFi()  # 创建一个wifi对象
                        ifaces = wifi.interfaces()[0]  # 取第一个无限网卡
                        print("无线网卡：" + ifaces.name())  # 输出无线网卡名称
                        ifaces.disconnect()  # 断开网卡连接
                        sleep(3)  # 缓冲3秒

                        wlan_profile = pywifi.Profile()  # 配置文件
                        wlan_profile.ssid = "ZJUWLAN"  # wifi名称
                        # profile.auth = const.AUTH_ALG_SHARED  # 需要密码
                        # profile.akm.append(const.AKM_TYPE_NONE)
                        # profile.cipher = const.CIPHER_TYPE_NONE
                        # profile.akm.append(const.AKM_TYPE_WPA2PSK)  # 加密类型
                        # profile.cipher = const.CIPHER_TYPE_CCMP  # 加密单元
                        # profile.key = '4000103000' #wifi密码

                        ifaces.remove_all_network_profiles()  # 删除其他配置文件
                        tmp_profile = ifaces.add_network_profile(wlan_profile)  # 加载配置文件

                        ifaces.connect(tmp_profile)  # 连接
                        sleep(10)  # 尝试10秒能否成功连接
                        if ifaces.status() == const.IFACE_CONNECTED:
                            log("网卡连接ZJUWLAN成功")
                            if not self.internet_connection():
                                try:
                                    login_brsr = webdriver.Edge()
                                    login_brsr.get('https://net3.zju.edu.cn/srun_portal_pc?ac_id=3&theme=zju')
                                    login_brsr.find_element(By.ID, 'username').send_keys(self.profile["wifi_username"])
                                    login_brsr.find_element(By.ID, 'password').send_keys(self.profile["wifi_password"])
                                    login_brsr.find_element(By.ID, 'login').send_keys(Keys.RETURN)
                                    sleep(1)
                                    login_brsr.quit()
                                except:
                                    log("ZJUWLAN账号登录失败")
                        else:
                            log("网卡连接ZJUWLAN失败")
            sleep(600)

    def check_status(self):
        while True:
            if not os.path.exists(self.edge_path):
                self.ready = False
                self.prompt.setText("【警告】未检测到Edge浏览器,请下载")
            else:
                if not self.internet_connection():
                    self.ready = False
                    self.prompt.setText("【警告】无网络连接")
                else:
                    if not os.path.exists(self.driver_path):
                        self.ready = False
                        if not self.download_driver():
                            self.prompt.setText("【警告】驱动下载失败")
                    else:
                        self.ready = True
                        self.prompt.setText("就绪")

            sleep(5)

    def check_driver_version(self):
        driver_ver = self.parser.GetFileVersion(self.driver_path)
        edge_ver = self.parser.GetFileVersion(self.edge_path)
        log("当前驱动版本：" + str(driver_ver))
        log("当前Edge版本：" + str(edge_ver))
        if driver_ver != edge_ver:
            log("驱动版本不匹配，重新下载")
            try:
                os.remove(self.driver_path)
            except:
                log("旧驱动删除失败")
                return False
            if not self.download_driver():
                log("驱动下载失败")
                return False
        return True

    def check_clocking(self):
        while True:
            if datetime.today().hour == self.profile["time"][0]:
                if datetime.today().minute == self.profile["time"][1]:
                    if datetime.today().second == self.profile["time"][2]:
                        log("------------------------------")
                        log("开始打卡")
                        if self.ready:
                            if self.check_driver_version():
                                self.clocking_in()
                            else:
                                log("打卡取消")
                        else:
                            log(self.prompt + "\n停止启动打卡")
            sleep(0.8)


def log(string):
    log_path = os.getenv('TEMP')[:len(os.getenv('TEMP')) - 4] + "AHPLOG.txt"
    with open(log_path, 'a') as f:
        f.write("\n" + str(datetime.now())[0:19] + " " + string)


def backup_pyfiles():
    if os.path.basename(sys.argv[0]).find('exe'):
        for py in glob.glob(os.getcwd() + '\\*.py'):
            shutil.copy(py, 'D:\\backup')







