import sys

import win32com.client
import pythoncom

# 이벤트 처리용 클래스
from PyQt5.QtWidgets import QMainWindow, QApplication

ID = "nqwrt"
PASSWD = "nn6729"
CERT = "공인인증서비밀번호"


class XASessionEventHandler:
    def __init__(self):
        self.user_obj = None
        self.com_obj = None

    def connect(self, user_obj, com_obj):
        self.user_obj = user_obj
        self.com_obj = com_obj

    def OnLogin(self, code, msg):
        self.user_obj.status = True
        print(code, msg)


# XASession 클래스
class XASession:
    def __init__(self):
        self.com_obj = win32com.client.Dispatch("XA_Session.XASession")
        self.event_handler = win32com.client.WithEvents(self.com_obj, XASessionEventHandler)
        self.event_handler.connect(self, self.com_obj)

        # demo.ebestsec.co.kr => 모의투자
        # hts.ebestsec.co.kr => 실투자
        self.com_obj.ConnectServer("demo.ebestsec.co.kr", 20001)
        self.status = False

    def login(self, id, passwd, cert):
        self.com_obj.Login(id, passwd, cert, 0, False)

        while not self.status:
            pythoncom.PumpWaitingMessages()


# 메인 윈도우
class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.session = XASession()
        self.session.login(ID, PASSWD, CERT)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MyWindow()
    window.show()
    app.exec_()

# if __name__ == "__main__":
#    session = XASession()
#    session.login("nqwrt", "nn6729", "공인인증서")
