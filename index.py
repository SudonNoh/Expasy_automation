import os
import sys
from warnings import WarningMessage
from PyQt5.QtWidgets import (
    QApplication,
    QWidget,
    QMainWindow,
    QFileDialog,
    QPushButton,
    QHBoxLayout,
    QVBoxLayout,
    QLabel,
    QMessageBox,
    QLineEdit,
    QDesktopWidget,
    QProgressDialog
)
from PyQt5.QtGui import QIcon
from expasy import *
import time

class MainApp(QMainWindow):
    
    def __init__(self):
        super().__init__()
        self.initUI()
        
    def initUI(self):
        
        self.setWindowTitle('Expasy')
        
        self.center()
        self.width=300
        self.height=200
        self.setFixedSize(self.width, self.height)

        self.SubApp = SubApp()
        self.setCentralWidget(self.SubApp)
        
        self.show()
        
    def center(self):
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())
        
        
class SubApp(QWidget):
    
    def __init__(self):
        super().__init__()
        self.initUI()
        
    def initUI(self):
        # 1. WIDGET
        
        # 1) File Open
        self.open_label = QLabel('Files: ')
        self.open_lineedit = QLineEdit()
        self.open_btn = QPushButton('Open File')
        
        # 2) Description
        self.change_label = QLabel('Open File to Create Excel File')
        
        # 3) Create Button
        self.btn = QPushButton('Create')
        
        self.open_btn.setMaximumWidth(100)
        self.btn.setMaximumWidth(100)
        
        # 2. widget's style
        self.change_label.setStyleSheet(
            "color: #0d3300;"
            "padding: 5px;"
            "font-weight: bold;"
            "background-color: #53ff1a;"
        )
        
        # 3. click action
        self.open_btn.clicked.connect(self.openFile)
        self.btn.clicked.connect(self.createFile)
        
        # layout 
        vbox = QVBoxLayout()
        hbox1 = QHBoxLayout()
        hbox2 = QHBoxLayout()
        hbox3 = QHBoxLayout()
        hbox4 = QHBoxLayout()
        
        hbox1.addStretch(1)
        hbox1.addWidget(self.open_label)
        hbox1.addWidget(self.open_lineedit)
        hbox1.addStretch(1)
        
        hbox2.addStretch(1)
        hbox2.addWidget(self.open_btn)
        hbox2.addStretch(1)
        
        hbox3.addStretch(1)
        hbox3.addWidget(self.change_label)
        hbox3.addStretch(1)
        
        hbox4.addStretch(1)
        hbox4.addWidget(self.btn)
        hbox4.addStretch(1)
        
        vbox.addStretch(2)
        vbox.addLayout(hbox1)
        vbox.addLayout(hbox2)
        vbox.addStretch(1)
        vbox.addLayout(hbox3)
        vbox.addStretch(1)
        vbox.addLayout(hbox4)
        vbox.addStretch(2)
        
        self.setLayout(vbox)
        
    def openFile(self):
        url, _ = QFileDialog.getOpenFileName(
            caption='Select One File',
            directory='./',
            filter="excel(*.xlsx)"
        )
        print(url)
        
        if not url:
            pass
        else:
            self.open_lineedit.setText(url)
            self.open_lineedit.setReadOnly(True)
            
            self.change_label.setText('Ready To Create Excel File')
            self.change_label.setStyleSheet(
                "color: #332200;"
                "padding: 5px;"
                "font-weight: bold;"
                "background-color: #ffaa00;"
                )
        
    def createFile(self):
        
        # pb = ProgressApp()
        self.ec = ExcelControl()
        self.sc = SeleniumControl()
        
        self.file_url = self.open_lineedit.text()
        self.sheet_name = 'FRONT'
        self.site_route = 'https://web.expasy.org/protparam'
        
        self.sc.site_enter(self.site_route)
        seq_data = self.ec.excel_read(url=self.file_url, sheet_name=self.sheet_name)
        self.sc.time_sleep(3)
        
        data_list = []
        for i in seq_data:
            self.sc.input_seq(i)
            self.sc.time_sleep(5)
            data_text = self.sc.get_body()
            self.data_list.append(data_text)
            self.sc.site_back()
            self.sc.time_sleep(5)
        
        self.sc.site_close()
        
        try:
            self.ec.make_excel_file(data_list=data_list, url=self.file_url, sheet_name='ExpasyProParam')
        except PermissionError:
            self.Warning_event
            
    def Warning_event(self):
        buttonReply = QMessageBox.warning(
                        self, 
                        self.file_url+'\n 위 파일을 닫아주세요.', 
                        QMessageBox.Yes | QMessageBox.No,
                        QMessageBox.Yes
                        )
        
        if buttonReply == QMessageBox.Yes:
            self.ec.make_excel_file(data_list=self.data_list, url=self.file_url, sheet_name='ExpasyProParam')
        
        else:
            pass
            
            
# class ProgressApp(QProgressDialog):
    
#         def __init__(self):
#             super().__init__()
#             self.show()
            
#             time.sleep(10)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    first = MainApp()
    sys.exit(app.exec_())