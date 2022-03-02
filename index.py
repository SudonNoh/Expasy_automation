import os
import sys
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
        self.sheet_name = 'FRONT'
        
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
        
        print('create')
        pb = ProgressApp()
        
        
class ProgressApp(QProgressDialog):
    
        def __init__(self):
            super().__init__()
            self.show()
            
            time.sleep(10)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    first = MainApp()
    sys.exit(app.exec_())