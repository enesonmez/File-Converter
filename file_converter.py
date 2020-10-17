# -*- coding: utf-8 -*-
#-----------------------------------------------------------------------------------------------------------------------------------
#   Date   : 
#   Author : Enes Sönmez
#-----------------------------------------------------------------------------------------------------------------------------------

import sys
import os
from PyQt5.QtWidgets import (QApplication ,QWidget, QMainWindow, QLabel, QPushButton, QProgressBar, QMenuBar, QAction,
                            QTextEdit, QFileDialog, QStackedLayout, QGridLayout, QHBoxLayout, QVBoxLayout)
from PyQt5.QtGui import QIcon, QFont, QPalette, QColor
from PyQt5.QtCore import Qt 
import inspect
import time
import docx2pdf
import pptx2pdf
import txt2pdf

"""Project General Function--------------------------------------------------------------------------------------------------------
--------------------------------------------------------------------------------------------------------------------------------"""
#Function returning variable name
def retrieve_name(var):
    callers_local_vars = inspect.currentframe().f_back.f_locals.items()
    return [var_name for var_name, var_val in callers_local_vars if var_val is var]

#Converting list elements to strings
def listItemAdd(array):
    value = ''
    for arr in array:
        value += arr + '\n'
    return value


"""StackedLayoutLayer Class---------------------------------------------------------------------------------------------------------
---------------------------------------------------------------------------------------------------------------------------------"""
class StackedLayoutLayer(QWidget):
    def __init__(self):
        super(StackedLayoutLayer,self).__init__()
        self.files = []

#---Create Layers of StackedLayout--------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------
    def setMainStackedLayout(self, layout, progressBar):
        self.progressBar = progressBar
        mainpage = self.mainpageUI()
        convertpage = self.convertPageUI('')

        layout.addWidget(mainpage)
        layout.addWidget(convertpage)

#---StackedLayout Main Page UI Layer------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------
    def mainpageUI(self):
        pdf = ['DOCX to PDF', 'PPTX to PDF', 'TXT to PDF', 'HTML to PDF','PDF Merge']
        docx = ['PDF to DOCX', 'PPTX to DOCX']
        pptx = ['PDF to PPTX']
        files = [pdf,docx,pptx]
        color = ['red','#3f9fdb','orange']
        widget = QWidget()
        grid = QGridLayout()
        
        for y, file in enumerate(zip(files,color)):
            #file[0] =>lists  file[1] => colors
            variable = str(retrieve_name(file[0])[0]).upper()
            label = QLabel(variable,self)
            label.setAlignment(Qt.AlignCenter)
            label.setFont(QFont('Times',11))
            label.setStyleSheet("background-color: " + file[1])
            grid.addWidget(label,0,y)
            for x, value in enumerate(file[0]):
                label = QLabel(value,self)
                label.setAlignment(Qt.AlignTop)
                label.setFont(QFont('Times',11))
                grid.addWidget(label,x+1,y)
        
        widget.setLayout(grid)
        return widget

#---StackedLayout Convert UI Layer-------------------------------------------------------------------------------------------------- 
#-----------------------------------------------------------------------------------------------------------------------------------
    def convertPageUI(self,title):
        widget = QWidget()
        grid = QGridLayout()
        self.label = QLabel(title,self)
        self.label.setAlignment(Qt.AlignTop)
        self.label.setFont(QFont('Times',11))
        grid.addWidget(self.label,0,0)

        clearbuton = QPushButton('Clear',self)
        clearbuton.setFont(QFont('Times',11))
        clearbuton.setStyleSheet("background-color:green")
        clearbuton.clicked.connect(self.reviewEditClear)
        grid.addWidget(clearbuton,0,6)


        self.reviewEdit = QTextEdit()
        self.reviewEdit.setFont(QFont('Times',11))
        self.reviewEdit.setEnabled(False)
        grid.addWidget(self.reviewEdit,1,0,1,7)

        label2 = QLabel('Choose file :',self)
        label2.setFont(QFont('Times',11))
        grid.addWidget(label2,2,0)

        buton = QPushButton('Select',self)
        buton.setFont(QFont('Times',11))
        buton.clicked.connect(self.openFileName)
        grid.addWidget(buton,2,1,1,6)

        butonCreate = QPushButton('Convert',self)
        butonCreate.setFont(QFont('Times',11))
        butonCreate.setStyleSheet("background-color:red")
        
        butonCreate.clicked.connect(self.convertSwitch)
        grid.addWidget(butonCreate,3,0,1,7)

        widget.setLayout(grid)
        return widget

#---ReviewEdit Tetxbox Clear-----------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------  
    def reviewEditClear(self):
        self.files.clear()
        self.reviewEdit.setText('')

#---Convert Process-----------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------        
    def convertSwitch(self):
        processName = self.label.text().lower()
        if processName == 'docx to pdf':
            self.docx2pdf()
        elif processName == 'pptx to pdf':
            self.pptx2pdf()
        elif processName == 'txt to pdf':
            self.txt2pdf()
        else:
            print(processName)
    
    def docx2pdf(self):
        if self.files:
            for file in self.files:
                file_in, file_out = self.filenameFix(file,'.pdf')
                self.progressBar.setValue(0)
                docx2pdf.docx2pdf(file_in,file_out)
                self.progressBar.setValue(100)
            self.convertPageUIWidgetReset()
    
    def pptx2pdf(self):
        if self.files:
            for file in self.files:
                file_in, file_out = self.filenameFix(file,'.pdf')
                self.progressBar.setValue(0)
                pptx2pdf.PPTXtoPDF(file_in,file_out)
                self.progressBar.setValue(100)
            self.convertPageUIWidgetReset()

    def txt2pdf(self):
        if self.files:
            for file in self.files:
                file_in, file_out = self.filenameFix(file,'.pdf')
                self.progressBar.setValue(0)
                txt2pdf.TXTtoPDF(file_in,file_out)
                self.progressBar.setValue(100)
            self.convertPageUIWidgetReset()
    
    def filenameFix(self,file,filetype):
        file_in = file
        path, filename = os.path.split(file)
        file_out = path + '/' + os.path.splitext(filename)[0] + filetype
        return file_in, file_out

    def convertPageUIWidgetReset(self):
        self.reviewEditClear()
        time.sleep(2)
        self.progressBar.reset() 
        

#---Open File Name Functions--------------------------------------------------------------------------------------------------------
#----------------------------------------------------------------------------------------------------------------------------------- 
    def openFileName(self):
        if self.label.text():
            fileType = self.label.text().split()[0].lower()
            if fileType == 'docx':
                self.openFileNameDialog(fileType.upper() + ' Files (*.' + fileType + ' *.doc)')
            elif fileType == 'pptx':
                self.openFileNameDialog(fileType.upper() + ' Files (*.' + fileType + ' *.ppt)')
            elif fileType == 'txt':
                self.openFileNameDialog(fileType.upper() + ' Files (*.' + fileType + ')')

    def openFileNameDialog(self,filetype):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName(self,"OpenFileName", "",filetype, options=options)
        x = ''
        try:
            x = str(self.files.index(fileName))
        except ValueError:
            if fileName and (not x):
                self.files.append(fileName)
                self.reviewEdit.setText(listItemAdd(self.files))


"""App Class------------------------------------------------------------------------------------------------------------------------
---------------------------------------------------------------------------------------------------------------------------------"""
class App(QMainWindow,StackedLayoutLayer):
    def __init__(self):
        super(App,self).__init__()
        self.files = []
        self.title = "File Converter"
        self.icon = 'img/icon.png'
        self.left = 500
        self.top = 200
        self.width = 640
        self.height = 480
        self.initUI()

        self.show()

#---App UI--------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------
    def initUI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left,self.top,self.width,self.height)#setting position and size what window
        self.setMaximumSize(self.width,self.height)
        self.setMinimumSize(self.width,self.height)
        self.setWindowIcon(QIcon(self.icon))

        #Layout
        self.pagelayout = QVBoxLayout()
        self.menulayout = QHBoxLayout()
        self.statuslayout = QHBoxLayout()
        self.layout = QStackedLayout()

        self.pagelayout.addLayout(self.menulayout)
        self.pagelayout.addLayout(self.layout)
        self.pagelayout.addLayout(self.statuslayout)

        #MenubarUI
        self.menubarUI()

        #StatusbarUI
        self.statusbarUI('No Process')

        #MainpageUI    
        self.setMainStackedLayout(self.layout,self.progressBar)

        widget = QWidget()
        widget.setLayout(self.pagelayout)
        self.setCentralWidget(widget)

#---MenuBar UI----------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------
    def menubarUI(self):
        #Menubar
        self.mainmenu = self.menuBar()
        self.mainmenu.setFont(QFont('Times',12))
        #File
        file = self.mainmenu.addMenu('File')
        file.setFont(QFont('Times',11))
        #Action Add (Main PAGE)
        mainpage = QAction('Main Page',self)
        mainpage.triggered.connect(self.mainPageLinkPress)
        file.addAction(mainpage)
        #Action Add (Exit)
        exitlink = QAction('Exit',self)
        exitlink.setShortcut('Ctrl+Q')
        exitlink.triggered.connect(self.close)
        file.addAction(exitlink)

        #PDF
        pdf = self.mainmenu.addMenu('PDF')
        pdf.setFont(QFont('Times',11))
        #Action Add (DOCX to PDF)
        docx2pdf = QAction(QIcon('img/doc2pdf.png'),'DOCX to PDF',self)
        docx2pdf.triggered.connect(self.docx2pdfLinkPress)
        pdf.addAction(docx2pdf)
        #Action Add (PPTX to PDF)
        pptx2pdf = QAction(QIcon('img/ppt2pdf.png'),'PPTX to PDF',self)
        pptx2pdf.triggered.connect(self.pptx2pdfLinkPress)
        pdf.addAction(pptx2pdf)
        #Action Add (PPTX to PDF)
        txt2pdf = QAction(QIcon('img/ppt2pdf.png'),'TXT to PDF',self)
        txt2pdf.triggered.connect(self.txt2pdfLinkPress)
        pdf.addAction(txt2pdf)


        self.menulayout.addWidget(self.mainmenu)

#---StatusBar UI--------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------
    def statusbarUI(self,notification):
        self.notificationlabel = QLabel('Process State : ' + notification,self)
        self.notificationlabel.setFont(QFont('Times',10))
        self.statuslayout.addWidget(self.notificationlabel)

        self.progressBar = QProgressBar()
        self.progressBar.setVisible(False)
        self.statuslayout.addWidget(self.progressBar)        
    
    def setLabelText(self,labelWidget,notification):
        labelWidget.setText('Process State : ' + notification)


#---StackedLayout Layer Switch------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------  
    def mainPageLinkPress(self):
        self.layout.setCurrentIndex(0)
        self.setLabelText(self.notificationlabel,'No Process')
        self.progressBar.setVisible(False)

    def docx2pdfLinkPress(self):
        self.stackPageUIWidgetSettigs(1,'DOCX to PDF')

    def pptx2pdfLinkPress(self):
        self.stackPageUIWidgetSettigs(1,'PPTX to PDF')
    
    def txt2pdfLinkPress(self):
        self.stackPageUIWidgetSettigs(1,'TXT to PDF')
    
    def stackPageUIWidgetSettigs(self,id,text):
        self.layout.setCurrentIndex(id)
        self.label.setText(text)
        self.reviewEditClear()
        self.setLabelText(self.notificationlabel,'')
        self.progressBar.setVisible(True)



if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec())