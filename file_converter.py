import sys
import os
from PyQt5.QtWidgets import *
from PyQt5.QtGui import QIcon, QFont, QPalette, QColor
from PyQt5.QtCore import Qt 
import inspect
import time
import docx2pdf
import pptx2pdf

#Değişken ismi döndüren fonksiyon
def retrieve_name(var):
    callers_local_vars = inspect.currentframe().f_back.f_locals.items()
    return [var_name for var_name, var_val in callers_local_vars if var_val is var]

#Liste elemanlarını string hale getirme
def listItemAdd(array):
    value = ''
    for arr in array:
        value += arr + '\n'
    return value


class App(QMainWindow):
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

#---Uygulama Arayüzü----------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------
    def initUI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left,self.top,self.width,self.height)#pencerenin konumunu ve boyutunu ayarlar.
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
        self.setMainStackedLayout(self.layout)

        widget = QWidget()
        widget.setLayout(self.pagelayout)
        self.setCentralWidget(widget)

#---MenuBar Arayüzü-----------------------------------------------------------------------------------------------------------------
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
        #Action Add (Docx to PDF)
        docx2pdf = QAction(QIcon('img/doc2pdf.png'),'Docx to PDF',self)
        docx2pdf.triggered.connect(self.docx2pdfLinkPress)
        pdf.addAction(docx2pdf)
        #Action Add (PPTX to PDF)
        pptx2pdf = QAction(QIcon('img/ppt2pdf.png'),'PPTX to PDF',self)
        #pptx2pdf.triggered.connect(self.pptx2pdfLinkPress)
        pdf.addAction(pptx2pdf)


        self.menulayout.addWidget(self.mainmenu)

#---StatusBar Arayüzü---------------------------------------------------------------------------------------------------------------
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

#---Stacked Layout Katmanlarını Oluşturma-------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------
    def setMainStackedLayout(self, layout):
        mainpage = self.mainpageUI()
        convertpage = self.convertPageUI('')

        layout.addWidget(mainpage)
        layout.addWidget(convertpage)

#---Stacked Layout Ana Sayfa Arayüzü------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------
    def mainpageUI(self):
        pdf = ['Docx to PDF', 'PPTX to PDF', 'PDF Merge']
        docx = ['PDF to Docx', 'PPTX to Docx']
        pptx = ['PDF to PPTX']
        files = [pdf,docx,pptx]
        color = ['red','blue','orange']
        widget = QWidget()
        grid = QGridLayout()
        
        for y, file in enumerate(zip(files,color)):
            #file[0] =>listeler  file[1] => renkler
            variable = str(retrieve_name(file[0])[0]).upper()
            label = QLabel(variable,self)
            label.setAlignment(Qt.AlignTop)
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

#---Stacked Layout Dönüştürme Arayüzü----------------------------------------------------------------------------------------------- 
#-----------------------------------------------------------------------------------------------------------------------------------
    def convertPageUI(self,title):
        widget = QWidget()
        grid = QGridLayout()
        self.label = QLabel(title,self)
        self.label.setAlignment(Qt.AlignTop)
        self.label.setFont(QFont('Times',11))
        grid.addWidget(self.label,0,0)

        self.reviewEdit = QTextEdit()
        self.reviewEdit.setFont(QFont('Times',11))
        self.reviewEdit.setEnabled(False)
        grid.addWidget(self.reviewEdit,1,0,1,7)

        label2 = QLabel('Choose file',self)
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

#---Stacked Layout Geçişleri--------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------  
    def mainPageLinkPress(self):
        self.layout.setCurrentIndex(0)
        self.setLabelText(self.notificationlabel,'No Process')
        self.progressBar.setVisible(False)

    def docx2pdfLinkPress(self):
        self.layout.setCurrentIndex(1)
        self.label.setText('Docx to PDF')
        self.files.clear()
        self.reviewEdit.setText('')
        self.setLabelText(self.notificationlabel,'')
        self.progressBar.setVisible(True)
        
#---Convert İşlemleri---------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------        
    def convertSwitch(self):
        processName = self.label.text().lower()
        if processName == 'docx to pdf':
            self.docx2pdf()
        else:
            print(processName)
    
    def docx2pdf(self):
        if self.files:
            for file in self.files:
                file_in = file
                path, filename = os.path.split(file)
                file_out = path + '/' + os.path.splitext(filename)[0] + '.pdf'
                #print(file_in,file_out)
                self.progressBar.setValue(0)
                docx2pdf.docx2pdf(file_in,file_out)
                self.progressBar.setValue(100)
            self.files.clear()
            self.reviewEdit.setText('')
            time.sleep(2)
            self.progressBar.reset()
                
        
        

#---Open File Name Fonksiyonları----------------------------------------------------------------------------------------------------
#----------------------------------------------------------------------------------------------------------------------------------- 
    def openFileName(self):
        if self.label.text():
            fileType = self.label.text().split()[0].lower()
            if fileType == 'docx':
                self.openFileNameDialog(fileType.title() + ' Files (*.' + fileType + ' *.doc)')

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




if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec())