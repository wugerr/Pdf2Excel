# -*- coding:utf-8 -*-


import pdfplumber
import pandas as pd
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtGui import QFont
import sys,os
import platform
from PyQt5.QtWidgets import QWidget, QMainWindow, QDesktopWidget, QPushButton, QHBoxLayout, QVBoxLayout,QGridLayout, QMessageBox
import time
from PyQt5.QtCore import QThread, pyqtSignal

#import numpy as np
#import xlsxwriter

#Jinja2, matplotlib


#起线程处理读pdf，写excel等业务逻辑
class ConverterThread(QThread):
    # 通过类成员对象定义信号对象  
    converterSignal = pyqtSignal(dict)
    pdfFolderPath = None
    xlsxFilePath = None
    

    def initParam(self, pdfFolderPath, xlsxFilePath):
        self.pdfFolderPath = pdfFolderPath
        self.xlsxFilePath = xlsxFilePath

    '''[读取pdf数据的信息]
    解释出的行数据如下
    ['Windwoo Design & Manufacture Limited\nFloor 6, Building B, No.9 East Zone,Shangxue Industrial Zone,BanTian,LongGang\nDistrict,Shenzhen China Contact: jonly', None, None, None]
    ['Invoice', None, None, None]
    ['Date:28-04-2017 Order NO：2017042806', None, None, None]
    ['Consignee and buyer company:Bao Cao \n OLAWUYI IBUANU\nADD：Ibuanu Olawuyi Tramways D1 room G Frederick road Salford Manchester UK M6\n6BY.\nContact:OLAWUYI IBUANU\nTel:+44787626111', None, None, None]
    ['Item', 'QTY', 'FOB\nSHENZHEN', 'Total Amount']
    ['bluetooth wooden\nspeaker', '1pcs(light\nbrown color)', '45.32usd', '45.32usd']
    ['', None, None, None]
    ['Total', '', '', '45.32usd']
    
    Args:
        pdfName,pdfFile
    '''

    def getDataFromPdf(self, pdfName, pdfFile):
        print("\r\n------getDataFromPdf----pdfFile:",pdfFile)
        orderDict = {}

        pdfNameList = pdfName.split("-")
        if len(pdfNameList)>2:
            orderDict["国家"] = pdfNameList[0].replace("invoice","")
            orderDict["业务员"] = pdfNameList[1]
            orderDict["产品"] = pdfNameList[2]

        with pdfplumber.open(pdfFile) as pdf:      
            firstPage = pdf.pages[0] 
            tables = firstPage.extract_tables()
            table = tables[0]
            
            #print(table)
            df = pd.DataFrame(table) 
            # 第一列当成表头： 
            df = pd.DataFrame(table[1:],columns=table[0]) 

            
            for rowList in table:
                if rowList[0].startswith("Date:"):#日期，订单号，'Date:28-04-2017 Order NO：2017042806'
                    tempStr = rowList[0]
                    tempStr = tempStr.replace("Date:","")
                    tempList = tempStr.split("Order NO：")
                    if len(tempList) == 2:
                        orderDict["日期"] = tempList[0]
                        orderDict["订单号"] = tempList[1]
                elif rowList[0].startswith("Consignee and buyer company:"):#客户名信息
                    tempStr = rowList[0]
                    tempList = tempStr.split("\n")
                    if len(tempList)>1:
                        buyInfo = tempList[0]
                        buyInfo = buyInfo.replace("Consignee and buyer company:","")
                        orderDict["客户名"] = buyInfo.strip()
                elif rowList[0]=="Total":#订单总金额
                    totalStr = rowList[3].replace("usd","")
                    orderDict["金额"] = float(totalStr)
        #orderDict["统计"] = ""
        print(orderDict)
        return orderDict


    #编写亮绿色和暗绿色的条件
    def color(self, row):
        if row['gender']=='F':
            return ['background-color: palegreen']
        elif row['grade']>=80:
            return ['background-color: limegreen']
        return ['background-color: palegreen']


    #编写黑色和白色的条件
    def font_color(self, row):
        if row['gender']=='F':
            return ['color: black']
        elif row['grade']>=80:
            return ['color: white']
        return ['color: black']

    '''[写excel文件]
    
    [把遍历文件夹的pdf信息写进excel文件中，每一个pdf文件是excel中的一行]
    
    '''
    def writeExcel(self,pdfDataList,allTaotal):
        print("-------write excel-------")

        fontSize = "20px"
        
        writer = pd.ExcelWriter(self.xlsxFilePath,  engine="xlsxwriter", datetime_format='yyyy mmm d hh:mm:ss', date_format='yyyy-mmmm-dd ')

        import pandas.io.formats.excel
        pandas.io.formats.excel.header_style = None

        df = pd.DataFrame(pdfDataList)
        headOrder = ["订单号","日期","产品","客户名","国家","金额","业务员"]
        df = df[headOrder]
        #df.at[2,7] = "所有订单总金额" 
        #df.at[3,7] = allTaotal 
        #df.style.applymap(self.color_change).to_excel(writer, sheet_name='Sheet1', index=False)
        

        print(df)

        style1 = [
            dict(selector="th", props=[("font-size", fontSize), ("text-align", "center"),("background-color", "#FCF3CF"),('width',"150px"),('height','20px')]),
            dict(selector="td", props=[("font-size", fontSize), ("text-align", "right"),('width',"150px"),('height','50px')]),
            dict(selector="caption", props=[("caption-side", "top"),("font-size",fontSize),("font-weight","bold"),("text-align", "left"),('height','50px'),('color','#E74C3C')])
        ]

        # overwrite需要pandas1.2.0
        style2 = {
            'F': [dict(selector='td', props=[('text-align','center'),("font-weight","bold"),("text-transform","capitalize")])],
            'B': [dict(selector='td', props=[('text-align','left'),("font-style","italic")])],
            'E': [dict(selector='td', props=[('text-align','center')])],
            'C': [dict(selector='td', props=[('text-decoration','underline'),('text-decoration-color','red'),('text-decoration-style','wavy')])]
        }

        format_dict = {'金额':'${0:,.0f}', '日期': '{:%Y-%m-%d}'}

        (df.style
            .background_gradient("Greens",subset="金额")
            .applymap(lambda x: 'font-size:'+fontSize)
            .applymap(lambda x: 'white-space:pre-wrap;')
            .applymap(lambda x: "background-color:#E0FFFF", subset="业务员")
            .applymap_index(lambda x: 'font-size:'+fontSize)
            .set_table_styles(style1).set_table_styles(style2,overwrite=False)
            .set_properties(**{'font-family': 'Microsoft Yahei','border-collapse': 'collapse',
                     'border-top': '1px solid black','border-bottom': '1px solid black','font-size': fontSize})
            .format(format_dict)
            .format(na_rep='-')
            #.apply_index("font-size:24px")
            #.applymap(self.color_change,subset=["金额"])
            .to_excel(writer, sheet_name='Sheet1', header=True, index=False))
        # 使用applymap并调用写好的函数
        
        worksheet = writer.sheets['Sheet1']
        
        
        
        #设置格式的范围
        dataLen = len(pdfDataList)
        formatRangeStr = "A2:E"+str(dataLen+1) #类似这种格式'A2:F15'，15是15条数据

        
        #隔行颜色
        workbook = writer.book
        rowColorStyle1 = workbook.add_format({'bg_color':   '#F0FFF0'})#'font_color': '#9C0006'F5F5F5
        rowColorStyle2 = workbook.add_format({'bg_color':   '#F0FFFF'})
        worksheet.conditional_format(formatRangeStr, {'type':'formula','criteria': '=mod(row(),2)=1','format':rowColorStyle1})
        worksheet.conditional_format(formatRangeStr, {'type':'formula','criteria': '=mod(row(),2)=0','format':rowColorStyle2})

        #formatFont = workbook.add_format({'font_color': 'green'})
        #worksheet.set_column(formatRangeStr, None, formatFont)
        #worksheet.set_row(3, 35)
        #worksheet.set_default_row(15,  {'font_size': 24})

        
        # 创建列名的样式
        headStyle = workbook.add_format({
            'font_size':20,
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#D7E4BC',
            'border': 1})
        statisticsStyle = workbook.add_format({'font_size':20,'fg_color': 'red'})

        # 从A1单元格开始写出一行数据，指定样式为header_format
        worksheet.write_row(0, 0,  df.columns, headStyle)

        totalStr = "所有订单总金额: %.2f" % allTaotal
        worksheet.write(2,7, totalStr, statisticsStyle)

        worksheet.autofit()

        bussStyle = workbook.add_format({'bg_color': '#A4D3EE'})
        
        worksheet.set_column("A:A", 18) 
        worksheet.set_column("B:B", 18) 
        worksheet.set_column("C:C", 40) 
        worksheet.set_column("D:D", 20) 
        worksheet.set_column("E:E", 12) 
        worksheet.set_column("F:F", 7) 
        worksheet.set_column("F:F", 10) 
        worksheet.set_column("G:G", 12) 
        worksheet.set_column("H:H", 45)
        
        writer._save()


    def highlight_max(self,s):
        print("highlight")


    def color_change(self,val):
        
        color = 'red' if val < float(100) else 'blue'
        return 'color: %s' % color  


    # 大业务逻辑
    def run(self):
        emitDict = {}
        pdfDataList = []
        pdfCount = 0 
        
        emitDict["flag"] = "startProgressBar"
        self.converterSignal.emit( emitDict)
        time.sleep(0.1)
        allTaotal = 0

        for fileName in os.listdir(self.pdfFolderPath):
            if not os.path.isdir(os.path.join(self.pdfFolderPath,fileName))  and fileName.endswith(".pdf"):
                pdfCount += 1
                pdfFile = os.path.join(self.pdfFolderPath,fileName)
                emitDict["flag"] = "startPdf"
                emitDict["pdfCount"] = pdfCount
                emitDict["fileName"] = fileName
                self.converterSignal.emit( emitDict)

                orderDict = self.getDataFromPdf(fileName, pdfFile)
                singleTotal = float(orderDict["金额"])
                allTaotal += singleTotal


                pdfDataList.append(orderDict)
                time.sleep(0.1)
        

        emitDict["flag"] = "startExcel"
        self.converterSignal.emit( emitDict)
        self.writeExcel(pdfDataList,allTaotal)
        emitDict["flag"] = "endExcel"
        self.converterSignal.emit( emitDict)


class Converter(QMainWindow):

    def __init__(self,parent=None):
        #super().__init__()
        super(Converter,self).__init__(parent)

        self.INVOICE_XLSX_NAME = "invoice.xlsx"
        self.pdfFolderPath = ""
        self.xlsxFilePath = ""
        self.pdfSum = 0

        if sys.platform.startswith('linux'):
            print('当前系统为 Linux')
            self.defaultOpenFolder = "/"
        elif sys.platform.startswith('win'):
            print('当前系统为 Windows')
            self.defaultOpenFolder = "C:/"
        elif sys.platform.startswith('darwin'):
            print('当前系统为 macOS')
            self.defaultOpenFolder = "/Users/seasago/Documents"
        else:
            self.defaultOpenFolder = "/"


        self.initUI()
        #self.resize(400, 300)
 

    
    def initUI(self):
        gridColumns = 2
        screen = QDesktopWidget().screenGeometry(); #获取屏幕大小
        
        self.setFixedSize(int(screen.width()/3), int(screen.height()/3))
        
        # 设置只显示关闭按钮
        self.setWindowFlags(QtCore.Qt.WindowCloseButtonHint)
        #self.setWindowFlags(QtCore.Qt.WindowMinimizeButtonHint) 

        gridLayout = QGridLayout()
        #gridLayout.addWidget(QPushButton(str(1)))

        boldFont = QFont()
        boldFont.setPointSize(18)
        boldFont.setBold(True)

        fontSize = QFont()
        fontSize.setPointSize(18)

        #选择生成的Excel文件路径：
        self.xlsxLabel0 = QtWidgets.QLabel(self)
        #self.xlsxLabel0.adjustSize()
        self.xlsxLabel0.setAlignment(QtCore.Qt.AlignLeft)
        #self.xlsxLabel0.setGeometry(QtCore.QRect(10, 10, 100, 21)) 
        self.xlsxLabel0.setFont(boldFont)
        gridLayout.addWidget(self.xlsxLabel0,0,0,1,gridColumns,QtCore.Qt.AlignBottom)

        self.pdfTextEdit = QtWidgets.QLineEdit(self)
        self.pdfTextEdit.setObjectName("pdfTextEdit")
        self.pdfTextEdit.setReadOnly(True)#
        self.pdfTextEdit.setFixedHeight(30)
        self.pdfTextEdit.setStyleSheet("QLineEdit:read-only { background-color: #CFCFCF; }")
        self.pdfTextEdit.setText(self.defaultOpenFolder)
        gridLayout.addWidget(self.pdfTextEdit,1,0,1,gridColumns)


        #选择PDF所在目录：
        self.pdfLabel0 = QtWidgets.QLabel(self)
        #self.pdfLabel0.adjustSize()
        self.pdfLabel0.setAlignment(QtCore.Qt.AlignLeft)
        #self.pdfLabel0.setGeometry(QtCore.QRect(10, 100, 120, 21)) 
        self.pdfLabel0.setFont(boldFont)
        gridLayout.addWidget(self.pdfLabel0,2,0,1,gridColumns,QtCore.Qt.AlignBottom)


        self.excelTextEdit = QtWidgets.QLineEdit(self)
        self.excelTextEdit.setObjectName("excelTextEdit")
        self.excelTextEdit.setReadOnly(True)#
        self.excelTextEdit.setFixedHeight(30)
        self.excelTextEdit.setStyleSheet("QLineEdit:read-only { background-color: #CFCFCF; }")
        self.excelTextEdit.setText(self.defaultOpenFolder+"/"+self.INVOICE_XLSX_NAME)
        gridLayout.addWidget(self.excelTextEdit,3,0,1,gridColumns)






        # 选择PDF文件目录的按钮
        self.pdfFolderButton = QtWidgets.QPushButton(self)
        #self.pdfFolderButton.setGeometry(QtCore.QRect(10, 120, 150, 28))
        self.pdfFolderButton.setObjectName("pdfFolderButton")
        self.pdfFolderButton.setStyleSheet("QPushButton:hover{color:gray}")
 
        self.pdfFolderButton.setFixedSize(180,40)
        gridLayout.addWidget(self.pdfFolderButton,4,0,2,1,QtCore.Qt.AlignVCenter)



        # 开始工作
        self.doButton = QtWidgets.QPushButton(self)
        self.doButton.setObjectName("xlsxFileButton")
        self.doButton.setStyleSheet(
            "QPushButton:hover{color:gray}"  # 光标移动到上面后的前景色
        )

        self.doButton.setFixedSize(200,40)
        gridLayout.addWidget(self.doButton,4,1,2,1,QtCore.Qt.AlignVCenter)#QtCore.Qt.AlignCenter|QtCore.Qt.AlignBottom

        #处理进度条
        self.progressBar = QtWidgets.QProgressBar(self)
        #self.progressBar.setGeometry(QtCore.QRect(30, 100, 291, 61))
        #self.progressBar.setProperty("value", 55)
        self.progressBar.setObjectName("progressBar")
        #self.progressBar.setStyleSheet("QProgressBar {border: 2px solid green; border-radius: 5px; background-color: #FFFFFF; text-align:center; font-size:12px}")
        self.progressBar.setVisible(False)
        self.progressBar.setFixedSize(int(self.width()*0.8),10)
        gridLayout.addWidget(self.progressBar,5,0,1,gridColumns,QtCore.Qt.AlignBottom|QtCore.Qt.AlignCenter)

        
        #

        #self.setLayout(gridLayout)
        widget = QWidget()
        widget.setLayout(gridLayout)
        self.setCentralWidget(widget)


        #定义状态栏
        self.statusbar = QtWidgets.QStatusBar(self)
        # 将状态栏设置为当前窗口的状态栏
        self.setStatusBar(self.statusbar)
        # 设置状态栏的对象名称
        self.statusbar.setObjectName("statusbar")
        #设置状态栏样式
        self.statusbar.setStyleSheet('QStatusBar::item {border: none;}')


        #self.setGeometry(300, 300, 400, 100)
        self.retranslateUi()
        #self.setWindowTitle("Buttons")
        self.pdfFolderButton.clicked.connect(self.setPdfFloder)
        #self.xlsxFileButton.clicked.connect(self.setXlsxPath)
        self.doButton.clicked.connect(self.do)

    def retranslateUi(self):
        self._translate = QtCore.QCoreApplication.translate
        self.setWindowTitle(self._translate("MainWindow", "一简办公"))
        self.xlsxLabel0.setText(self._translate("MainWindow", "输出Excel文件："))
        self.pdfLabel0.setText(self._translate("MainWindow", "选择的PDF目录是："))
        self.pdfFolderButton.setText(self._translate("MainWindow", "1.选择PDF所在目录"))
        #self.xlsxFileButton.setText(self._translate("MainWindow", "2.选择生成的Excel文件路径"))
        self.doButton.setText(self._translate("MainWindow", "2.生成excel文件"))



    def setPdfFloder(self,Filepath):
        print("选择PDF所在目录")
        self.pdfFolderPath = QtWidgets.QFileDialog.getExistingDirectory(None,"选取PDF所在文件夹",self.defaultOpenFolder) 
        self.xlsxFilePath = self.pdfFolderPath+"/"+self.INVOICE_XLSX_NAME

        self.pdfTextEdit.setText((self.pdfFolderPath))
        self.excelTextEdit.setText(self.xlsxFilePath)



    def handleUI(self,dataDict):
        print("--------handleUI-------")
        flag = dataDict["flag"]
        if flag == "startProgressBar":
            print("\r\n-------xxxxxxxxxxxx-------\r\n")
            self.statusbar.showMessage(self._translate("MainWindow", "开始读取pdf文件信息"),0)
            #pdfSum = dataDict["pdfSum"]
            self.progressBar.setRange(0,self.pdfSum)
            self.progressBar.setVisible(True)
        elif flag == "startPdf":
            pdfCount = dataDict["pdfCount"]
            pdfSumStr = str(self.pdfSum)
            fileName = dataDict["fileName"]

            progressMsg = "正在处理pdf文件："+str(pdfCount)+"/"+pdfSumStr +" "+fileName
            #self.statusLabel.setText(self._translate("MainWindow", progressMsg))
            self.statusbar.showMessage(self._translate("MainWindow", progressMsg),0)
            self.progressBar.setValue(pdfCount)
        elif flag == "startExcel":
            self.statusbar.showMessage(self._translate("MainWindow", "开始导出excel文件"),0)
        elif flag == "endExcel":
            self.statusbar.showMessage(self._translate("MainWindow", "导出excel文件完成"),0)
        elif flag == "noPdf":
            self.statusbar.showMessage(self._translate("MainWindow", ("指定的路径不存在PDF文件："+self.pdfFolderPath)),0)
        elif flag == "noPdfFolder":
            self.statusbar.showMessage(self._translate("MainWindow", (self.pdfFolderPath+"指定的路径不存在")),0)
        
        
        QtWidgets.QApplication.processEvents()
        time.sleep(0.1)


    def startThread(self):
        self.converterThread = ConverterThread()
        self.converterThread.converterSignal.connect(self.handleUI)
        self.converterThread.initParam(self.pdfFolderPath,   self.xlsxFilePath)
        self.converterThread.start()

    def do(self):
        print("------生成excel文件------")

        if self.pdfFolderPath and os.path.exists(self.pdfFolderPath):
            pdfSum = 0 
            for fileName in os.listdir(self.pdfFolderPath):
                if not os.path.isdir(os.path.join(self.pdfFolderPath,fileName)) and fileName.endswith(".pdf"):
                    pdfSum += 1
            if pdfSum > 0 :
                self.pdfSum = pdfSum
                if os.path.exists(self.xlsxFilePath):
                    reply = QMessageBox().question(self, "注意", "对应excel文件已经存在pdf目录，你确认要覆盖吗？", QMessageBox.Yes|QMessageBox.No, QMessageBox.No)
                    if reply == QMessageBox.Yes:
                        self.startThread()

                else:
                    self.startThread()
            else:
                QMessageBox.critical(self, '错误',"选择的目录不存在pdf文件", QMessageBox.Close, QMessageBox.Close)

        else:
            QMessageBox.critical(self,    # 父窗口QWidget
                     '错误',    # 窗口标题
                     "选择的PDF目录不存在",      # 窗口提示信息
                      QMessageBox.Close,    
                      # 窗口内添加按钮-QMessageBox.StandardButton，可重复添加使用 | 隔开；如果不写，会有个默认的QMessageBox.StandardButton
                      QMessageBox.Close
                      )




if __name__ == "__main__":
    QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)#
    app = QtWidgets.QApplication(sys.argv)
    converter = Converter()
    converter.show()

    sys.exit(app.exec_())
        


