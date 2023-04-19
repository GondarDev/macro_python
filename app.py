import os
import sys


import win32ui
import win32con

from tkinter import filedialog

import re

from PySide6.QtCore import (QCoreApplication, QMetaObject, QObject, QRect)
from PySide6.QtWidgets import (QApplication, QPushButton, QMainWindow, QMessageBox, QLabel)
from PyQt6.QtCore import pyqtSignal
from PyQt6 import QtWidgets, QtCore

import numpy as np

import pandas as pd
pd.set_option('expand_frame_repr', False)


global file_toupdate
file_toupdate=""

global update_list
update_list=[]

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.s=MergeUI()
        

    def setupUi(self, Widget):
       
        if not Widget.objectName():
            Widget.setObjectName(u"Widget")
        Widget.resize(200, 280)

        self.label = QLabel(Widget)
        self.label.setObjectName(u"label")
        self.label.setGeometry(QRect(50, 30, 100, 30))

        self.pushButton = QPushButton(Widget)
        self.pushButton.setObjectName(u"pushButton")
        self.pushButton.setGeometry(QRect(50, 60, 100, 30))
        self.pushButton.clicked.connect(self.click_merge)

        # add Button2
        self.label_2 = QLabel(Widget)
        self.label_2.setObjectName(u"label")
        self.label_2.setGeometry(QRect(50, 110, 100, 30))
        
        self.pushButton_1 = QPushButton(Widget)
        self.pushButton_1.setObjectName(u"pushButton_1")
        self.pushButton_1.setGeometry(QRect(50, 140, 100, 30))
        self.pushButton_1.clicked.connect(self.click_import)

        self.pushButton_2 = QPushButton(Widget)
        self.pushButton_2.setObjectName(u"pushButton_2")
        self.pushButton_2.setGeometry(QRect(50, 190, 100, 30))
        self.pushButton_2.clicked.connect(self.click_update)

        
        # setListitem

        self.retranslateUi(Widget)

        QMetaObject.connectSlotsByName(Widget)

        # setupUi

    def retranslateUi(self, Widget):
        Widget.setWindowTitle(
            QCoreApplication.translate("Widget", u"Macro", None))
        self.pushButton.setText(
            QCoreApplication.translate("Widget", u"Merge", None))
        self.pushButton_2.setText(
            QCoreApplication.translate("Widget", u"Update", None))
        self.pushButton_1.setText(
            QCoreApplication.translate("Widget", u"Import", None))
        self.label.setText(QCoreApplication.translate("Widget", u"Merge Section", None))

        self.label_2.setText(QCoreApplication.translate("Widget", u"Update Section", None))
    
    # retranslateUi
    def click_import(self):
        global file_toupdate
        file_toupdate = filedialog.askopenfilename()
        if file_toupdate:
            if file_toupdate.split(".")[len(file_toupdate.split("."))-1]!="csv":
                win32ui.MessageBox("Please Select CSV file", "Alert")
                return 0
            self.pushButton_1.setText("File selected")
    
    def click_merge(self):
        self.s.show()
    
    # click_update
    def click_update(self):

        # self.k.show()

        title=["Baltimore", "China", "LA", "Printful"]

        i=0

        while(i<4):
            response = win32ui.MessageBox("Do you want to update "+title[i]+" warehouse?", "Message", win32con.MB_YESNO)

            if response == win32con.IDYES:

                self.k=SelectSizeUI(i)
                i=i+1

                res=self.k.show()
                self.k.data_signal.connect(self.receive_data)
                
                # pass
            if response == win32con.IDNO:
                i=i+1

        global update_list
        self.update_result(update_list)

    def update_result(self, updatelist):

        df = pd.read_csv(file_toupdate)
        for each_list in updatelist:

            location=each_list['location']
            filepath=each_list['filepath']
            df1=pd.read_excel(filepath, sheet_name=0)
            if each_list['sizeList']==[]:
                each_list['sizeList']=[".*1.52[ ]?[m]?[ ]?[*X][ ]?3[ ]?m.*", ".*1.52[ ]?[m]?[ ]?[*X][ ]?5[ ]?m.*", ".*1.52[ ]?[m]?[ ]?[*X][ ]?10[ ]?m.*", ".*1.52[ ]?[m]?[ ]?[*X][ ]?15[ ]?m.*", ".*1.52[ ]?[m]?[ ]?[*X][ ]?18[ ]?m.*", ".*1.22[ ]?[m]?[ ]?[*X][ ]?3[ ]?m.*", ".*1.22[ ]?[m]?[ ]?[*X][ ]?5[ ]?m.*", ".*1.22[ ]?[m]?[ ]?[*X][ ]?10[ ]?m.*", ".*1.22[ ]?[m]?[ ]?[*X][ ]?15[ ]?m.*", ".*1.22[ ]?[m]?[ ]?[*X][ ]?18[ ]?m.*"]
            sizelist=each_list['sizeList']

            for index, row in df1.iterrows():                

                for each_sizeList in sizelist:

                    df['On hand'] = np.where((df['Option2 Value'].str.match(each_sizeList)&(df['Location'] == location) & (df['SKU'] == row[0])), row[len(row)-1], df['On hand'])

        response1 = win32ui.MessageBox("Do you want to Zero the China Warehouse if LA Quantity is>0", "Message", win32con.MB_YESNO)
        # print(df)
        if response1 == win32con.IDYES:
            row_index = df[(df['Location'] == "LA") & (pd.to_numeric(df['On hand'], errors='coerce').fillna(0) > 0)].index
            # row_index = df[(df['Location'] == "LA")].index
            # print(row_index)
            china_row = df.loc[row_index-1, 'On hand']=0
        
        response2 = win32ui.MessageBox("Do you want to Zero both LA & China if Baltimore Quantity is >0", "Message", win32con.MB_YESNO)

        if response2 == win32con.IDYES:

            row_index = df[(df['Location'] == "Baltimore") & (pd.to_numeric(df['On hand'], errors='coerce').fillna(0) > 0)].index
            # print(row_index)
            
            china_row = df.loc[row_index+1, 'On hand']=0
            LA_row= df.loc[row_index+2, 'On hand']=0
        response = win32ui.MessageBox("Do you want to overwrite to the previous file?", "Message", win32con.MB_YESNO)

        if response == win32con.IDYES:

            os.remove(file_toupdate)
            df.to_csv(file_toupdate, index=False, header=True)
            self.hide()

        else:

            try:

                with filedialog.asksaveasfile(mode='w', defaultextension=".csv") as file:
                    
                    df.to_csv(file.name, header=True, index=False)

                self.hide()

                return 0

            except:

                self.hide()

                return 0

    def receive_data(self, eachdata):

        global update_list
        data_copy = eachdata.copy()
        update_list.append(data_copy)

class MergeUI(QtWidgets.QWidget):
    global filename
    filename=""
    def __init__(self):
        QtWidgets.QWidget.__init__(self)

        self.setWindowTitle("Merge")
        self.setFixedHeight(350)
        self.setFixedWidth(400)
        self.pushbutton_1=QtWidgets.QPushButton("Import file to merge", self)
        self.pushbutton_1.setGeometry(QtCore.QRect(40, 100, 200, 40))
        self.pushbutton_1.clicked.connect(self.click_import)
        self.pushbutton=QtWidgets.QPushButton("Merge", self)
        self.pushbutton.setGeometry(QtCore.QRect(40, 180, 200, 40))
        self.pushbutton.clicked.connect(self.click_pushbutton)

    def click_import(self):
        global filename
        filename = filedialog.askopenfilename()
        if filename:
            if filename.split(".")[len(filename.split("."))-1]!="xlsx":
                QMessageBox.critical(self, "Warning!!!",'Please select xlsx file!!!')
            self.pushbutton_1.setText("File selected")

    def click_pushbutton(self):

        global filename

        if filename=="":
            
            QMessageBox.critical(self, "Warning", "No file Selected")
            return 0
        
        df_result=pd.DataFrame()

        df_result['Code'] = []
        df_result['Size'] = []
        df_result['Quantity']=[]

        for i in range(len(pd.read_excel(filename, None).keys())):
            df=pd.read_excel(filename, i)
            if("Unnamed" in df.columns[2]):
                df=pd.read_excel(filename, i, header=1)
            if("Code" not in df.columns):
                continue
            df.dropna(subset = ['Code'], inplace = True) 
            if("Size" in df.columns):
                df1=df.loc[:, ["Code","Size", df.columns[len(df.columns)-1]]]
                df1.columns=["Code", "Size", "every"]
                df_result = pd.merge(df_result, df1, how="outer", on=["Code", "Size"])
                df_result=df_result.replace([np.NaN], 0)
                df_result['Quantity'] = df_result['Quantity'] + df_result['every'] 
                df_result=df_result.loc[:, ["Code","Size","Quantity"]]
            else:
                df1=df.loc[:, ["Code",df.columns[len(df.columns)-1]]]
                df1.columns=["Code", "every"]
                df_result = pd.merge(df_result, df1, how="outer", on=["Code"])
                df_result=df_result.replace([np.NaN], 0)
                df_result['Quantity'] = df_result['Quantity'] + df_result['every'] 
                df_result=df_result.loc[:, ["Code","Quantity"]]

        indexTotal = df_result[ (df_result['Code'] == "Total") | (df_result['Code'] == "Sold") ].index

        df_result.drop(indexTotal , inplace=True)

        response = win32ui.MessageBox("Do you want to overwrite to the previous file?", "Message", win32con.MB_YESNO)

        if response == win32con.IDYES:
            os.remove(filename)
            df_result.to_excel(filename, index=False, header=True)
            self.hide()
        else:
            try:
                with filedialog.asksaveasfile(mode='w', defaultextension=".xlsx") as file:
                    df_result.to_excel(file.name, header=True, index=False)
                self.hide()
            except:
                self.hide()
                return 0;     

class SelectSizeUI(QtWidgets.QWidget):
        
    data_signal = pyqtSignal(dict)
    def __init__(self, data):
        QtWidgets.QWidget.__init__(self)

        global send_data
        send_data={'sizeList':[]}
        title=["Baltimore", "China", "LA", "Printful"]

        send_data['location']=title[data]
        self.setWindowTitle("Update "+title[data])
        self.setFixedHeight(350)
        self.setFixedWidth(800)

        self.pushbutton_1=QtWidgets.QPushButton("Import file to fill with", self)
        self.pushbutton_1.setGeometry(QtCore.QRect(600, 5, 180, 40))
        self.pushbutton_1.clicked.connect(self.click_import)

        self.label = QtWidgets.QLabel("Select Size",self)
        self.label.setGeometry(QtCore.QRect(40, 40, 200, 16))
        
        size_array_152=["4.98ft x 9.84ft-1.66yd x 3.28yd(1.52m x 3m)", "4.98ft*16.4ft-1.66yd*5.47yd(1.52m*5m)", "4.98ft*32.8ft-1.66yd*10.9yd(1.52m*10m)", "4.98ft*49.2ft-1.66yd*16.4yd(1.52m*15m)", "4.98ft*59ft-1.66yd*19.6yd(1.52m*18m)"]
        
        size_array_122=["4ft x 9.84ft-1.3yd x 3.28yd(1.22m x 3m)", "4ft x 16.4ft-1.3yd x 5.5yd(1.22m x 5m)", "4ft x ft-32.8yd x 10.9yd(1.22m x 10m)", "4ft x 49.2ft-16.4yd x 16.4yd(1.22m x 15m)", "4ft x 59ft-1.3yd x 19.6yd(1.22m x 18m)"]

        each_array_152=[".*1.52[ ]?[m]?[ ]?[*X][ ]?3[ ]?m.*", ".*1.52[ ]?[m]?[ ]?[*X][ ]?5[ ]?m.*", ".*1.52[ ]?[m]?[ ]?[*X][ ]?10[ ]?m.*", ".*1.52[ ]?[m]?[ ]?[*X][ ]?15[ ]?m.*", ".*1.52[ ]?[m]?[ ]?[*X][ ]?18[ ]?m.*"]

        each_array_122=[".*1.22[ ]?[m]?[ ]?[*X][ ]?3[ ]?m.*", ".*1.22[ ]?[m]?[ ]?[*X][ ]?5[ ]?m.*", ".*1.22[ ]?[m]?[ ]?[*X][ ]?10[ ]?m.*", ".*1.22[ ]?[m]?[ ]?[*X][ ]?15[ ]?m.*", ".*1.22[ ]?[m]?[ ]?[*X][ ]?18[ ]?m.*"]
        
        for index, each in enumerate(each_array_152):
            self.each = QtWidgets.QCheckBox(size_array_152[index], self)
            self.each.setGeometry(QtCore.QRect(40, index*50+50, 300, 40))
            self.each.stateChanged.connect(
                lambda state, each=each: self.checkbox_clicked(state, each))
        
        for index, each in enumerate(each_array_122):
            self.each = QtWidgets.QCheckBox(size_array_122[index], self)
            self.each.setGeometry(QtCore.QRect(440, index*50+50, 300, 40))
            self.each.stateChanged.connect(
                lambda state, each=each: self.checkbox_clicked(state, each))
    
        self.pushbutton=QtWidgets.QPushButton("OK", self)
        self.pushbutton.setGeometry(QtCore.QRect(500, 300, 100, 40))
        self.pushbutton.clicked.connect(self.click_pushbutton)
        self.show()

    def checkbox_clicked(self, state, each):
        if(state >= 2):
            send_data['sizeList'].append(each)
        if(state <= 0):
            send_data['sizeList'].remove(each)
    
    def click_pushbutton(self):
        global filename
        if(filename==""):
            QMessageBox.critical(self, "Warning!!!",'Please import file!!!')
            return 0
        global send_data
        
        self.data_signal.emit(send_data)
        filename=""
        self.hide()
    def click_import(self):
        global filename
        filename = filedialog.askopenfilename()
        if filename:
            if filename.split(".")[len(filename.split("."))-1]!="xlsx":
                QMessageBox.critical(self, "Warning!!!",'Please select xlsx file!!!')
            self.pushbutton_1.setText("File selected")
            self.data=filename
            global send_data
            send_data['filepath']=filename


app = QApplication(sys.argv)

w = MainWindow()
w.show()
app.exec()