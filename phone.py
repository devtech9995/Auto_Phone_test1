# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'Mainwindow.ui'
#
# Created by: PyQt5 UI code generator 5.15.6
#
import os
import re
import cv2
import time
import pyperclip
import pyautogui
import win32gui
import win32con
import ctypes
import pickle
import pytesseract
import numpy as np
import datetime
import csv
import keyboard
import configparser
from PyQt5.QtGui import QIcon, QFont
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtWidgets import QFileDialog, QMainWindow, QApplication, QMessageBox
import pickle
from selenium import webdriver
# import browser

def load_pickle(filename):
    with open(filename, 'rb') as handle:
        data = pickle.load(handle)
        return data

def save_pickle(dic):
    with open('filename.pickle', 'wb') as handle:
        pickle.dump(dic, handle, protocol=pickle.HIGHEST_PROTOCOL)  

class Ui_MainWindow(QMainWindow):
    flag_active  = True
    flag_error = [1] * 4
    deal_number = [0] * 4        
    phone_numbers = [[]] * 4
    check_result = {}
    rowStep = 520
    colStep = 960
    file_path = [''] * 4
    save_filename = [''] * 4
    result_filename = [''] * 4
    
    movie = ''

    replaceDir = {
        'o':'p', 'B':'8', 'X':'x', 'C':'c', '?':'2', 'Z':'2', 'z':'2', 'i':'7', 'S':'8', 
        '/':'7', 'é':'6', 'j':'3', ' ':'', 'W':'w', '&':'8', '“':'', 'J':'2', 's':'3', 
        "D":'b', '*':'f', '.':'', '£':'f', 'O':'o', 'P':'p', 'V':'v', 'r':'w', '`':'', 
        '_':'', 't':'f', '-':'', 't':'f', 'h':'b', 'A':'4', 'T':'7', '9':'g', 'q':'g', 
        ',':'', 'l':'7', '0':'n', '{':'f', '(':'f', 'I':'2', '$':'5', '^':'' , '1':'2', 
        '¥':'y', '[':'7', 'N':'n', 'u':'w', 'F':'f', 'G':'g', '#':'f', 'E':'6', 'k':'x', 
        'K':'x', '‘':'', 'Y':'y', '§':'5', '¢':'c', '€':'c'
    }

    class csv_open():
        file_path = ''
        def __init__(self, file_path):
            self.file_path = file_path
        
        def get_phone_number(self):
            phone_number = []
            with open(self.file_path, 'r', errors="ignore") as file:
                reader = csv.reader(file) 
                for row in reader:
                    # phone_numbers_in_row = re.findall(r'\d{3}[-\.\s]??\d{3}[-\.\s]??\d{4}', row[0])
                    phone_numbers_in_row = [s for s in row[0].split() if len(s) == 12]
                    # Remove hyphens and spaces from phone numbers
                    phone_numbers_in_row = [number.replace('-', '').replace(' ', '') for number in phone_numbers_in_row]
                    phone_number.extend(phone_numbers_in_row)
                    # tmp = str(row[0].replace('-', ''))     
                    # phone_number.append(tmp)
            return phone_number
    
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(538, 284)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.splitter = QtWidgets.QSplitter(self.centralwidget)
        self.splitter.setGeometry(QtCore.QRect(20, 20, 500, 241))
        self.splitter.setOrientation(QtCore.Qt.Vertical)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.splitter.sizePolicy().hasHeightForWidth())
        self.splitter.setSizePolicy(sizePolicy)
        self.splitter.setObjectName("splitter")
        self.groupBox = QtWidgets.QGroupBox(self.splitter)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.groupBox.sizePolicy().hasHeightForWidth())
        self.groupBox.setSizePolicy(sizePolicy)
        self.groupBox.setMinimumSize(QtCore.QSize(0, 90))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.groupBox.sizePolicy().hasHeightForWidth())
        self.groupBox.setSizePolicy(sizePolicy)
        self.groupBox.setObjectName("groupBox")
        self.layoutWidget = QtWidgets.QWidget(self.groupBox)
        self.layoutWidget.setGeometry(QtCore.QRect(30, 20, 451, 30))
        self.layoutWidget.setObjectName("layoutWidget")
        self.horizontalLayout = QtWidgets.QHBoxLayout(self.layoutWidget)
        self.horizontalLayout.setSizeConstraint(QtWidgets.QLayout.SetFixedSize)
        self.horizontalLayout.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.lineEdit = QtWidgets.QLineEdit(self.layoutWidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.lineEdit.sizePolicy().hasHeightForWidth())
        self.lineEdit.setSizePolicy(sizePolicy)
        self.lineEdit.setMinimumSize(QtCore.QSize(410, 28))
        self.lineEdit.setMaximumSize(QtCore.QSize(16777215, 28))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.lineEdit.sizePolicy().hasHeightForWidth())
        self.lineEdit.setSizePolicy(sizePolicy)
        self.lineEdit.setObjectName("lineEdit")
        self.horizontalLayout.addWidget(self.lineEdit)
        self.add_csv = QtWidgets.QToolButton(self.layoutWidget)
        self.add_csv.setMinimumSize(QtCore.QSize(32, 28))
        self.add_csv.setMaximumSize(QtCore.QSize(32, 28))
        self.add_csv.setObjectName("add_csv")
        self.add_csv.setFont(QFont('Times', 12))
        self.add_csv.clicked.connect(self.open_file)
        self.horizontalLayout.addWidget(self.add_csv)
        self.layoutWidget1 = QtWidgets.QWidget(self.groupBox)
        self.layoutWidget1.setGeometry(QtCore.QRect(30, 50, 451, 30))
        self.layoutWidget1.setObjectName("layoutWidget1")
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout(self.layoutWidget1)
        self.horizontalLayout_2.setSizeConstraint(QtWidgets.QLayout.SetFixedSize)
        self.horizontalLayout_2.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.lineEdit_2 = QtWidgets.QLineEdit(self.layoutWidget1)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.lineEdit_2.sizePolicy().hasHeightForWidth())
        self.lineEdit_2.setSizePolicy(sizePolicy)
        self.lineEdit_2.setMinimumSize(QtCore.QSize(410, 28))
        self.lineEdit_2.setMaximumSize(QtCore.QSize(16777215, 28))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.lineEdit_2.sizePolicy().hasHeightForWidth())
        self.lineEdit_2.setSizePolicy(sizePolicy)
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.horizontalLayout_2.addWidget(self.lineEdit_2)
        self.add_csv_2 = QtWidgets.QToolButton(self.layoutWidget1)
        self.add_csv_2.setMinimumSize(QtCore.QSize(32, 28))
        self.add_csv_2.setMaximumSize(QtCore.QSize(32, 28))
        self.add_csv_2.setObjectName("add_csv_2")
        self.add_csv_2.clicked.connect(self.open_file)
        self.add_csv_2.setFont(QFont('Times', 12))
        self.horizontalLayout_2.addWidget(self.add_csv_2)
        self.groupBox_2 = QtWidgets.QGroupBox(self.splitter)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.groupBox_2.sizePolicy().hasHeightForWidth())
        self.groupBox_2.setSizePolicy(sizePolicy)
        self.groupBox_2.setMinimumSize(QtCore.QSize(0, 90))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.groupBox_2.sizePolicy().hasHeightForWidth())
        self.groupBox_2.setSizePolicy(sizePolicy)
        self.groupBox_2.setObjectName("groupBox_2")
        self.layoutWidget2 = QtWidgets.QWidget(self.groupBox_2)
        self.layoutWidget2.setGeometry(QtCore.QRect(30, 20, 451, 30))
        self.layoutWidget2.setObjectName("layoutWidget2")
        self.horizontalLayout_3 = QtWidgets.QHBoxLayout(self.layoutWidget2)
        self.horizontalLayout_3.setSizeConstraint(QtWidgets.QLayout.SetFixedSize)
        self.horizontalLayout_3.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        self.lineEdit_3 = QtWidgets.QLineEdit(self.layoutWidget2)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.lineEdit_3.sizePolicy().hasHeightForWidth())
        self.lineEdit_3.setSizePolicy(sizePolicy)
        self.lineEdit_3.setMinimumSize(QtCore.QSize(410, 28))
        self.lineEdit_3.setMaximumSize(QtCore.QSize(16777215, 28))
        self.lineEdit_3.setObjectName("lineEdit_3")
        self.horizontalLayout_3.addWidget(self.lineEdit_3)
        self.add_csv_3 = QtWidgets.QToolButton(self.layoutWidget2)
        self.add_csv_3.setMinimumSize(QtCore.QSize(32, 28))
        self.add_csv_3.setMaximumSize(QtCore.QSize(32, 28))
        self.add_csv_3.setFont(QFont('Times', 12))
        self.add_csv_3.setObjectName("add_csv_3")
        self.add_csv_3.clicked.connect(self.open_file)
        self.horizontalLayout_3.addWidget(self.add_csv_3)
        self.layoutWidget3 = QtWidgets.QWidget(self.groupBox_2)
        self.layoutWidget3.setGeometry(QtCore.QRect(30, 50, 451, 30))
        self.layoutWidget3.setObjectName("layoutWidget3")
        self.horizontalLayout_4 = QtWidgets.QHBoxLayout(self.layoutWidget3)
        self.horizontalLayout_4.setSizeConstraint(QtWidgets.QLayout.SetFixedSize)
        self.horizontalLayout_4.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_4.setObjectName("horizontalLayout_4")
        self.lineEdit_4 = QtWidgets.QLineEdit(self.layoutWidget3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.lineEdit_4.sizePolicy().hasHeightForWidth())
        self.lineEdit_4.setSizePolicy(sizePolicy)
        self.lineEdit_4.setMinimumSize(QtCore.QSize(410, 28))
        self.lineEdit_4.setMaximumSize(QtCore.QSize(16777215, 28))
        self.lineEdit_4.setObjectName("lineEdit_4")
        self.horizontalLayout_4.addWidget(self.lineEdit_4)
        self.add_csv_4 = QtWidgets.QToolButton(self.layoutWidget3)
        self.add_csv_4.setMinimumSize(QtCore.QSize(32, 28))
        self.add_csv_4.setMaximumSize(QtCore.QSize(32, 28))
        self.add_csv_4.setObjectName("add_csv_4")
        self.add_csv_4.clicked.connect(self.open_file)
        self.add_csv_2.setFont(QFont('Times', 12))
        self.horizontalLayout_4.addWidget(self.add_csv_4)
        self.layoutWidget4 = QtWidgets.QWidget(self.splitter)
        self.layoutWidget4.setObjectName("layoutWidget4")
        self.horizontalLayout_5 = QtWidgets.QHBoxLayout(self.layoutWidget4)
        self.horizontalLayout_5.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_5.setObjectName("horizontalLayout_5")
        self.start_emul = QtWidgets.QPushButton(self.layoutWidget4)
        self.start_emul.setMinimumSize(QtCore.QSize(120, 32))
        self.start_emul.setFont(QFont('Times', 14))
        self.start_emul.clicked.connect(self.execute)
        self.start_emul.setObjectName("start_emul")
        self.horizontalLayout_5.addWidget(self.start_emul)
        self.pauseEmul = QtWidgets.QPushButton(self.layoutWidget4)
        self.pauseEmul.setMinimumSize(QtCore.QSize(120, 32))
        self.pauseEmul.setFont(QFont('Times', 14))
        self.pauseEmul.setObjectName("pauseEmul")
        self.pauseEmul.clicked.connect(self.pause)
        self.horizontalLayout_5.addWidget(self.pauseEmul)
        self.clear_list = QtWidgets.QPushButton(self.layoutWidget4)
        self.clear_list.setMinimumSize(QtCore.QSize(120, 32))
        self.clear_list.setFont(QFont('Times', 14))
        self.clear_list.setObjectName("clear_list")
        self.clear_list.clicked.connect(self.clear_lists)
        self.horizontalLayout_5.addWidget(self.clear_list)
        self.clear_list_2 = QtWidgets.QPushButton(self.layoutWidget4)
        self.clear_list_2.setMinimumSize(QtCore.QSize(120, 32))
        self.clear_list_2.setFont(QFont('Times', 14))
        self.clear_list_2.setObjectName("clear_list_2")
        self.horizontalLayout_5.addWidget(self.clear_list_2)
        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "NTT フレッツ光判定システム"))
        self.add_csv.setText(_translate("MainWindow", "..."))
        self.add_csv_2.setText(_translate("MainWindow", "..."))
        self.add_csv_3.setText(_translate("MainWindow", "..."))
        self.add_csv_4.setText(_translate("MainWindow", "..."))
        self.start_emul.setText(_translate("MainWindow", "開  始"))
        self.pauseEmul.setText(_translate("MainWindow", "停  止"))
        self.clear_list.setText(_translate("MainWindow", "削  除"))
        self.clear_list_2.setText(_translate("MainWindow", "保  管"))
        self.groupBox.setTitle(_translate("MainWindow", "東日本"))
        self.groupBox_2.setTitle(_translate("MainWindow", "西日本"))
        
    def double_click(self, point):
        pyautogui.doubleClick(point)
    
    # Click event
    def click(self, point):
        pyautogui.click(point)
    
    # Get the snapshot of CAPTCHA image
    def get_snapshot(self, point_img, img_width, img_height):
        snapshot = pyautogui.screenshot(region=(point_img[0], point_img[1], img_width, img_height))
        image = np.array(snapshot)
        return image
    
    # Reference a dictionary
    def filterText(self, captcha):
        result = ''
        for each in captcha:
            if each in self.replaceDir:
                result += self.replaceDir[each]
            else:
                result += each
        # print(result, ' ', len(result))
        return result
    
    # Recognize the CAPTCHA image using OCR
    def perform_ocr(self, img):
        img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)        
        # tmp = np.ones((10, img.shape[1]), np.uint8) * 255 
        img = np.vstack((img, np.ones((10, img.shape[1])) * 255)).astype(np.uint8)        
            
        img[:, :52] = 255
        img[:, 165:] = 255

        img = cv2.GaussianBlur(img, (5, 3), 1)        

        _, thresh = cv2.threshold(img, 84, 255, cv2.THRESH_BINARY)
        bottom = thresh[43:]

        thresh = cv2.GaussianBlur(thresh, (3, 3), 1)        

        kernel = np.ones((2, 1), np.uint8)
        dilated = cv2.dilate(thresh, kernel, iterations=3)
                
        
        kernel = np.ones((2, 1), np.uint8)
        eroded = cv2.erode(dilated, kernel, iterations=1)
        eroded[43:, :] = bottom        
        text = self.filterText(pytesseract.image_to_string(eroded).strip())
        if text == '' :
            text = 'error' 
        return text
    
    def pause(self):        
        QApplication.processEvents()   
        _translate = QtCore.QCoreApplication.translate
        self.add_csv.setEnabled(self.flag_active)
        self.clear_list.setEnabled(self.flag_active)
        self.flag_active = not self.flag_active
        if self.flag_active == True:
            print(self.flag_active, "Resumed.")
            self.pauseEmul.setText(_translate("Form", "停  止"))
            self.execute()
        else:
            save_pickle(self.check_result)
            filenames = ''
            i = 0
            for each in self.file_path:
                if len(each) > 0:
                    new_filename = each.split("/")[-1]
                    current_datetime = datetime.datetime.now()
                    formatted_datetime = current_datetime.strftime("%H-%M-%S_%Y%m%d")
                    tmpresult = filenames.join("完了_" + formatted_datetime + "_" + new_filename)    
                    cwd = os.getcwd()
                    for filename in os.listdir(cwd):
                        if new_filename in filename:
                            if "完了" in filename:
                                file_path = os.path.join(cwd, filename)
                                os.remove(file_path)
                                break
                    print(tmpresult)
                    # if new_filename in self.save_filename:
                    #     index = self.save_filename.index(new_filename)
                    #     tmpresult = self.result_filename[index]
                    #     mode = 'r+'
                    # else:
                    #     tmpresult = filenames.join("完了_" + formatted_datetime + "_" + new_filename)
                    #     mode = 'w'
                    # self.result_filename[i] = tmpresult
                    # self.save_filename[i] = new_filename
                    with open(tmpresult, mode='w', newline='', errors="ignore") as file:    
                        reader = csv.reader(open(each, mode='r', errors="ignore"))
                        writer = csv.writer(file)
                        data = load_pickle("filename.pickle")

                        keys = data.keys()
                        for row in reader:
                            temp_key = str(row[0]).replace('-', '')
                            if temp_key in keys:
                                line = ""                    
                                status = "todo"
                                if data[temp_key] == 0:
                                    status = "failed"
                                    line = r"線路無"
                                if data[temp_key] == 1:
                                    status = "success"
                                    line = r"線路有り"
                                row.insert(1, status)
                                row.insert(20, line)
                                writer.writerow(row)
                    # current_name = tmpresult
                    # new_name = filenames.join("完了_" + formatted_datetime + "_" + new_filename)
                    # self.result_filename[i] = new_name
                    # os.rename(current_name, new_name)
                    # print(self.result_filename[i])
                    # i += 1
            print(self.flag_active, "Stopped! Press play button to resume.")
            self.pauseEmul.setText(_translate("Form", "再  開"))
        
    # Monitoring the title of window
    def monitor(self):
        hwnd = win32gui.GetForegroundWindow()
        window_title = win32gui.GetWindowText(hwnd)
        return window_title
    
    def flatten(self, matrix):
        flat_list = []
        for row in matrix:
            flat_list.extend(row)
        return flat_list
    
    def calculate_csv_num(self):
        count = 0
        for element in self.file_path:
            if element == '':
                count += 1
        count = 4-count
        return(count)
    
    def clear_lists(self):
        self.file_path.clear()
        self.lineEdit.clear()
        self.lineEdit_2.clear()
        self.lineEdit_3.clear()
        self.lineEdit_4.clear()
        self.start_emul.setEnabled(True)
        _translate = QtCore.QCoreApplication.translate
        self.pauseEmul.setText(_translate("Form", "停  止"))
        self.flag_active = True
        print("Initiated.")
        
    def open_file(self):
        sender_button = self.sender()
        file, _ = QFileDialog.getOpenFileName(self, 'Open File', '', 'CSV Files (*.csv*)')
        if not self.file_path.__contains__(file):
            if file:
                print("Selected file:", file)
                filename = file.split("/")[-1] + "; "
                
                if sender_button == self.add_csv:
                    self.lineEdit.setText(filename)
                    self.init(0, file)
                elif sender_button == self.add_csv_2:
                    self.lineEdit_2.setText(filename)
                    self.init(1, file)
                elif sender_button == self.add_csv_3:
                    self.lineEdit_3.setText(filename)
                    self.init(2, file)
                elif sender_button == self.add_csv_4:
                    self.lineEdit_4.setText(filename)
                    self.init(3, file)
                
            else:
                print("No file selected.")
        
    def init(self, index, file):
        self.file_path[index] = file
        self.phone_numbers[index] = self.csv_open(file).get_phone_number()
        for phone in self.phone_numbers[index]:            
            self.check_result[str(phone)] = -1
        # print(self.check_result[index])
                
    def execute(self):
        # Decide to load pickle file  0:no, 1:load
        config = configparser.ConfigParser()
        config.read('setting.ini')
        # print(config)
        if config['main']['key'] == '1':
            print("config creating...")
            data = {}
            save_pickle(data)
        else:
            print("config loading")
            self.check_result = load_pickle('filename.pickle')
        print(self.check_result)    
        # print(self.file_path)
        csv_num = self.calculate_csv_num()
                       
        self.add_csv.setEnabled(False)
        self.clear_list.setEnabled(False)
        self.clear_list_2.setEnabled(False)
        self.start_emul.setEnabled(False)

        phoneNum_list = [[], [], [], []]
        self.click((730,30))
        while self.flag_active:
        # while True:
            for row in range(2):
                for col in range(2):
                    QApplication.processEvents()
                    window_title = self.monitor()
                    if "線路情報開示画面" in window_title and self.flag_active:
                        index = row * 2 + col
                        if self.file_path[index] == '':
                            continue                           
                        print(index, window_title, self.flag_active)                        
                        x = col * self.colStep 
                        y = row * self.rowStep
                        self.point_num = (x+170, y+280)
                        self.point_img = (x+16, y+324)
                        self.point_recog_text = (x+120, y+450)
                        self.point_btn = (x+100, y+500)
                        
                        init_click_point = (x + 720, y + 30)
                        self.click(init_click_point) 
                        if self.flag_error[index] == 1:
                            self.double_click(self.point_num)
                            phoneNum = self.phone_numbers[index][self.deal_number[index]]
                            while self.check_result[str(phoneNum)] != -1:
                                self.deal_number[index] += 1
                                phoneNum = self.phone_numbers[index][self.deal_number[index]]
                            keyboard.write(str(phoneNum))
                            
                            # print(str(phoneNum), '      ', len(str(phoneNum)))

                            phoneNum_list[index].append(str(phoneNum))                        
                        captcha_img = np.array(([255]))
                        while captcha_img.mean() == 255:
                            captcha_img = self.get_snapshot(self.point_img, 200, 50)
                        text = self.perform_ocr(captcha_img)
                        self.double_click(self.point_recog_text)
                        keyboard.write(text)
                        QApplication.processEvents()
                        time.sleep(8/csv_num)
                        self.click(self.point_btn)
                        QApplication.processEvents()
                        time.sleep(8/csv_num)
                        QApplication.processEvents()
            time.sleep(1)    
            for row in range(2):
                for col in range(2):
                    QApplication.processEvents()
                    if self.flag_active: 
                        index = row * 2 + col
                        if self.file_path[index] == '':
                            continue 
                        x = col * self.colStep 
                        y = row * self.rowStep
                        self.point_error_text = (x+266, y+216)                        
                        error_img = self.get_snapshot(self.point_error_text, 350, 37)
                        error_text = pytesseract.image_to_string(error_img, lang='jpn', config='--psm 6').strip()
                        # print(error_text)
                        tested_phone = self.phone_numbers[index][self.deal_number[index]]
                        if "ご指定の電話番号による検索はできませんでした。" in error_text:
                            self.check_result[tested_phone] = 0  # failed                            
                            self.deal_number[index] += 1
                            # print(str(self.deal_number)+"件 処理")
                            self.flag_error[index] = 1
                        elif "入力された画像" in error_text:
                            self.flag_error[index] = 0
                        else:
                            self.flag_error[index] = 1
                            self.check_result[tested_phone] = 1
                        print(self.check_result[tested_phone], '    ', tested_phone)
                        print("================================================")

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    # browser.simulate()
    MainWindow = QtWidgets.QMainWindow()
    MainWindow.setWindowIcon(QIcon("icon.png"))
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    # hwnd = ctypes.windll.user32.GetForegroundWindow()
    hwnd = ctypes.windll.user32.FindWindowW(None, "NTT フレッツ光判定システム")
    win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
    sys.exit(app.exec_())