# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'Mainwindow.ui'
#
# Created by: PyQt5 UI code generator 5.15.6
#
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
from PyQt5.QtWidgets import QFileDialog, QMainWindow, QApplication,  
import pickle
from selenium import webdriver
import browser0

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
    rowStep = 515
    colStep = 960
    file_path = [''] * 4
    
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
                    phone_numbers_in_row = re.findall(r'\d{3}[-\.\s]??\d{3}[-\.\s]??\d{4}', row[0])
                    # Remove hyphens and spaces from phone numbers
                    phone_numbers_in_row = [number.replace('-', '').replace(' ', '') for number in phone_numbers_in_row]
                    phone_number.extend(phone_numbers_in_row)
                    # tmp = str(row[0].replace('-', ''))     
                    # phone_number.append(tmp)
            return phone_number
    
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(570, 237)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.splitter = QtWidgets.QSplitter(self.centralwidget)
        self.splitter.setGeometry(QtCore.QRect(30, 20, 500, 181))
        self.splitter.setOrientation(QtCore.Qt.Vertical)
        self.splitter.setObjectName("splitter")
        self.widget = QtWidgets.QWidget(self.splitter)
        self.widget.setObjectName("widget")
        self.horizontalLayout = QtWidgets.QHBoxLayout(self.widget)
        self.horizontalLayout.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.lineEdit = QtWidgets.QLineEdit(self.widget)
        self.lineEdit.setMinimumSize(QtCore.QSize(0, 28))
        self.lineEdit.setMaximumSize(QtCore.QSize(16777215, 28))
        self.lineEdit.setObjectName("lineEdit")
        self.horizontalLayout.addWidget(self.lineEdit)
        self.add_csv = QtWidgets.QToolButton(self.widget)
        self.add_csv.setMinimumSize(QtCore.QSize(32, 28))
        self.add_csv.setMaximumSize(QtCore.QSize(32, 28))
        self.add_csv.setObjectName("add_csv")
        self.add_csv.clicked.connect(self.open_file)
        self.horizontalLayout.addWidget(self.add_csv)
        self.widget1 = QtWidgets.QWidget(self.splitter)
        self.widget1.setObjectName("widget1")
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout(self.widget1)
        self.horizontalLayout_2.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.lineEdit_2 = QtWidgets.QLineEdit(self.widget1)
        self.lineEdit_2.setMinimumSize(QtCore.QSize(0, 28))
        self.lineEdit_2.setMaximumSize(QtCore.QSize(16777215, 28))
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.horizontalLayout_2.addWidget(self.lineEdit_2)
        self.add_csv_2 = QtWidgets.QToolButton(self.widget1)
        self.add_csv_2.setMinimumSize(QtCore.QSize(32, 28))
        self.add_csv_2.setMaximumSize(QtCore.QSize(32, 28))
        self.add_csv_2.setObjectName("add_csv_2")
        self.add_csv_2.clicked.connect(self.open_file)
        self.horizontalLayout_2.addWidget(self.add_csv_2)
        self.widget2 = QtWidgets.QWidget(self.splitter)
        self.widget2.setObjectName("widget2")
        self.horizontalLayout_3 = QtWidgets.QHBoxLayout(self.widget2)
        self.horizontalLayout_3.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        self.lineEdit_3 = QtWidgets.QLineEdit(self.widget2)
        self.lineEdit_3.setMinimumSize(QtCore.QSize(0, 28))
        self.lineEdit_3.setMaximumSize(QtCore.QSize(16777215, 28))
        self.lineEdit_3.setObjectName("lineEdit_3")
        self.horizontalLayout_3.addWidget(self.lineEdit_3)
        self.add_csv_3 = QtWidgets.QToolButton(self.widget2)
        self.add_csv_3.setMinimumSize(QtCore.QSize(32, 28))
        self.add_csv_3.setMaximumSize(QtCore.QSize(32, 28))
        self.add_csv_3.setObjectName("add_csv_3")
        self.add_csv_3.clicked.connect(self.open_file)
        self.horizontalLayout_3.addWidget(self.add_csv_3)
        self.widget3 = QtWidgets.QWidget(self.splitter)
        self.widget3.setObjectName("widget3")
        self.horizontalLayout_4 = QtWidgets.QHBoxLayout(self.widget3)
        self.horizontalLayout_4.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_4.setObjectName("horizontalLayout_4")
        self.lineEdit_4 = QtWidgets.QLineEdit(self.widget3)
        self.lineEdit_4.setMinimumSize(QtCore.QSize(0, 28))
        self.lineEdit_4.setMaximumSize(QtCore.QSize(16777215, 28))
        self.lineEdit_4.setObjectName("lineEdit_4")
        self.horizontalLayout_4.addWidget(self.lineEdit_4)
        self.add_csv_4 = QtWidgets.QToolButton(self.widget3)
        self.add_csv_4.setMinimumSize(QtCore.QSize(32, 28))
        self.add_csv_4.setMaximumSize(QtCore.QSize(32, 28))
        self.add_csv_4.setObjectName("add_csv_4")
        self.add_csv_4.clicked.connect(self.open_file)
        self.horizontalLayout_4.addWidget(self.add_csv_4)
        self.widget4 = QtWidgets.QWidget(self.splitter)
        self.widget4.setObjectName("widget4")
        self.horizontalLayout_5 = QtWidgets.QHBoxLayout(self.widget4)
        self.horizontalLayout_5.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_5.setObjectName("horizontalLayout_5")
        self.start_emul = QtWidgets.QPushButton(self.widget4)
        self.start_emul.setMinimumSize(QtCore.QSize(120, 32))
        self.start_emul.setFont(QFont('Times', 14))
        self.start_emul.setObjectName("start_emul")
        self.start_emul.clicked.connect(self.execute)
        self.horizontalLayout_5.addWidget(self.start_emul)
        self.pauseEmul = QtWidgets.QPushButton(self.widget4)
        self.pauseEmul.setMinimumSize(QtCore.QSize(120, 32))
        self.pauseEmul.setFont(QFont('Times', 14))
        self.pauseEmul.setObjectName("pauseEmul")
        self.pauseEmul.clicked.connect(self.pause)
        
        self.horizontalLayout_5.addWidget(self.pauseEmul)
        self.clear_list = QtWidgets.QPushButton(self.widget4)
        self.clear_list.setMinimumSize(QtCore.QSize(120, 32))
        self.clear_list.setFont(QFont('Times', 14))
        self.clear_list.setObjectName("clear_list")
        self.clear_list.clicked.connect(self.clear_lists)
        self.horizontalLayout_5.addWidget(self.clear_list)
        self.clear_list_2 = QtWidgets.QPushButton(self.widget4)
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
        MainWindow.setWindowTitle(_translate("MainWindow", "NTT西日本 フレッツ光判定システム"))
        self.add_csv.setText(_translate("MainWindow", "..."))
        self.add_csv_2.setText(_translate("MainWindow", "..."))
        self.add_csv_3.setText(_translate("MainWindow", "..."))
        self.add_csv_4.setText(_translate("MainWindow", "..."))
        self.start_emul.setText(_translate("MainWindow", "開  始"))
        self.pauseEmul.setText(_translate("MainWindow", "停  止"))
        self.clear_list.setText(_translate("MainWindow", "削  除"))
        self.clear_list_2.setText(_translate("MainWindow", "保  管"))
        
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
        save_pickle(self.check_result)
        filenames = ''
        for each in self.file_path:
            if len(each) > 0:
                current_datetime = datetime.datetime.now()
                formatted_datetime = current_datetime.strftime("%H-%M-%S_%Y%m%d")
                tmpresult = filenames.join("result_" + formatted_datetime + "_" + each.split("/")[-1])    
                print(tmpresult)     
                with open(tmpresult, mode='w', newline='', encoding='latin-1', errors="ignore") as file:    
                    reader = csv.reader(open(each, mode='r', encoding='latin-1', errors="ignore"))
                    writer = csv.writer(file)
                    data = load_pickle("filename.pickle")

                    keys = data.keys()
                    for row in reader:
                        temp_key = str(row[0]).replace('-', '')
                        if temp_key in keys:                    
                            status = "todo"
                            if data[temp_key] == 0:
                                status = "failed"
                            if data[temp_key] == 1:
                                status = "success"
                            row.insert(1, status)
                            writer.writerow(row)
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
        # print(self.phone_numbers[index])
        for phone in self.phone_numbers[index]:            
            self.check_result[str(phone)] = -1
        # print(self.check_result[index])
                
    def execute(self):
        # Decide to load pickle file  0:no, 1:load
        config = configparser.ConfigParser()
        config.read('setting.ini')
        # print(config)
        if config['main']['key'] == '0':
            print("config creating...")
            data = {}
            save_pickle(data)
        else:
            print("config loading")
            self.check_result = load_pickle('filename.pickle')
            
        # print(self.file_path)
        csv_num = self.calculate_csv_num()
                       
        self.add_csv.setEnabled(False)
        self.clear_list.setEnabled(False)
        self.clear_list_2.setEnabled(False)
        self.start_emul.setEnabled(False)

        phoneNum_list = [[], [], [], []]
        
        self.click((400,400))
        while self.flag_active:
        # while True:
            for row in range(2):
                for col in range(2):
                    QApplication.processEvents()
                    window_title = self.monitor()
                    print(window_title)
                    if "線路情報開示画面" in window_title and self.flag_active:
                        index = row * 2 + col
                        if self.file_path[index] == '':
                            continue                           
                        print(index, window_title, self.flag_active)                        
                        x = col * self.colStep 
                        y = row * self.rowStep
                        self.point_num = (x+170, y+260)
                        self.point_img = (x+16, y+294)
                        self.point_recog_text = (x+100, y+420)
                        self.point_btn = (x+70, y+470)
                        
                        init_click_point = (x + 400, y + 400)
                        self.click(init_click_point) 
                        # print(self.flag_error[index])               
                        if self.flag_error[index] == 1:
                            self.double_click(self.point_num)
                            phoneNum = self.phone_numbers[index][self.deal_number[index]]
                            # print(phoneNum)
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
                        print("================================================")
                        QApplication.processEvents()
                
            for row in range(2):
                for col in range(2):
                    QApplication.processEvents()
                    if self.flag_active: 
                            index = row * 2 + col
                            if self.file_path[index] == '':
                                continue 
                            x = col * self.colStep 
                            y = row * self.rowStep
                            self.point_error_text = (x+262, y+192)                        
                            error_img = self.get_snapshot(self.point_error_text, 347, 40)
                            error_text = pytesseract.image_to_string(error_img, lang='jpn', config='--psm 6').strip()
                            tested_phone = self.phone_numbers[index][self.deal_number[index]]
                            if "ご指定の電話番号による検索はできませんでした。" in error_text:
                                self.check_result[tested_phone] = 0  # failed                            
                                
                                self.deal_number[index] += 1
                                # print(str(self.deal_number)+"件 処理")
                                self.flag_error[index] = 1
                            elif "入力された画像認証" in error_text:                        
                                self.flag_error[index] = 0
                            else:
                                self.check_result[tested_phone] = 1

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    MainWindow.setWindowIcon(QIcon("icon.png"))
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    hwnd = ctypes.windll.user32.GetForegroundWindow()
    win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
    sys.exit(app.exec_())
