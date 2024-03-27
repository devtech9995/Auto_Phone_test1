import cv2
import numpy as np
import pytesseract
import pyautogui
from PyQt5.QtWidgets import QFileDialog, QMessageBox
from PyQt5 import QtWidgets, QtCore  
from PyQt5.QtGui import QFont  


replaceDir = {
    'o':'p', 'B':'8', 'X':'x', 'C':'c', '?':'2', 'Z':'2', 'z':'2', 'i':'7', 'S':'8', 
    '/':'7', 'é':'6', 'j':'3', ' ':'', 'W':'w', '&':'8', '“':'', 'J':'2', 's':'3', 
    "D":'b', '*':'f', '.':'', '£':'f', 'O':'o', 'P':'p', 'V':'v', 'r':'w', '`':'', 
    '_':'', 't':'f', '-':'', 't':'f', 'h':'b', 'A':'4', 'T':'7', '9':'g', 'q':'g', 
    ',':'', 'l':'7', '0':'n', '{':'f', '(':'f', 'I':'2', '$':'5', '^':'' , '1':'7', 
    '¥':'y', '[':'7', 'N':'n', 'u':'w', 'F':'f', 'G':'g', '#':'f', 'E':'6', 'k':'x', 
    'K':'x', '‘':'', 'Y':'y', '§':'5', '¢':'c', '€':'c', '%':'x'
}

def filterText(captcha):
    result = ''      
    for each in captcha:
        if each in replaceDir:
            result += replaceDir[each]
        else:
            result += each
    return result

def perform_ocr(img):
    img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)        
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
    text = filterText(pytesseract.image_to_string(eroded).strip())
    if text == '' :
        text = 'error' 
    return text

def execute(opt):
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    if opt == 1:
        file, _ = QFileDialog.getOpenFileName(MainWindow, 'Open File', '', 'Image Files (*.bmp *.jpg *.png)')
        img_array = np.fromfile(file, np.uint8) 
        img_file = cv2.imdecode(img_array,  cv2.IMREAD_COLOR)
        # img_file = cv2.imread(file)
    else:
        print("Input a coordinate of left edge in recognizing image")
        x = float(input("Enter the value of x: "))
        y = float(input("Enter the value of y: "))
        snapshot = pyautogui.screenshot(region=(x, y, 200, 50))
        img_file = np.array(snapshot)
    recog_text = perform_ocr(img_file)
    from PyQt5.QtWidgets import QMessageBox
    msg_box = QMessageBox()
    msg_box.setWindowTitle("result")
    msg_box.setText(recog_text)
    msg_box.setFont(QFont('Times', 12))
    msg_box.setWindowFlags(msg_box.windowFlags() | QtCore.Qt.WindowStaysOnTopHint)
    msg_box.exec_()
    print(recog_text)
        
