import cv2
import numpy as np
import pytesseract

replaceDir = {
    'o':'p', 'B':'8', 'X':'x', 'C':'c', '?':'2', 'Z':'2', 'z':'2', 'i':'7', 'S':'8', 
    '/':'7', 'é':'6', 'j':'3', ' ':'', 'W':'w', '&':'8', '“':'', 'J':'2', 's':'3', 
    "D":'b', '*':'f', '.':'', '£':'f', 'O':'o', 'P':'p', 'V':'v', 'r':'w', '`':'', 
    '_':'', 't':'f', '-':'', 't':'f', 'h':'b', 'A':'4', 'T':'7', '9':'g', 'q':'g', 
    ',':'', 'l':'7', '0':'n', '{':'f', '(':'f', 'I':'2', '$':'5', '^':'' , '1':'7', 
    '¥':'y', '[':'7', 'N':'n', 'u':'w', 'F':'f', 'G':'g', '#':'f', 'E':'6', 'k':'x', 
    'K':'x', '‘':'', 'Y':'y', '§':'5', '¢':'c', '€':'c', '%':'x'
}


class Image_ocr:
    def __init__(self, file_path):
        self.file_path = file_path
        # self.perform_ocr(file_path)

    def filterText(captcha):
        result = ''      
        for each in captcha:
            if each in replaceDir:
                result += replaceDir[each]
            else:
                result += each
        return result
    
    def perform_ocr(self):
        # Perform OCR on the image and extract the text
        img_array = np.fromfile(self.file_path, np.uint8) 
        img_file = cv2.imdecode(img_array,  cv2.IMREAD_COLOR)
        img = cv2.cvtColor(img_file, cv2.COLOR_BGR2GRAY)        
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

        return output_file_path