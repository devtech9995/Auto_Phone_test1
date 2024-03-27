import os
import cv2
import numpy as np
import pytesseract       
image_folder = './all'
result_folder = "./result"
images = []
print(os.listdir(result_folder))

replaceDir = {
    'o':'p', 'B':'8', 'X':'x', 'C':'c', '?':'2', 'Z':'2', 'z':'2', 'i':'7', 'S':'8', 
    '/':'7', 'é':'6', 'j':'3', ' ':'', 'W':'w', '&':'8', '“':'', 'J':'2', 's':'3', 
    "D":'b', '*':'f', '.':'', '£':'f', 'O':'o', 'P':'p', 'V':'v', 'r':'w', '`':'', 
    '_':'', 't':'f', '-':'', 't':'f', 'h':'b', 'A':'4', 'T':'7', '9':'g', 'q':'g', 
    ',':'', 'l':'7', '0':'n', '{':'f', '(':'f', 'I':'2', '$':'5', '^':'' , '1':'7', 
    '¥':'y', '[':'7', 'N':'n', 'u':'w', 'F':'f', 'G':'g', '#':'f', 'E':'6', 'k':'x', 
    'K':'x', '‘':'', 'Y':'y', '§':'5', '¢':'c', '€':'c'
}
# replaceDir = {}
def fileterText(captcha):
    result = ''      
    for each in captcha:
        if each in replaceDir:
            result += replaceDir[each]
        else:
            result += each
    return result

for filename in os.listdir(result_folder):
        file_path = os.path.join(result_folder, filename)
        if os.path.isfile(file_path):
            os.remove(file_path)            

def removeStrikeLine(img):
    # img = img[8:, 52:170]
    # img[:, :52] = 255
    # img[:, 165:] = 255

    img[:, :52] = 255
    img[:, 165:] = 255

    img = cv2.GaussianBlur(img, (5, 3), 1)        

    _, img = cv2.threshold(img, 84, 255, cv2.THRESH_BINARY)
    bottom = img[43:]

    img = cv2.GaussianBlur(img, (3, 3), 1)        

    kernel = np.ones((2, 1), np.uint8)
    img = cv2.dilate(img, kernel, iterations=3)
            
    
    kernel = np.ones((2, 1), np.uint8)
    img = cv2.erode(img, kernel, iterations=1)
    img[43:, :] = bottom        

    # kernal = np.ones((1, 2), np.uint8)
    # img = cv2.erode(img, kernal, iterations=1)

    return img

for filename in os.listdir(image_folder):
    if filename.endswith('.jpg') or filename.endswith('.png') or filename.endswith('.bmp'):
        img_path = os.path.join(image_folder, filename)
        img = cv2.imread(img_path, cv2.IMREAD_GRAYSCALE)        
        tmp = img[49]
        for x in range(3):
            img = np.vstack((img, tmp)).astype(np.uint8)
        tmp = np.ones((2, img.shape[1]), np.uint8) * 255         
        img = np.vstack((img, tmp)).astype(np.uint8)
        res = removeStrikeLine(img)
        text = pytesseract.image_to_string(res).strip()        
        print(fileterText(text), len(text))
        cv2.imwrite('result/{}'.format( filename[:-4] + "_" + fileterText(text) + ".bmp"), res)                        
        # cv2.imwrite('result/{}'.format(filename), img)