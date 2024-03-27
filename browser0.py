from selenium import webdriver
import pyautogui
import keyboard
import time

def click(x, y, t): 
    for row in range(2):
        for col in range(2):
            click_x = x + row*960
            click_y = y + col*520
            pyautogui.click((click_x, click_y))
            time.sleep(t)
            
def simulate():
    url = "https://www.ntt-east.co.jp/line-info/consent.html"
    ieOptions = webdriver.IeOptions()
    ieOptions.add_additional_option("ie.edgechromium", True)
    ieOptions.add_additional_option("ie.edgepath",'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe')
    zoom_level = 0.75

    driver1 = webdriver.Ie(options=ieOptions)
    driver1.set_window_size(height=520, width=960)
    driver1.set_window_position(x=0, y=0)
    driver1.execute_script(f"document.body.style.zoom = '{zoom_level}';")

    driver2 = webdriver.Ie(options=ieOptions)
    driver2.set_window_size(height=520, width=960)
    driver2.set_window_position(x=960, y=0)

    driver3 = webdriver.Ie(options=ieOptions)
    driver3.set_window_size(height=520, width=960)
    driver3.set_window_position(x=0, y=520)

    driver4 = webdriver.Ie(options=ieOptions)
    driver4.set_window_size(height=520, width=960)
    driver4.set_window_position(x=960, y=520)


    for row in range(2):
        for col in range(2):
            x = 300 + row*960
            y = 60 + col*520
            pyautogui.click((x,y))
            time.sleep(1)
            keyboard.write(url)
            time.sleep(1)
            keyboard.press('Enter')
            time.sleep(1)
            
    time.sleep(25)

    click(890,450,0.2)  # scroll down
    click(890,450,0.2)  # scroll down

    click(30,350,0)     # next button click
    time.sleep(25)

    click(870,100,0)    # X click
    click(150,340,15)   # login button click
    
    time.sleep(10)
    
    for row in range(2):
        for col in range(2):
            x = 400 + row*960
            y = 400 + col*520
            pyautogui.click((x,y))
            time.sleep(1)
            pyautogui.hold('ctr')
            pyautogui.scroll(-30)
            time.sleep(1)
            # keyboard.press('Enter')
            # time.sleep(1)


            
# exit(0)

# # driver.get('https://lios-web.ipd.ntt-east.co.jp/LiosApp1/LoginPub')
# from selenium import webdriver
# from selenium.webdriver.common.by import By  
# from selenium.webdriver.common.keys import Keys
# import pyautogui
# import keyboard
# import time

# def click(x, y, t): 
#     for row in range(2):
#         for col in range(2):
#             click_x = x + row*960
#             click_y = y + col*520
#             pyautogui.click((click_x, click_y))
#             time.sleep(t)
            

# url = "https://www.ntt-east.co.jp/line-info/consent.html"
# ieOptions = webdriver.IeOptions()
# ieOptions.add_additional_option("ie.edgechromium", True)
# ieOptions.add_additional_option("ie.edgepath",'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe')
# zoom_level = 0.75

# driver1 = webdriver.Ie(options=ieOptions)
# driver1.set_window_size(height=520, width=960)
# driver1.set_window_position(x=0, y=0)

# driver2 = webdriver.Ie(options=ieOptions)
# driver2.set_window_size(height=520, width=960)
# driver2.set_window_position(x=960, y=0)

# driver3 = webdriver.Ie(options=ieOptions)
# driver3.set_window_size(height=520, width=960)
# driver3.set_window_position(x=0, y=520)

# driver4 = webdriver.Ie(options=ieOptions)
# driver4.set_window_size(height=520, width=960)
# driver4.set_window_position(x=960, y=520)

# for row in range(2):
#     for col in range(2):
#         x = 300 + row*960
#         y = 90 + col*520
#         pyautogui.click((x,y))
#         time.sleep(0.2)
#         keyboard.write(url)
#         time.sleep(0.1)
#         keyboard.press('Enter')
#         time.sleep(0.1)
        
# time.sleep(2)

# click(870,450,0.01)  # scroll down
# click(870,450,0.01)  
# click(870,450,0.01)  

# click(40,380,0)     # next button click
# time.sleep(2)

# click(840,140,0)    # X click
# time.sleep(1)
# click(200,450,15)   # login button click

# for row in range(2):
#     for col in range(2):
#         x = 500 + row*960
#         y = 400 + col*520
#         pyautogui.click((x,y))
#         time.sleep(1)
#         pyautogui.scroll(-30)
#         pyautogui.keyDown('ctrl')
#         pyautogui.scroll(-20)
#         pyautogui.scroll(-20)
#         pyautogui.keyUp('ctrl')
#         time.sleep(1)
