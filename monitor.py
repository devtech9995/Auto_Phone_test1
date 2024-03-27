import psutil
from selenium import webdriver

driver = webdriver.Chrome()
driver.quit()

def monitor_browser_process(browser_name):
    for proc in psutil.process_iter(['pid', 'name']):
        print(proc.info['name'])
        print(proc.info[''])
        if proc.info['name'] == browser_name:
            return True
    return False

# Specify the name of the browser executable
browser_name = 'chrome.exe'

# Monitor the browser process
while monitor_browser_process(browser_name):
    pass

print(f"The {browser_name} process has been closed")
