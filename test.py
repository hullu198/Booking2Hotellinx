from selenium import webdriver
from pywinauto import application
import datetime
driver

f = open("log.txt")
app = application.Application()
app.connect(path=r"C:\Program Files\Hotellinx Suite\Foflinx.exe")

#luuppaa varaukset läpi
#Lue varauksen tiedot verkosta
#Hae hotellinxista onko samanlainen varaus jo olemassa
#jos ei, tee
#jos on, seuraava varaus
#Kirjoita logi!

def login():
    driver.quit()
    driver = webdriver.Chrome("C:/WinPython-32bit-3.5.1.3/python-3.5.1/Scripts/booking_varaus/chromedriver.exe")
    driver.get('https://admin.booking.com')
    login_el = driver.find_element_by_id('loginname')
    login_el.send_keys("178223")
    password_el = driver.find_element_by_id('password')
    password_el.send_keys("Huippu246")
    password_el.submit()
    if(driver.title=="Welcome to the Extranet · Booking.com"):
        log("Login succesful")
        return True
    else:
        log("Failed to login to Booking.com")
        return False

class wait_for_page_load(object):
    def __init__(self, browser):
        self.browser = browser
    def __enter__(self):
        self.old_page = self.browser.find_element_by_tag_name('html')
    def page_has_loaded(self):
        new_page = self.browser.find_element_by_tag_name('html')
        return new_page.id != self.old_page.id
    def __exit__(self, *_):
        wait_for(self.page_has_loaded)

def log(txt):
    f.write(str(txt) + ' - {:%d.%m.%Y %H:%M:%S}'.format(datetime.datetime.now()) + "\n")
    print(str(txt)+ ' - {:%H:%M:%S}'.format(datetime.datetime.now()))


driver.quit()
