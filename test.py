from selenium import webdriver
from pywinauto import application
import datetime, calendar, re
driver = 0

f = open("log.txt","w")
app = application.Application()
app.connect(path=r"C:\Program Files\Hotellinx Suite\Foflinx.exe")
months = dict((v,k) for k,v in enumerate(calendar.month_name))

#luuppaa varaukset läpi
#Lue varauksen tiedot verkosta
#Hae hotellinxista onko samanlainen varaus jo olemassa
#jos ei, tee
#jos on, seuraava varaus
#Kirjoita logi!

def login():
    global driver
    driver = webdriver.Chrome("C:/WinPython-32bit-3.5.1.3/python-3.5.1/Scripts/booking_varaus/chromedriver.exe")
    driver.get('https://admin.booking.com')
    #Some indicator for user to wait, it loads long before doing anything
    login_el = driver.find_element_by_id('loginname')
    login_el.send_keys("178223")
    password_el = driver.find_element_by_id('password')
    password_el.send_keys("Huippu246")
    password_el.submit()
    #Special cases (like "Unacknowleged Reservations....")
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

def go_to_bookings():
    tmp = "Bookings"
    driver.execute_script("return $('a:contains("+tmp+")')[0]").click()
    if(driver.title == "Bookings · Booking.com"):
        return True
    elif(driver.title == " · Booking.com"):
        log("Verification needed")
        #Send sms to right number
        #Any way to verificate automatically?
        #continue executing when number has been entered
    elif(driver.title == "Unacknowleged Reservations · Booking.com"):
        log("Unacknowleged Reservations")
        driver.find_element_by_class_name("btn").click()
        log("Clicked 'Acknowlege all'")
        return True
    else:
        log("Failed to enter bookings")
        return False

def loop_bookings():
    if(str(driver.title) != "Bookings · Booking.com"):
        log("incorrect page, it's " + str(driver.title))
        return False
    bookings = driver.execute_script("return $('tbody>tr')")
    for i in bookings:
        values = i.find_elements_by_tag_name("td")
        name = values[4].text
        arrival = parse_date(values[6].text)
        departure = parse_date(values[7].text)
        log(name + " " + arrival + " " + departure)
        #if(reservation_exists(name, arrival,departure)):
        """
        TODO:
        check if exists in hotellinx
        if doesn't:
            fill details
            open reservation
            get card data
            fill it under "Tilausm."
        """
        pass        

def reservation_exists(name, arrival, leave):
    return True
    
def parse_date(txt):
    return str(re.search('(\d{2}(?= ))',txt).group(0))+str(months[re.search('[A-z]+(?!,)(?= )',txt).group(0)])+str(re.search('(\d{4})',txt).group(0))

def end():
    f.close()
    driver.quit()

def log(txt):
    f.write(str(txt) + ' - {:%d.%m.%Y %H:%M:%S}'.format(datetime.datetime.now()) + "\n")
    print(str(txt)+ ' - {:%H:%M:%S}'.format(datetime.datetime.now()))
if(login()):
    go_to_bookings()
    loop_bookings()
end()
