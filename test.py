from selenium import webdriver
from pywinauto import application
from pywinauto.application import Application
import datetime, calendar, re
driver = 0

f = open("log.txt","w")
app = application.Application()
app.connect(path=r"C:\Program Files\Hotellinx Suite\Foflinx.exe")
months = dict((v,k) for k,v in enumerate(calendar.month_name))
titles = {'Booking · Booking.com': 'booking',
          'Booking.com Extranet': 'login',
          'Welcome to the Extranet · Booking.com': 'welcome',
          'Bookings · Booking.com': 'bookings',
          'Credit card details · Booking.com': 'card_details',
          '· Booking.com': 'verification',
          'Unacknowleged Reservations · Booking.com': 'unacknowleged'
          }
#luuppaa varaukset läpi
#Lue varauksen tiedot verkosta
#Hae hotellinxista onko samanlainen varaus jo olemassa
#jos ei, tee
#jos on, seuraava varaus
#Kirjoita logi!
#Maksutiedot OsoiteQ

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
    page = current_page()
    if(page=="welcome"):
        log("Login succesful")
        return True
    elif(page=="verification"):
        log("sms verification needed")
        verification()
    elif(page=="unacknowleged"):
        #click that button
        enter_function()
    else:
        log("Failed to login to Booking.com, current page's title is " + driver.title + "current_page() returned " + page)
        enter_function()

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
    elif(driver.title == "· Booking.com"):
        verification()
    elif(driver.title == "Unacknowleged Reservations · Booking.com"):
        log("Unacknowleged Reservations")
        driver.find_element_by_class_name("btn").click()
        log("Clicked 'Acknowlege all'")
        return True
    else:
        log("Failed to enter bookings")
        input("Press enter to continue: ")
        return False

def verification():
    if(driver.title != "· Booking.com"):
        return True
    driver.find_element_by_id("text_option").click()
    driver.execute_script("""$('select[name="phone_id_sms"] option:first-child').attr("selected", "selected");""")
    driver.find_elements_by_class_name("send-me-pin")[1].click()
    input("Press enter after you have entered pin correctly: ")
    if(len(driver.find_elements_by_class_name("success"))>0):
        driver.execute_script('$("a.btn-default")[0].click()')
        return True
    else:
        driver.find_element_by_class_name("send-me-pin").click()
        enter_function()
        log("Didn't find success text, abort")
    #Send sms to right number
    #Any way to verificate automatically?
    #continue executing when number has been entered

def loop_bookings():
    if(str(driver.title) != "Bookings · Booking.com"):
        log("incorrect page, it's " + str(driver.title))
        input("Press enter to continue: ")
        enter_function()
        return False
    bookings = driver.execute_script("return $('tbody>tr')")
    for i in bookings:
        values = i.find_elements_by_tag_name("td")
        name = values[4].text
        arrival = parse_date(values[6].text)
        departure = parse_date(values[7].text)
        log(name + " " + arrival + " " + departure)
        if(search_reservation(name,arrival,departure)):
            log("reservation for "+name+" on "+arrival + " exists")
        else:
            log("Make reservation for " + name + " on " + arrival)
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

def current_page():
    page = titles[driver.title]
    if(page == "verification"):
        if(len(driver.find_elements_by_class_name("enter-pin-div"))>0):
            return "enter_pin"
        elif(len(driver.find_elements_by_class_name("success"))>0):
            return "verified"
        else:
            return page
    else:
        return page
        
def enter_function():
    x = input("Please enter function name: ")
    while(True):
        if(x ==""):
            return False
            break
        try:
            globals()[x]()
            break
            return True
        except:
            x = input("Incorrect function, try again: " + str(sys.exc_info()[0]))

def solve():
    page= current_page()
    if(driver.title == "· Booking.com"):
        verification()
        solve()
    elif(page == "unacknowleged"):
        log("Unacknowleged Reservations")
        driver.find_element_by_class_name("btn").click()
        solve()
    else:
        log("Everything seems to be OK")
        return True

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
    
def search_reservation(name="",arrival="",departure=""):
    app = Application().Connect(title=u'FrontOffice I-S - Hotelli', class_name='ThunderRT6FormDC')
    thunderrtformdc = app.ThunderRT6FormDC
    menu_item = thunderrtformdc.MenuItem(u'Etsi->&Varaukset...\tF7')
    menu_item.Click()
    thunderrtformdc2 = app.Varaukset
    thunderrttextbox = thunderrtformdc2[u'Edit4']
    thunderrttextbox.ClickInput()
    thunderrttextbox.TypeKeys(name.split(" ")[0])
    edit = thunderrtformdc2[u'Edit3']
    edit.ClickInput()
    edit.TypeKeys(arrival)
    thunderrtcommandbutton = thunderrtformdc2[u'H&ae']
    thunderrtcommandbutton.Click()
    if(listviewwndclass.ItemCount()==0):
        return False
    else:
        return True
    
if(login()):
    go_to_bookings()
    loop_bookings()
end()


