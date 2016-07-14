from selenium import webdriver
from pywinauto import application
from pywinauto.application import Application
import datetime, calendar, re, time 
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
elements = []
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
    #Maybe ask to input pin to input?
    #Set focus to pin element
    #Handle if pressed "continue" or something, twice
    if(len(driver.find_elements_by_class_name("success"))>0):
        driver.execute_script('$("a.btn-default")[0].click()')
        return True
    else:
        driver.find_element_by_class_name("send-me-pin").click()
        enter_function()
        log("Didn't find success text, abort")
    #Any way to verificate automatically?
def test():
    name = "Anssi Tenhunen"
    arrival = "140716"
    departure = "170716"
    total = "249"
    comment = "Tällä testataan toimiiko ohjelma"
    yritys = "Tenhusen syyhysormi"
    kontakti_0 = "kontakti0"
    kontakti_1 = "kontakti_1"
    puhelin = "+358 45 898 2272"
    email = "anssi.tenhunen@gmail.com"
    viite = "viite"
    agentti_yritys = "yritys"
    agentti_id = "id"
    agentti_kontakti_0 = "kontakti_0"
    agentti_kontakti_1 = "kontakti_1"
    agentti_puhelin = "puhelin"
    agentti_laskuta = False
    hintaryhma = "CN"
    segmentti = ""
    info = comment
    tyyppi = "D"
    saapuu = arrival
    lahtee = departure
    huone_maara = "1"
    aikuisia = "2"
    lapsia = "1"
    lisat = "3"
    vapaat = "4"
    tuloviesti = "Hola hola!"
    kampanja = ""
    ryhma_id = ""
    hinta = "D"
    ennakko = ""
    varausreitti = "b"
    virkailija = ""
    summa = total
    app = Application().Connect(title=u'FrontOffice I-S - Kioski vastaanotto', class_name='ThunderRT6FormDC')
    thunderrtformdc = app.ThunderRT6FormDC
    bring_front(thunderrtformdc)
    thunderrtformdc.TypeKeys("{F3}")
    time.sleep(1)
    #Thos f*cking elements keep changing their id's, we have to do this with pure keyboard input
    thunderrtformdc.TypeKeys(name.split(" ")[1]+"{TAB}"+name.split(" ")[0]+"{TAB}"+ yritys+"{TAB}"+kontakti_0+"{TAB}"+kontakti_1+"{TAB}"+puhelin+"{TAB}"+"{TAB}"+email+"{TAB}"+"{TAB}"+viite+"{TAB}"+"{TAB}"+agentti_yritys+"{TAB}"+agentti_id+"{TAB}"+agentti_kontakti_0+"{TAB}"+agentti_kontakti_1+"{TAB}"+agentti_puhelin+"{TAB}"+"{TAB}"+"{TAB}"+hintaryhma+"{TAB}"+segmentti+"{TAB}"+info+"{TAB}"+"{TAB}"+tyyppi+"{TAB}"+saapuu+"{TAB}"+"{TAB}"+"{TAB}"+"{TAB}"+lahtee+"{TAB}"+"{BACKSPACE}"*3+huone_maara+"{TAB}"+"{TAB}"+"{BACKSPACE}"*3+aikuisia+"{TAB}"+"{BACKSPACE}"*3+lapsia+"{TAB}"+"{TAB}"+lisat+"{TAB}"+vapaat+"{TAB}"+tuloviesti+"{TAB}"+kampanja+"{TAB}"+"{TAB}"+"{TAB}"+"{TAB}"+ryhma_id+"{TAB}"+"{TAB}"+hinta+"{TAB}"+"{TAB}"+"{TAB}"+"{TAB}"+"{TAB}"+"{TAB}"+"{TAB}"+ennakko+"{TAB}"+"{TAB}"+"{TAB}"+"{TAB}"+varausreitti+"{TAB}"+virkailija,with_spaces=True)
    time.sleep(1)
    input("Done writing, is everything correct?")
    
def loop_bookings():
    if(str(driver.title) != "Bookings · Booking.com"):
        log("incorrect page, it's " + str(driver.title))
        input("Press enter to continue: ")
        enter_function()
        return False
    bookings = driver.execute_script("return $('tbody>tr')")
    links = driver.execute_script("return $('tr>td.align_right>a')")
    index = -1
    for i in bookings:
        index +=1
        #Check "Saapuvat" radio button
        values = i.find_elements_by_tag_name("td")
        name = values[4].text
        arrival = parse_date(values[6].text)
        departure = parse_date(values[7].text)
        total = str(re.search('([0-9]+)',values[9].text).group(0))
        comment = values[3].text
        yritys = ""
        kontakti_0 = ""
        kontakti_1 = ""
        puhelin = ""
        email = ""
        viite = ""
        agentti_yritys = ""
        agentti_id = ""
        agentti_kontakti_0 = ""
        agentti_kontakti_1 = ""
        agentti_puhelin = ""
        agentti_laskuta = False
        hintaryhma = "c"+""
        segmentti = ""
        info = comment
        tyyppi = ""
        saapuu = arrival
        lahtee = departure
        huone_maara = "1"
        aikuisia = "1"
        lapsia = "0"
        lisat = "0"
        vapaat = "0"
        tuloviesti = ""
        kampanja = ""
        ryhma_id = ""
        hinta = ""
        ennakko = ""
        varausreitti = "b"
        virkailija = ""
        summa = total
        log(name + " " + arrival + " " + departure)
        if(search_reservation(name,arrival,departure)):
            log("reservation for "+name+" on "+arrival + " exists")
            continue
        else:
            """
            TODO:
            Get credit card details
            find out room type
            enter details to reservation
            """
            log("Make reservation for " + name + " on " + arrival)
            log("Open reservation at Booking")
            links[index].click()
            solve()
            puhelin = driver.execute_script('return $(".phone-info")[0].innerText').replace(" ","")
            tyyppi = driver.execute_script('return $("tbody>tr:eq(4)>td:eq(1)").text()').strip()
            # TODO:
            # It seems like these are not correct, some of them change  (seems to bee only u'*' elements
            # end of title after "-" changes
            app = Application().Connect(title=u'FrontOffice I-S - Hotelli', class_name='ThunderRT6FormDC')
            thunderrtformdc = app.ThunderRT6FormDC
            thunderrtformdc.TypeKeys("{F3}")
            time.sleep(1)
            #Thos f*cking elements keep changing their id's, we have to do this with pure keyboard input
            thunderrtformdc.TypeKeys(name.split(" ")[1]+"{TAB}"+name.split(" ")[0]+"{TAB}"+ yritys+"{TAB}"+kontakti_0+"{TAB}"+kontakti_1+"{TAB}"+puhelin+"{TAB}"+"{TAB}"+email+"{TAB}"+"{TAB}"+viite+"{TAB}"+"{TAB}"+agentti_yritys+"{TAB}"+agentti_id+"{TAB}"+agentti_kontakti_0+"{TAB}"+agentti_kontakti_1+"{TAB}"+agentti_puhelin+"{TAB}"+"{TAB}"+"{TAB}"+hintaryhma+"{TAB}"+segmentti+"{TAB}"+info+"{TAB}"+"{TAB}"+tyyppi+"{TAB}"+saapuu+"{TAB}"+"{TAB}"+"{TAB}"+"{TAB}"+lahtee+"{TAB}"+"{BACKSPACE}"*3+huone_maara+"{TAB}"+"{TAB}"+"{BACKSPACE}"*3+aikuisia+"{TAB}"+"{BACKSPACE}"*3+lapsia+"{TAB}"+"{TAB}"+lisat+"{TAB}"+vapaat+"{TAB}"+tuloviesti+"{TAB}"+kampanja+"{TAB}"+"{TAB}"+"{TAB}"+"{TAB}"+ryhma_id+"{TAB}"+"{TAB}"+hinta+"{TAB}"+"{TAB}"+"{TAB}"+"{TAB}"+"{TAB}"+"{TAB}"+"{TAB}"+ennakko+"{TAB}"+"{TAB}"+"{TAB}"+"{TAB}"+varausreitti+"{TAB}"+virkailija,with_spaces=True)
            time.sleep(1)
            if(input("Done writing, is everything correct? (n or blank")=="n"):
                break
            thunderrtformdc.TypeKeys("{F4}")
            """
            saapuu_el.TypeKeys(arrival)
            saapuu_el.DrawOutline()
            input(arrival)
            lahtee_el.TypeKeys(departure)
            lahtee_el.DrawOutline()
            input(departure)
            info_el.TypeKeys(comment)
            info_el.DrawOutline()
            input()
            puhelin_el.TypeKeys(number)
            _el.DrawOutline()
            input()
            
            """
            go_to_bookings()
            return False
        """
        TODO:
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
    return str(re.search('(\d{2}(?= ))',txt).group(0))+str(months[re.search('[A-z]+(?!,)(?= )',txt).group(0)]).zfill(2)+str(re.search('(\d{4})',txt).group(0))

def end():
    f.close()
    driver.quit()

def log(txt):
    f.write(str(txt) + ' - {:%d.%m.%Y %H:%M:%S}'.format(datetime.datetime.now()) + "\n")
    print(str(txt)+ ' - {:%H:%M:%S}'.format(datetime.datetime.now()))
    
def search_reservation(name="",arrival="",departure=""):
    app = Application().Connect(title=u'FrontOffice I-S - Hotelli', class_name='ThunderRT6FormDC')
    thunderrtformdc = app.ThunderRT6FormDC
    bring_front(thunderrtformdc)
    menu_item = thunderrtformdc.MenuItem(u'Etsi->&Varaukset...\tF7')
    menu_item.Click()
    thunderrtformdc2 = app.Varaukset
    #click "saapuvat"
    thunderrttextbox = thunderrtformdc2[u'Edit4']
    for i in range(2):
        thunderrttextbox.ClickInput()
        thunderrttextbox.TypeKeys("{BACKSPACE}"*25)
        thunderrttextbox.TypeKeys(name.split(" ")[i])
        edit = thunderrtformdc2[u'Edit3']
        edit.ClickInput()
        edit.TypeKeys(arrival)
        thunderrtcommandbutton = thunderrtformdc2[u'H&ae']
        thunderrtcommandbutton.Click()
        time.sleep(1)
        listviewwndclass = thunderrtformdc.Children()[15]
        if(listviewwndclass.ItemCount()>0):
            thunderrtformdc.TypeKeys("{ESC}")
            return True
        else:
            continue
    thunderrtformdc.TypeKeys("{ESC}")
    return False

def bring_front(window):
    window.Minimize()
    window.Maximize()
    window.SetFocus()
if(login()):
    go_to_bookings()
    loop_bookings()
end()

