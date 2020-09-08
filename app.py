
from selenium import webdriver
import selenium
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
from selenium.webdriver import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import random
import time,sys,os,pyautogui
import pandas as pd
from sys import platform
cur_path = sys.path[0]
pyautogui.FAILSAFE = False
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)

if platform == "linux" or platform == "linux2":
    # linux
    path = resource_path('driver/chromedriver')
else:
    path = resource_path('driver/chromedriver.exe')
    # Windows...

def asking():
    prompt = input("you want to continue(y/n)? ")
    if(prompt == "y" ):
        pass
    else:
        print("asking again...")
        asking()
def toForm():    
    while(True):
        try:
            time.sleep(10)
            driver.find_elements_by_css_selector("button.float-right")[0].click()
            print("success : clicked on open project")
            time.sleep(8)
            driver.get("https://beta.paydaydocinc.com/#/addForm")
            print("success: clicked add form")
            break
        except Exception as e:
            print(e)
            print("something went wrong...refreshing tha page")
            driver.refresh()
        
print("\n\nProcessing.....")

driver =webdriver.Chrome(path)
# open link
# driver.set_page_load_timeout(120)

try:
    driver.get("https://beta.paydaydocinc.com/")
    print("success : Loaded...")
except TimeoutException as e:
    print("info : website taking too long to load...stopped")
    driver.refresh()

# driver.find_element_by_xpath('//form/div[0]/input[@class="form-control"]').send_keys('KHI3873')
# driver.find_element_by_xpath('//form/div[1]/input[@class="form-control"]').send_keys("2urbawrd")
while True:
    try:
        inputs = driver.find_elements_by_css_selector("input.form-control")

        d = ["email","password"]
        inputs[0].send_keys(d[0])
        inputs[1].send_keys(d[1])

        driver.find_elements_by_tag_name("button")[0].click()
        print("success : logging in...")
        break
    except Exception as e:
        print(e)
        print("error : something went wrong... refreshing")
        driver.refresh()

toForm()
time.sleep(8)
# extracting data
filename = "data.xlsx"
df = pd.read_excel(filename)


# badge = driver.find_elements_by_xpath("//span[@class='badge']/font")
# finding completed forms number
badge = driver.find_element_by_tag_name('font').text

badge = int(badge) + 1
# writing current form #
while(True):
    try: 
        time.sleep(5)
        driver.find_element_by_id('appendedInputButton').send_keys(str(badge))
        print("success : typed '" + str(badge)+"'")
        # driver.find_element_by_xpath("//button[text()='Go!' and @type='button']").click()
        btns = driver.find_element_by_css_selector('button.btn-secondary').click()
        print("success : clicking Go!")
        break
    except Exception as e:
        print(e)
        pyautogui.hotkey(['win','up'])
        print("error : something went wrong ... refreshing page")
        driver.refresh()
        time.sleep(5)
        toForm()



form_num = df['form num'].tolist()
com_code = df['company code'].tolist()
com_name = df['Company Name'].tolist()
com_add = df['Company Address'].tolist()
zip_code = df['Zip Code'].tolist()
fax = df['Fax'].tolist()
web = df['Website'].tolist()
email = df['Email'].tolist()
contact = df['Contact'].tolist()
state = df['State'].tolist()
country = df['Country'].tolist()
head_qua = df['Head Quarter'].tolist()
num_of_emp = df['Number of Employees'].tolist()
ind = df['Industry'].tolist()
brand_amb = df['Brand Ambassador'].tolist()
med_part = df['Media Partner'].tolist()
soc_med = df['Social Media'].tolist()
fra_part = df['Franchise Partner'].tolist()
inv = df['Investor'].tolist()
ad_med = df['Advertising Media'].tolist()
product = df['Product'].tolist()
services = df['Services'].tolist()
manager = df['Manager'].tolist()
sub_class = df['Sub Classification'].tolist()
reg_date = df['Registration date'].tolist()
year_rev = df['Yearly Revenue'].tolist()
land_mark = df['Land Mark'].tolist()
acc_aud = df['Account Audit'].tolist()
Currency = df['Currency'].tolist()
year_exp = df['Yearly Expense'].tolist()








# i = 0
# while i<len(form_num)-1:
#     driver.find_element_by_id("Form Number").send_keys(form_num[i])
#     driver.find_element_by_id("Company Code").send_keys(com_code[i])
#     driver.find_element_by_id("Company Name").send_keys(com_name[i])
#     asking()
#     i += 1


print("success : complete")


