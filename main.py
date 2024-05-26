import time
import multiprocessing
import pandas as pd
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import Select

df = pd.read_excel('input.xlsx', header = 1)
inputDict = df.to_dict()
totalNo = len(inputDict['S/NO'])

def run(row):
    time.sleep(10)
    fullName = inputDict['Name'][row]
    for sep in ['BINTE', 'BIN', 'BTE', '|']:
        if sep in fullName:
            firstName = f"{sep} {fullName.split(sep)[1].strip()}"
            lastName = f"{fullName.split(sep)[0].strip()}"
            break
    age = inputDict['Age'][row]
    nric = inputDict['NRIC'][row]
    dobObj = inputDict['D.O.B'][row]
    dobDD = str(dobObj.day).zfill(2)
    dobMM = str(dobObj.month).zfill(2)
    dobYY = str(dobObj.year)
    gender = inputDict['Gender'][row]
    email = 'REDACTED'
    mobileNo = inputDict['Contact No.'][row]
    passportNo = inputDict['Passport No. '][row] #space after PNo.

    planCapitalise = str(inputDict['PLAN'][row]).title()
    planTypes = ['Classic', 'Superior', 'Premier']
    for plan in planTypes:
        if plan in planCapitalise:
            planType = plan

    # input('start webdriver')
    driver = webdriver.Chrome()

    driver.get('https://online.aig.com.sg/AIGSingapore/apps/services/www/AIGSG/desktopbrowser/default/AIGSG.html#purchase-travel')

    #Fill in standard info (Zone 3, Prod Code, Sub Code, Staff Name)
    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//div[@data-target='#zone-3']"))).click()
    driver.find_element(By.XPATH, "//input[@placeholder='Producer code']").send_keys('REDACTED')
    driver.find_element(By.XPATH, "//input[@placeholder='Enter sub code']").clear()
    driver.find_element(By.XPATH, "//input[@placeholder='Enter sub code']").send_keys('REDACTED')
    driver.find_element(By.XPATH, "//input[@placeholder='Enter your staff name']").send_keys('REDACTED')

    #Fill in Traveller Age
    # age = '25'
    driver.find_element(By.XPATH, "//input[@name='adultage1']").send_keys(age)

    #Fill in Travel Period
    depDate = '03/01/2024'
    arrDate = '13/01/2024'
    driver.find_element(By.XPATH, "//input[@name='travelDepartureDate']").clear()
    driver.find_element(By.XPATH, "//input[@name='travelDepartureDate']").send_keys(depDate)
    driver.find_element(By.XPATH, "//input[@name='travelDepartureDate']").click()
    time.sleep(1) #IMPROVEMENT : auto detect the arr date getting filled with dep date, THEN clear and fill
    driver.find_element(By.XPATH, "//input[@name='travelArrivalDate']").clear()
    driver.find_element(By.XPATH, "//input[@name='travelArrivalDate']").click()
    driver.find_element(By.XPATH, "//input[@name='travelArrivalDate']").send_keys(arrDate)

    #Click button to step 2
    driver.find_element(By.CLASS_NAME, "step-button.step2").click()

    #Click chosen plan
    # planType = 'Classic'
    # planType = 'Superior'
    # planType = 'Premier'

    # WebDriverWait(driver, 8).until(EC.visibility_of_element_located((By.XPATH, f"//div[@data-target='#{planType.lower()}']")))
    WebDriverWait(driver, 8).until(EC.element_to_be_clickable((By.XPATH, f"//button[@onclick='goPurchaseTravelGuard_3()']")))
    time.sleep(1)
    driver.find_element(By.XPATH, f"//input[@value='{planType}']").click()

    #Click button to step 3
    driver.find_element(By.XPATH, f"//button[@onclick='goPurchaseTravelGuard_3()']").click()

    #Fill in Personal Information
    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.CLASS_NAME, "pur-trv-grd-ins-first-name")))
    driver.find_element(By.CLASS_NAME, "pur-trv-grd-ins-first-name").send_keys(firstName) #STILL NOT PERFECT
    driver.find_element(By.CLASS_NAME, "pur-trv-grd-ins-last-name").send_keys(lastName) #STILL NOT PERFECT
    driver.find_element(By.CLASS_NAME, "pur-trv-grd-ins-nric").send_keys(nric)


    dobDDi = Select(driver.find_element(By.XPATH, "//select[@name='travel_dd']"))
    dobDDi.select_by_value(dobDD)
    dobMMi = Select(driver.find_element(By.XPATH, "//select[@name='travel_mm']"))
    dobMMi.select_by_value(dobMM)
    dobYYi = Select(driver.find_element(By.XPATH, "//select[@name='travel_yy']"))
    dobYYi.select_by_value(dobYY)

    driver.find_element(By.XPATH, f"//input[@value='{gender}']").click()
    driver.find_element(By.CLASS_NAME, "pur-trv-grd-ins-email").send_keys(email)
    driver.find_element(By.CLASS_NAME, "pur-trv-grd-ins-mob").send_keys(mobileNo)
    driver.find_element(By.XPATH, f"//input[@name='poc'][@value='Yes']").click()
    driver.find_element(By.CLASS_NAME, "pur-trv-grd-ins-passport-no").send_keys(passportNo)

    #Click button to step 4
    driver.find_element(By.XPATH, f"//button[@onclick='goPurchaseTravelGuard_4()']").click()
    WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, "//button[@onclick='goPurchaseTravelGuard_5()']")))
    driver.find_element(By.XPATH, f"//input[@onclick='agreeDisagreeTrvPur(this)']").click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//button[@onclick='goPurchaseTravelGuard_5()']")))
    time.sleep(1)
    driver.find_element(By.XPATH, f"//button[@onclick='goPurchaseTravelGuard_5()']").click()

    time.sleep(100000) #Unlimited sleep to keep browser open for manual payment completion

if __name__ == '__main__':
    processes = []
    for row in range(11):
        p = multiprocessing.Process(target=run, args=(row,))
        p.start()
        processes.append(p)

    for p in processes:
        p.join()

