from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService  # Similar thing for firefox also!
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import WebDriverException, SessionNotCreatedException, \
    ElementClickInterceptedException, ElementNotInteractableException, NoSuchElementException
from time import sleep, time
import win32com.client as win
import os
import json
from termcolor import colored
import sys


CLICKER = 0


def start_driver():
    options = webdriver.ChromeOptions()
    options.add_extension(r'addons/ublock_1_44_0_0.crx')
    if os.path.exists(os.getcwd() + r'\drivers\.wdm\drivers.json'):
        with open(os.getcwd() + r'\drivers\.wdm\drivers.json') as f: path = json.load(f)
        driver = webdriver.Chrome(service=ChromeService(path.popitem()[1]['binary_path']), options=options)
    else:
        driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager(path=os.getcwd() + r'\drivers').install()),
                                  options=options)
    return driver


def get_data_last_day(days=1):

    def click_cookie(counter=0):
        if counter == 3:
            raise SystemError("Не удалось нажать на кнопку с принятием куки")
        try:
            driver.find_element(By.ID, "onetrust-accept-btn-handler").click()
        except NoSuchElementException:
            sleep(2)
            click_cookie(counter + 1)

    def click_yesterday():
        global CLICKER
        if CLICKER == 3: raise SystemError("Не удалось нажать на кнопку открытия предыдущего дня")
        try:
            driver.find_element(By.CLASS_NAME, 'calendar__navigation--yesterday').click()
        except NoSuchElementException or ElementNotInteractableException:
            CLICKER += 1
            get_data_last_day()

    driver = start_driver()
    driver.get('https://www.livescore.in/ru/')
    sleep(10)
    click_cookie()
    for i in range(days):
        click_yesterday()
        sleep(10)
    try:
        list_home_teams = driver.find_elements(By.CSS_SELECTOR, '.event__participant--home')
        list_away_teams = driver.find_elements(By.CSS_SELECTOR, '.event__participant--away')
        list_home_res = driver.find_elements(By.CSS_SELECTOR, '.event__score--home')
        list_away_res = driver.find_elements(By.CSS_SELECTOR, '.event__score--away')
    except Exception as e:
        print(f'something went wrong: {e}')
    else:
        dict_res = dict()
        if len(list_home_teams) == len(list_away_teams) == len(list_home_res) == len(list_away_res):
            print('OK!')
            for i in range(len(list_home_teams)):
                dict_res[list_home_teams[i].text] = [list_away_teams[i].text,
                                                     list_home_res[i].text,
                                                     list_away_res[i].text]
                #print(list_home_teams[i].text, list_away_teams[i].text, list_home_res[i].text, list_away_res[i].text)
                #print('-' * 100)
        else:
            print('NOT OK!')
            raise SystemError("Количество домашних команд не равно количеству команд противника")
        return dict_res
    finally:
        driver.quit()


def write_data_excel(path, day=1):
    dict_res = get_data_last_day(day)
    xl = win.Dispatch('Excel.Application')
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.open(path)
    except Exception as e:
        raise SystemExit(f'exception occurred while opening workbook: {e}')
    sheet = wb.Worksheets('List_21')
    int_index = 7
    try:
        l_mistakes = []
        while sheet.cells(int_index, 9).value is not None:
            if sheet.cells(int_index, 9).value not in dict_res:
                l_mistakes.append(sheet.cells(int_index, 9).value)
            else:
                sheet.cells(int_index + 24, 13).value = dict_res[sheet.cells(int_index, 9).value][1]
                sheet.cells(int_index + 24, 14).value = dict_res[sheet.cells(int_index, 9).value][2]
            int_index += 36
    except Exception as exp:
        print("EXCEPTION OCCURRED:", exp)
    finally:
        wb.Save()
        wb.Close()
    print('All Done!')
    print('MISTAKES:')
    for i in l_mistakes: print(i)


if __name__ == '__main__':
    if len(sys.argv) == 3:
        write_data_excel(path=sys.argv[1], day=int(sys.argv[2]))
    else: write_data_excel(input("Введите путь к файлу с матчами ВЧЕРАШНЕГО дня:"))
