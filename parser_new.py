from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService  # Similar thing for firefox also!
#from subprocess import CREATE_NO_WINDOW  # This flag will only be available in windows
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import WebDriverException, SessionNotCreatedException, \
    ElementClickInterceptedException, ElementNotInteractableException, NoSuchElementException
from time import sleep, time
import win32com.client as win
import os
import json
from termcolor import colored
from datetime import datetime, timedelta
import sys
from shutil import rmtree
import pandas as pd


class Parser:

    CLICKER = 0
    ROW = 7

    def __init__(self, date_to_parse=1, headless=False):
        self.url = 'https://livescore.in/ru'
        self.driver = self.get_driver(headless)
        self.date_counter = date_to_parse
        self.xl, self.wb, self.sheet = None, None, None
        self.df = pd.DataFrame(columns=['country', 'coef_W', 'coef_D', 'coef_L', 'team1', 'stat_home1', 'stat_away1',
                                        'res_team1', 'team2', 'stat_home2', 'stat_away2', 'res_team2', 'res_home',
                                        'res_away', 'stat_res'])
        self.get_path()

    @staticmethod
    def get_driver(headless):
        options = webdriver.ChromeOptions()
        options.add_extension(r'addons/ublock_1_44_0_0.crx')
        options.headless = headless
        if os.path.exists(os.getcwd() + r'\drivers\.wdm\drivers.json'):
            with open(os.getcwd() + r'\drivers\.wdm\drivers.json') as f: path = json.load(f)
            try:
                driver = webdriver.Chrome(service=ChromeService(path.popitem()[1]['binary_path']), options=options)
            except SessionNotCreatedException:
                rmtree("drivers", ignore_errors=True)
                os.mkdir("drivers")
                driver = webdriver.Chrome(
                    service=ChromeService(ChromeDriverManager(path=os.getcwd() + r'\drivers').install()),
                    options=options)
        else:
            driver = webdriver.Chrome(
                service=ChromeService(ChromeDriverManager(path=os.getcwd() + r'\drivers').install()),
                options=options)
        return driver

    @staticmethod
    def get_path():
        if not os.path.exists(r'result'): os.mkdir(r'result')

    @staticmethod
    def check_coefficients(l_current_teams_coeffs):
        if len(l_current_teams_coeffs) < 3: return False
        i = 0
        while i < len(l_current_teams_coeffs):  # Обычно встречаются коэффы только от двух букмекеров, поэтому
            # элементов не больше шести (3 у первого и 3 у второго)
            if l_current_teams_coeffs[i][1] == '-':
                if i < 3:
                    del l_current_teams_coeffs[:3]  # Ошибка у первого букмекера
                    i = 0
                else:
                    del l_current_teams_coeffs[3:]  # Ошибка у второго букмекера
                    break
            i += 1
        return len(l_current_teams_coeffs) > 0

    def parse_me_daddy(self):

        def click_cookie(counter=0):
            if counter == 3:
                raise SystemError("exception: Не удалось нажать на кнопку с принятием куки")
            try:
                self.driver.find_element(By.ID, "onetrust-accept-btn-handler").click()
            except NoSuchElementException:
                sleep(3)
                click_cookie(counter + 1)

        def click_tomorrow():
            if self.CLICKER == 3: raise SystemError("exception: Не удалось нажать на кнопку открытия нужного дня")
            try:
                self.driver.find_element(By.CLASS_NAME, 'calendar__navigation--tomorrow').click()
            except NoSuchElementException or ElementNotInteractableException:
                self.CLICKER += 1
                self.parse_me_daddy()

        def parse_odd_tab(counter=0):
            if counter == 50: raise SystemError("exception: Fatal Error")
            try:
                l_all_home_teams[i].click()
            except ElementClickInterceptedException:
                body = self.driver.find_element(By.CSS_SELECTOR, 'body')
                for _ in range(3): body.send_keys(Keys.ARROW_DOWN)
                parse_odd_tab(counter+1)
            except ElementNotInteractableException:  # FIX THIS SHEEEIDT
                print("exception: ElementNotInteractableException", flush=True)
                parse_odd_tab(counter+1)
            sleep(1)
            if len(self.driver.window_handles) == 1:
                print("strange thing..", flush=True)
                parse_odd_tab(counter+1)
            coefficients_tab = self.driver.window_handles[1]
            self.driver.switch_to.window(coefficients_tab)
            s_country = self.driver.find_element(By.CLASS_NAME, 'tournamentHeader__country').text
            current_teams_table = self.driver.find_element(By.CLASS_NAME, 'container__detail')
            l_current_teams_coeffs = current_teams_table.find_elements(By.CLASS_NAME, 'oddsValueInner')
            l_current_teams_coeffs = [[['1', 'X', '2'][n], l_current_teams_coeffs[n].get_attribute("textContent")]
                                      for n in range(len(l_current_teams_coeffs))]
            s_status = self.driver.find_element(By.CLASS_NAME, 'detailScore__status').text
            if not self.check_coefficients(l_current_teams_coeffs):
                print(f"skipping {team1_name} VS {team2_name} match because of odds issues..", flush=True)
                self.driver.close()
                self.driver.switch_to.window(main_tab)
                return None, None, None, None
            if 'TKP' in s_status or 'Отменен' in s_status or 'Перенесен' in s_status:
                print(f"skipping {team1_name} VS {team2_name} match because of ТКР/Отменен/Перенесен..", flush=True)
                self.driver.close()
                self.driver.switch_to.window(main_tab)
                return None, None, None, None
            l_teams_sel = self.driver.find_elements(By.CSS_SELECTOR, '.participant__participantLink--team')
            for ind in l_teams_sel: ind.click()
            self.driver.close()
            self.driver.switch_to.window(main_tab)
            return s_country, s_status, l_current_teams_coeffs, l_teams_sel

        def parse_scores_tab(sel_tab, team):

            def check_scores():
                if len(l_home_scores) != 10 or len(l_away_scores) != 10 or len(l_results) != 10: return False
                for ind in range(10):
                    if not l_home_scores[ind].isdigit() or not l_away_scores[ind].isdigit(): return False
                return True

            self.driver.switch_to.window(self.driver.window_handles[-1])
            sleep(1)
            s_name = self.driver.find_element(By.CLASS_NAME, 'heading__name').text
            l_home_scores = [m.text for m in self.driver.find_elements(By.CSS_SELECTOR, '.event__score--home')
                             if m.text != '-']
            l_away_scores = [m.text for m in self.driver.find_elements(By.CSS_SELECTOR, '.event__score--away')
                             if m.text != '-']
            l_results = [m.text for m in self.driver.find_elements(By.CLASS_NAME, 'formIcon') if m.text in 'ВНП']
            self.driver.close()
            self.driver.switch_to.window(main_tab)
            if check_scores(): return s_name, l_home_scores, l_away_scores, l_results
            else: return None, None, None, None

        self.driver.get(self.url)
        sleep(10)
        click_cookie()
        sleep(10)
        for i in range(self.date_counter):
            click_tomorrow()
            sleep(10)
        l_all_home_teams = self.driver.find_elements(By.CSS_SELECTOR, '.event__participant--home')
        l_all_away_teams = self.driver.find_elements(By.CSS_SELECTOR, '.event__participant--away')
        main_tab = self.driver.window_handles[0]
        print(f'all_len {len(l_all_home_teams)}')#, flush=True)
        i = 0
        while i < len(l_all_home_teams):
            l_info = []
            print(f'current {i}')#, flush=True)
            try:
                team1_name = l_all_home_teams[i].text
                team2_name = l_all_away_teams[i].text
            except Exception as e:
                print(f"exception in opening coefficients tab. Index: {i}. Description: {e}")#, flush=True)
                continue
            if '(Ж)' in team1_name or '(Ж)' in team2_name:
                print(f"skipping {team1_name} VS {team2_name} match because of female teams..")#, flush=True)
                i += 1
                continue
            s_country, s_status, l_current_teams_coeffs, l_cur_teams = parse_odd_tab()
            if l_cur_teams is None:
                i += 1
                continue
            l_info.append([s_country, s_status, l_current_teams_coeffs])
            for j in range(len(l_cur_teams)):
                s_name, l_home_scores, l_away_scores, l_results = parse_scores_tab(l_cur_teams[j], j)
                if s_name is not None: l_info.append([s_name, l_home_scores, l_away_scores, l_results])
            if len(l_info) == 3:
                print(f"l_info: {l_info}")
                self.df = self.df.append({'country': s_country, 'coef_W': l_current_teams_coeffs[0][1],
                                          'coef_D': l_current_teams_coeffs[1][1],
                                          'coef_L': l_current_teams_coeffs[2][1], 'team1': l_info[1][0],
                                          'stat_home1': l_info[1][1], 'stat_away1': l_info[1][2],
                                          'res_team1': l_info[1][3], 'team2': l_info[2][0], 'stat_home2': l_info[2][1],
                                          'stat_away2': l_info[2][2], 'res_team2': l_info[2][3], 'res_home': None,
                                          'res_away': None, 'stat_res': None}, ignore_index=True)
                print(f"writing info about {team1_name} VS {team2_name} match")
            else:
                print(f"skipping {team1_name} VS {team2_name} match because of not enough matches (<10)..")
            i += 1
        #self.driver.quit()
        #self.df.to_csv(r'/data/test.csv', encoding='utf-8')
        self.quitter()
        print('all done')

    def quitter(self):
        self.driver.quit()
        s_date = datetime.strftime(datetime.today().date() + timedelta(days=self.date_counter), format="%d-%m-%Y")
        if os.path.exists(rf'result\{s_date}.csv'): os.remove(rf'result\{s_date}.csv')
        self.df.to_csv(rf'result\{s_date}.csv', encoding='utf-8', sep='\t')


if __name__ == '__main__':
    if len(sys.argv) > 1:
        aboba = Parser(date_to_parse=int(sys.argv[1]))
    else:
        aboba = Parser()
    try:
        aboba.parse_me_daddy()
    finally:
        aboba.quitter()
