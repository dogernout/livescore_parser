import win32com.client as win
import pandas as pd
from datetime import datetime, timedelta
import os


class XlBook:

    s_template_path = os.getcwd() + r'\templates\Шаблон_матчей.xlsm'
    s_sheet = 'List_21'

    def __init__(self, path_to_csv: str, days_delta: int):
        self.xl = win.Dispatch('Excel.Application')
        self.xl.DisplayAlerts = False
        self.wb = self.xl.Workbooks.open(self.s_template_path)
        self.sheet = self.wb.Worksheets(self.s_sheet)
        self.df = pd.read_csv(path_to_csv, encoding='utf-8', sep='\t')
        self.date_counter = days_delta

    def write_data(self):
        k = 7
        for i in self.df.index:
            print(i, ':', self.df['team1'][i], '-', self.df['team2'][i])
            self.sheet.Cells(k, 9).value = self.df['team1'][i]
            self.sheet.Cells(k, 18).value = self.df['team2'][i]
            self.sheet.Cells(k, 25).value = self.df['country'][i]
            k += 4
            listz, params = [0] * 6, ['stat_home1', 'stat_home2', 'stat_away1', 'stat_away2', 'res_team1', 'res_team2']
            for ind in range(6):
                listz[ind] = [jind for jind in self.df[params[ind]][i] if jind.isdigit() or jind in 'ВНП']
            d_result = {'В': 3, 'Н': 1, 'П': 0}
            for j in range(10):
                self.sheet.Cells(k + j, 14).value = listz[0][j]
                self.sheet.Cells(k + j, 15).value = listz[2][j]
                self.sheet.Cells(k + j, 16).value = d_result[listz[4][j]]
                self.sheet.Cells(k + j, 30).value = listz[1][j]
                self.sheet.Cells(k + j, 31).value = listz[3][j]
                self.sheet.Cells(k + j, 32).value = d_result[listz[5][j]]
            k += 26
            self.sheet.Cells(k, 21).value = self.df['coef_W'][i]
            self.sheet.Cells(k, 22).value = self.df['coef_D'][i]
            self.sheet.Cells(k, 23).value = self.df['coef_L'][i]
            k += 6

    def quitter(self):
        s_date = datetime.strftime(datetime.today().date() - timedelta(days=self.date_counter), format="%d-%m-%Y")

        if os.path.exists(rf'data\{s_date}.xlsm'): os.remove(rf'data\{s_date}.xlsm')
        self.wb.SaveAs(os.getcwd() + rf'\data\{s_date}.xlsm')
        self.wb.Close()
        self.xl.Quit()


if __name__ == '__main__':
    try:
        a = XlBook(path_to_csv=r'result\10-12-2022.csv', days_delta=2)
        a.write_data()
    finally:
        a.quitter()
