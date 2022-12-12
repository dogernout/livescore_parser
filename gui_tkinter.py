from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from datetime import datetime, timedelta
import subprocess
import parser_new
import parser_last_day
import os
import threading
from time import sleep


def show_window():

    def chozen_date(event=None):
        pass

    def chozen_settings(event=None):
        pass

    def chozen_mode(event=None):
        if ivar_mode.get():  # Fill Results
            fr_mode0.pack_forget()
            fr_mode1.pack(side=TOP, fill=BOTH, expand=TRUE)
        else:
            fr_mode1.pack_forget()
            fr_mode0.pack(side=TOP, fill=BOTH, expand=TRUE)

    def btn_clicked(event=None):
        if not ivar_mode.get():
            fr_mode0.pack_forget()
            fr_modes.pack_forget()
            fr_mode_run = Frame(window, borderwidth=0)
            fr_mode_run.pack(side=TOP, fill=BOTH, expand=TRUE)
            lbl_percent = Label(fr_mode_run, font=('arial', 14), text='0%', width=8, background=s_maincolor)
            lbl_percent.grid(row=0, column=0, pady=16, sticky=NW)
            window.update()

            prg_bar = ttk.Progressbar(fr_mode_run, orient=HORIZONTAL, length=610, mode='determinate')
            prg_bar.grid(row=1, column=0, columnspan=2, sticky=NSEW, padx=(16, 16))
            window.update()

            fr_lbx = Frame(fr_mode_run, borderwidth=0)
            fr_lbx.grid(row=2, column=0, columnspan=2, pady=16)

            lbx_status = Listbox(fr_lbx, font=('arial', 10), height=5, relief=GROOVE, borderwidth=3, width=84)
            lbx_status.pack(side=LEFT, fill=BOTH)
            vsb = Scrollbar(fr_lbx, orient=VERTICAL, command=lbx_status.yview)
            vsb.pack(side=RIGHT, fill=BOTH)
            lbx_status.config(yscrollcommand=vsb.set)
            window.update()

            process = subprocess.Popen(f"py -u parser_new.py {ivar_date.get() + 1} {ivar_settings.get()}", shell=True,
                                       stdout=subprocess.PIPE, universal_newlines=True, encoding='utf-8')
            while True:
                try:
                    output = process.stdout.readline().strip()
                except Exception as e:
                    print('exception:', e)
                    output = ' '
                if process.poll() is not None: break
                if 'all_len' in output:
                    i_len = int(output.split()[1])
                if 'current' in output:
                    i_current = int(output.split()[1])
                    prg_bar['value'] = round((610 * i_current / i_len) / 610 * 100, 2)
                    lbl_percent.config(text=str(round((610 * i_current / i_len) / 610 * 100, 2)) + '%')
                    window.update()
                if 'writing' in output or 'skipping' in output:
                    lbx_status.insert(0, output)
                    window.update()
                window.update()

            lbl_percent.config(text='All Done')
            prg_bar['value'] = 100
            window.update()
            lbl_percent.config(text='Now filling data in xlsm..')
            s_date = datetime.strftime(datetime.today().date() + timedelta(days=ivar_date.get() + 1), format="%d-%m-%Y")
            process = subprocess.Popen(fr"py -u csv_to_xlsm.py result\{s_date}.csv {ivar_date.get() + 1}", shell=True,
                                       stdout=subprocess.PIPE, universal_newlines=True, encoding='utf-8')
            while True:
                try:
                    output = process.stdout.readline().strip()
                except Exception as e:
                    print('exception:', e)
                    output = ' '
                if 'all_len' in output:
                    i_len = int(output.split()[1])
                elif 'current' in output:
                    i_current = int(output.split()[1])
                    prg_bar['value'] = round((610 * i_current / i_len) / 610 * 100, 2)
                    lbl_percent.config(text=str(round((610 * i_current / i_len) / 610 * 100, 2)) + '%')
                else:
                    lbx_status.insert(0, output)
                if process.poll() is not None: break

                window.update()
            prg_bar['value'] = 100
            lbl_percent.config(text='END')
            window.update()

        else:
            s_filename = filedialog.askopenfilename(initialdir=os.getcwd(), title='Choose file to upload',
                                                    filetypes=(('excel files', '*.xls*'),))
            print(s_filename, len(s_filename), f"date - {ivar_date.get() + 1}", sep='\n')
            if not len(s_filename): return
            fr_modes.pack_forget()
            fr_mode0.pack_forget()
            fr_mode1.pack_forget()
            fr_run = Frame(window, borderwidth=0)
            fr_run.pack(side=TOP, fill=BOTH, expand=TRUE)
            prg_bar = ttk.Progressbar(fr_run, orient=HORIZONTAL, length=610, mode='determinate')
            prg_bar.grid(row=1, column=0, columnspan=2, sticky=NSEW, padx=(16, 16))
            prg_bar['value'] = 0
            prg_bar.start(interval=32)
            lbl_status = Label(fr_run, text='Work In Progress', font=('arial', 14))
            lbl_status.grid(row=0, column=0, sticky=EW, columnspan=2)
            window.update()
            process = subprocess.Popen(f"parser_last_day.py {s_filename} {ivar_date.get() + 1}", shell=True,
                                       stdout=subprocess.PIPE, universal_newlines=True)
            while True:
                try:
                    output = process.stdout.readline().strip()
                except Exception as e:
                    print('exception:', e)
                    output = ' '
                if process.poll() is not None: break
                if 'all_len' in output:
                    print(output)
                    prg_bar['value'] = 2
                    window.update()
                if 'current' in output:
                    print(output)
                    prg_bar['value'] = 10
                    window.update()
                print(output)

            lbl_status.config(text='All Done')
            window.update()


    def btn_upload_clicked(event=None):
        s_filename = filedialog.askopenfilename(initialdir=os.getcwd(), title='Choose file to upload',
                                                filetypes=(('excel files', '*.xls*'),))
        print(s_filename, len(s_filename), sep='\n')

    def gridder(fr_to, l_dates1, l_text):
        fr_day = LabelFrame(fr_to, borderwidth=3, relief=GROOVE, text='Choose Day:', font=('arial', 14),
                            labelanchor=N,
                            bg=s_maincolor)
        fr_day.grid(row=1, column=0, padx=32, pady=10)
        fr_settings = LabelFrame(fr_to, borderwidth=3, relief=GROOVE, text='Show Browser:', font=('arial', 14),
                                 labelanchor=N, bg=s_maincolor)
        fr_settings.grid(row=1, column=1, padx=92, pady=10, sticky=N)
        for i in range(len(l_dates1)):
            rbtn_date = Radiobutton(fr_day, value=i, variable=ivar_date, text=l_dates1[i] + f' ({l_text[0]})' * (not i),
                                    command=chozen_date, font=('arial', 12), bg=s_maincolor, anchor=W,
                                    activebackground=s_maincolor, width=20)
            rbtn_date.grid(row=i, column=0)

        for i in range(2):
            rbtn_settings = Radiobutton(fr_settings, value=i, variable=ivar_settings, text='Yes' if i else 'No (later)',
                                        command=chozen_settings, font=('arial', 12), bg=s_maincolor, anchor=W,
                                        activebackground=s_maincolor, width=20, state=DISABLED)
            rbtn_settings.grid(row=i, column=0, padx=(10, 10))

        if not ivar_mode.get():
            btn_run = Button(fr_to, text=l_text[1], font=('arial', 14), bg=s_maincolor, relief=GROOVE,
                             borderwidth=3, command=btn_clicked, width=20)
            btn_run.grid(row=2, column=1, sticky=N)
        else:
            btn_upload = Button(fr_to, text=l_text[1], font=('arial', 14), bg=s_maincolor, relief=GROOVE,
                                borderwidth=3, command=btn_upload_clicked, width=19)
            btn_upload.grid(row=2, column=1, sticky=N)

    window = Tk()
    window.title("Livescore Selenium Parser")
    s_maincolor = 'white smoke'
    window.configure(background=s_maincolor)
    window.iconphoto(False, PhotoImage(file=r'icons\my_logo.png'))
    window.geometry(f"640x240+{window.winfo_screenwidth() // 2 - 320}+{window.winfo_screenheight() // 2 - 120}")

    fr_mode0 = Frame(window, borderwidth=0)
    fr_mode0.pack(side=BOTTOM, fill=BOTH, expand=TRUE)

    fr_mode1 = Frame(window, borderwidth=0)

    fr_modes = Frame(window, borderwidth=0)
    fr_modes.pack(side=TOP, fill=BOTH, expand=TRUE)

    ivar_mode = IntVar(value=0)
    for i in range(2):
        rbtn_mode = Radiobutton(fr_modes, value=i, variable=ivar_mode, text='Fill Results' if i else 'Parser',
                                command=chozen_mode, font=('arial', 12), bg=s_maincolor, anchor=W,
                                activebackground=s_maincolor, width=20)
        rbtn_mode.grid(row=0, column=i, pady=4, padx=64)

    l_dates = [datetime.strftime(datetime.today().date() + timedelta(days=i), format="%d-%m-%Y") for i in range(1, 4)]
    l_dates2 = [datetime.strftime(datetime.today().date() - timedelta(days=i), format="%d-%m-%Y") for i in range(1, 4)]
    ivar_date = IntVar(value=0)

    ivar_settings = IntVar(value=1)

    gridder(fr_mode0, l_dates, ['tomorrow', 'Start Parser'])
    gridder(fr_mode1, l_dates2, ['yesterday', 'Upload File'])

    window.mainloop()


if __name__ == '__main__':
    if not os.path.exists(os.getcwd()+r'/data'): os.mkdir(os.getcwd()+'/data')
    show_window()
