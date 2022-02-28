#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Feb  9 16:36:45 2022

@author: terrydennison
"""

from tkinter import *
import tkinter
from PIL import ImageTk, Image
from tkinter import filedialog
import os
import pandas as pd
import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg


# Creates main window of program
main_window = Tk()
main_window.geometry('1600x1200')
main_window.title("Pueblo Cooperative Care")


# Opens file explorer


def browseFiles():
    global filepath

    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Select Excel File",
                                          filetypes=(("Excel",
                                                      "*.xlsx"),
                                                     ("all files",
                                                      "*.*")))

    filepath = os.path.abspath(filename)
# Creates AP Summmary window


def create_AP_Summary_frame():

    ap_summary_window = Tk()
    ap_summary_window.geometry('1600x1200')

    main_window.withdraw()

    ap_summary_window.title('AP Summary File Upload')
    txt_file = Text(ap_summary_window, height=4,
                    width=50, bg='white', fg='black', font=('Sans Serif', 23, 'italic bold'))
    txt_file.place(x=100, y=350)

    btn_home = Button(ap_summary_window, text="Back to Main", command=lambda: [ap_summary_window.destroy(), main_window.deiconify()],
                      width=20, height=3, fg='green')

    btn_home.place(x=100, y=550)

    btn_file_upload = Button(ap_summary_window, text="Find File", command=lambda: [txt_file.configure(state='normal'), txt_file.delete('1.0', "end"),
                                                                                   browseFiles(), txt_file.insert(INSERT, filepath),
                                                                                   txt_file.configure(state='disabled')], width=20, height=3, fg='green')

    btn_file_upload.place(x=400, y=550)

    def anaylze_AP_summary():
        df = pd.read_excel(
            filepath, sheet_name=1)

        df['Unnamed: 0'] = np.nan
        df.dropna(how='all', axis=1, inplace=True)

        df = pd.melt(df, id_vars='Unnamed: 1',
                     var_name='Date Ranges', value_name='Amount')

        def create_plot():
            axs = sns.catplot(x='Date Ranges', y='Amount',
                              hue='Unnamed: 1', data=df, kind='bar')
            axs.fig.suptitle('AP Summary')
            return axs.fig

        figure = create_plot()
        canvas = FigureCanvasTkAgg(figure, master=ap_summary_window)
        canvas.draw()
        canvas.get_tk_widget().pack(side="left", fill="both", expand=True)

    btn_analyze = Button(ap_summary_window, text="Analyze File", command=anaylze_AP_summary,
                         width=20, height=3, fg='green')

    btn_analyze.place(x=700, y=550)

    ap_summary_window.mainloop()


# Creates main window for AP Summary Option
btn_AP_Summary = Button(main_window, text="AP Summary", command=create_AP_Summary_frame,
                        width=40, height=3, fg='green')
btn_AP_Summary.place(x=50, y=100)


def create_AR_Summary_frame():

    ar_summary_window = Tk()
    ar_summary_window.geometry('1600x1200')

    main_window.destroy()

    ar_summary_window.title('AR Summary File Upload')
    txt_file = Text(ar_summary_window, height=4,
                    width=120, bg='white', fg='black', )
    txt_file.place(x=100, y=350)

    btn_file_upload = Button(ar_summary_window, text="Select File", command=browseFiles,
                             width=20, height=3, fg='green')

    btn_file_upload.place(x=950, y=350)

    ar_summary_window.mainloop()


btn_AR_Summary = Button(main_window, text="AR Summary", command=create_AR_Summary_frame,
                        width=40, height=3, fg='green')
btn_AR_Summary.place(x=50, y=200)


def create_Assets_Liabilities_frame():

    assets_liabilites_window = Tk()
    assets_liabilites_window.geometry('1600x1200')

    main_window.destroy()

    assets_liabilites_window.title('Assets Liabilities File Upload')
    txt_file = Text(assets_liabilites_window, height=4,
                    width=120, bg='white', fg='black', )
    txt_file.place(x=100, y=350)

    btn_file_upload = Button(assets_liabilites_window, text="Select File", command=browseFiles,
                             width=20, height=3, fg='green')

    btn_file_upload.place(x=950, y=350)

    assets_liabilites_window.mainloop()


btn_Assets_Liabilities = Button(main_window, text="Statement of Assets-Liabilities",
                                command=create_Assets_Liabilities_frame, width=40, height=3, fg='green')
btn_Assets_Liabilities.place(x=50, y=300)


def create_Revenue_Expenses_frame():

    revenues_expsenses_window = Tk()
    revenues_expsenses_window.geometry('1600x1200')

    main_window.destroy()

    revenues_expsenses_window.title('Revenue and Expenses File Upload')

    txt_file = Text(revenues_expsenses_window, height=4,
                    width=120, bg='white', fg='black', )
    txt_file.place(x=100, y=350)

    btn_file_upload = Button(revenues_expsenses_window, text="Select File", command=browseFiles,
                             width=20, height=3, fg='green')

    btn_file_upload.place(x=950, y=350)

    revenues_expsenses_window.mainloop()


btn_Revenue_Expenses = Button(main_window, text="Statement of Revenue and Expenses",
                              command=create_Revenue_Expenses_frame, width=40, height=3, fg='green')
btn_Revenue_Expenses.place(x=50, y=400)


def create_Revenue_Expenses_Comparison_frame():

    revenues_expsenses_comparison_window = Tk()
    revenues_expsenses_comparison_window .geometry('1600x1200')

    main_window.destroy()

    revenues_expsenses_comparison_window.title(
        'Revenue and Expenses Comparison File Upload')

    txt_file = Text(revenues_expsenses_comparison_window, height=4,
                    width=120, bg='white', fg='black', )
    txt_file.place(x=100, y=350)

    btn_file_upload = Button(revenues_expsenses_comparison_window, text="Select File", command=browseFiles,
                             width=20, height=3, fg='green')

    btn_file_upload.place(x=950, y=350)

    revenues_expsenses_comparison_window.mainloop()


btn_Revenue_Expenses_Comparison = Button(main_window, text="Statement of Revenue and Expenses Comparison",
                                         command=create_Revenue_Expenses_Comparison_frame, width=40, height=3, fg='green')
btn_Revenue_Expenses_Comparison.place(x=50, y=500)


def create_Revenue_Expenses_Budget_frame():

    revenues_expsenses_budget_window = Tk()
    revenues_expsenses_budget_window .geometry('1600x1200')

    main_window.destroy()

    revenues_expsenses_budget_window.title(
        'Revenue and Expenses Budget File Upload')

    txt_file = Text(revenues_expsenses_budget_window, height=4,
                    width=120, bg='white', fg='black', )
    txt_file.place(x=100, y=350)

    btn_file_upload = Button(revenues_expsenses_budget_window, text="Select File", command=browseFiles,
                             width=20, height=3, fg='green')

    btn_file_upload.place(x=950, y=350)

    revenues_expsenses_budget_window.mainloop()


btn_Revenue_Expenses_Budget = Button(main_window, text="Statement of Revenue and Expenses Budget",
                                     command=create_Revenue_Expenses_Budget_frame, width=40, height=3, fg='green')
btn_Revenue_Expenses_Budget.place(x=50, y=600)


def create_Revenue_Expenses_Program_frame():

    revenues_expsenses_program_window = Tk()
    revenues_expsenses_program_window .geometry('1600x1200')

    main_window.destroy()

    revenues_expsenses_program_window.title(
        'Revenue and Expenses by Program File Upload')

    txt_file = Text(revenues_expsenses_program_window, height=4,
                    width=120, bg='white', fg='black', )
    txt_file.place(x=100, y=350)

    btn_file_upload = Button(revenues_expsenses_program_window, text="Select File", command=browseFiles,
                             width=20, height=3, fg='green')

    btn_file_upload.place(x=950, y=350)

    revenues_expsenses_program_window.mainloop()


btn_Revenue_Expenses_Program = Button(main_window, text="Statement of Revenue and Expenses by Program",
                                      command=create_Revenue_Expenses_Program_frame, width=40, height=3, fg='green')
btn_Revenue_Expenses_Program.place(x=50, y=700)

# Exit button to close program
btn_exit = Button(main_window, text="Exit",
                  command=main_window.destroy, width=40, height=3, fg='green')
btn_exit.place(x=1150, y=900)

# loads Pueblo Cooperative Care image
load = Image.open(
    "/Users/terrydennison/Desktop/Python/Spyder/Pueblo IT/Pueblo-Coop-Center-History.gif")
render = ImageTk.PhotoImage((load))
photo = Label(main_window, image=render)
photo.place(x=600, y=100)


main_window.mainloop()
