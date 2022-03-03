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
import io
import os
import pandas as pd
import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import tempfile
import subprocess
import tempfile


# Creates main window of program
main_window = Tk()
main_window.geometry('1600x1200')
main_window.title("Pueblo Cooperative Care")


#Opens file explorer
def browseFiles():
    global filepath

    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Select Excel File",
                                          filetypes=(("Excel",
                                                      "*.xlsx"),
                                                     ("all files",
                                                      "*.*")))

    filepath = os.path.abspath(filename)
#@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ All Features Working Except: pdf_print() @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
#@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ Not Yet Built: Save_As_PDF fuction @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

#Creates AP Summmary window
def create_AP_Summary_frame():
    #Creates GUI object and sets size of GUI
    ap_summary_window = Tk()
    ap_summary_window.geometry('1600x1200')
    
    #Hides 'Main Screen' GUI
    main_window.withdraw()
    
    #Sets title of 'AP Summary File Upload' GUI
    ap_summary_window.title('AP Summary File Upload')
    
    #Creates textbox for 'AP Summary Frame' GUI and sets location on GUI
    txt_file = Text(ap_summary_window, height=4,
                    width=50, bg='white', fg='black', font=('Sans Serif', 23, 'italic bold'))
    txt_file.place(x=100, y=350)
    
    #Button for going back to 'Pueblo Cooperative Care (Home)' GUI
    btn_home = Button(ap_summary_window, text="Back to Main", command=lambda: [ap_summary_window.destroy(), main_window.deiconify()],
                      width=20, height=3, fg='green')

    #Button for selecting 'AP Summary' file for upload and location on 'AP Summary Frame' GUI
    btn_home.place(x=100, y=550)
    btn_file_upload = Button(ap_summary_window, text="Find File", command=lambda: [txt_file.configure(state='normal'), txt_file.delete('1.0', "end"),
                                                                                   browseFiles(), txt_file.insert(INSERT, filepath),
                                                                                   txt_file.configure(state='disabled')], width=20, height=3, fg='green')
    btn_file_upload.place(x=400, y=550)

    #Nested fuction for ETL of 'AP Summary' .xslx file
    def anaylze_AP_summary():
        df = pd.read_excel(
            filepath, sheet_name=1)

        df['Unnamed: 0'] = np.nan
        df.dropna(how='all', axis=1, inplace=True)

        df = pd.melt(df, id_vars='Unnamed: 1',
                     var_name='Date Ranges', value_name='Amount')
        
        #Nested function for creating figure for 'grouped bar chart'
        def create_plot():
            axs = sns.catplot(x='Date Ranges', y='Amount',
                              hue='Unnamed: 1', data=df, kind='bar')
            axs.fig.suptitle('AP Summary')
            return axs.fig
        #Creates window for graphical display of data and size of GUI frame
        ap_analyze_window = Tk()
        ap_analyze_window.geometry('1600x1200')
        #Calls create_plot() function to put 'grouped bar chart' in 'Canvas' frame
        figure = create_plot()
        canvas = FigureCanvasTkAgg(figure, master=ap_analyze_window)
        canvas.draw()
        #Sets position of created 'Canvas'
        canvas.get_tk_widget().pack(side='top')
        canvas.get_tk_widget().place(bordermode=OUTSIDE)
        
        #$$$$$$$$$$$$$$$$$$$$$$$ Place Holder for Pie Graph $$$$$$$$$$$$$$$$$$$$$$$$$$$$






        

        #$$$$$$$$$$$$$$$$$$$$$$$ Place Holder for Table $$$$$$$$$$$$$$$$$$$$$$$$$$$$
        #Hides 'AP Summary GUI'
        ap_summary_window.withdraw()
        
        #WIP
        def pdf_Print():
            f= 5
          # fp = tempfile.TemporaryFile()
          # fp.write(ap_analyze_window)

          # ps = canvas..postscript(colormode='color')
          # img = Image.open(io.BytesIO(ps.encode('utf-8')))
          # img.save('filename.jpg', 'jpeg')

          # postscript_file = "tmp_snapshot.ps"
          # subprocess.call(["impg", "-window", "td", postscript_file])
        
        #WIP Button for Printing 'Canvas' and button position
        btn_print = Button(ap_analyze_window, text="Print as PDF", command=pdf_Print(),
                           width=20, height=3, fg='green')

        btn_print.place(x=100, y=550)
        
        #WIP Button for saving file as 'PDF' and button position
        btn_pdf = Button(ap_analyze_window, text="Save as PDF", command=lambda: [ap_summary_window.destroy(), main_window.deiconify()],
                         width=20, height=3, fg='green')
        btn_pdf.place(x=500, y=550)
        
        #Display's 'AP Analyze' GUI frame
        ap_analyze_window.mainloop()

    #Button event handler for calling 'analyze_ap_summary' function and button position
    btn_analyze = Button(ap_summary_window, text="Analyze File", command=anaylze_AP_summary,
                         width=20, height=3, fg='green')
    btn_analyze.place(x=700, y=550)
    
    #Displays 'AP Summary' GUI    
    ap_summary_window.mainloop()


# Creates Button on main window for 'AP Summary' Option
btn_AP_Summary = Button(main_window, text="AP Summary", command=create_AP_Summary_frame,
                        width=40, height=3, fg='green')
btn_AP_Summary.place(x=50, y=100)

#@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ Duplicate Code from 'AP Summary' Section Above For Each Section @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

#Creates AR Summmary window
def create_AR_Summary_frame():
    #Creates GUI object and sets size of GUI
    ar_summary_window = Tk()
    ar_summary_window.geometry('1600x1200')
    
    #Hides 'Main Screen' GUI
    main_window.withdraw()
    
    #Sets title of 'AR Summary File Upload' GUI
    ar_summary_window.title('AR Summary File Upload')
    
    #Creates textbox for 'AR Summary Frame' GUI and sets location on GUI
    ar_txt_file = Text(ar_summary_window, height=4,
                    width=50, bg='white', fg='black', font=('Sans Serif', 23, 'italic bold'))
    ar_txt_file.place(x=100, y=350)
    
    #Button for going back to 'Pueblo Cooperative Care (Home)' GUI
    btn_home = Button(ar_summary_window, text="Back to Main", command=lambda: [ar_summary_window.destroy(), main_window.deiconify()],
                      width=20, height=3, fg='green')

    #Button for selecting 'AR Summary' file for upload and location on 'AR Summary Frame' GUI
    btn_home.place(x=100, y=550)
    btn_file_upload = Button(ar_summary_window, text="Find File", command=lambda: [ar_txt_file.configure(state='normal'), ar_txt_file.delete('1.0', "end"),
                                                                                   browseFiles(), ar_txt_file.insert(INSERT, filepath),
                                                                                   ar_txt_file.configure(state='disabled')], width=20, height=3, fg='green')
    btn_file_upload.place(x=400, y=550)

    #Nested fuction for ETL of 'AR Summary' .xslx file
    def anaylze_AR_summary():
        df = pd.read_excel(
            filepath, sheet_name=1)

        df['Unnamed: 0'] = np.nan
        df.dropna(how='all', axis=1, inplace=True)

        df = pd.melt(df, id_vars='Unnamed: 1',
                     var_name='Date Ranges', value_name='Amount')
        
        #Nested function for creating figure for 'grouped bar chart'
        def create_plot():
            axs = sns.catplot(x='Date Ranges', y='Amount',
                              hue='Unnamed: 1', data=df, kind='bar')
            axs.fig.suptitle('AR Summary')
            return axs.fig
        #Creates window for grARhical display of data and size of GUI frame
        ar_analyze_window = Tk()
        ar_analyze_window.geometry('1600x1200')
        #Calls create_plot() function to put 'grouped bar chart' in 'Canvas' frame
        figure = create_plot()
        canvas = FigureCanvasTkAgg(figure, master=ar_analyze_window)
        canvas.draw()
        #Sets position of created 'Canvas'
        canvas.get_tk_widget().pack(side='top')
        canvas.get_tk_widget().place(bordermode=OUTSIDE)
        
        #Hides 'AR Summary GUI'
        ar_summary_window.withdraw()
        
        #WIP
        def pdf_Print():
            f= 5
          # fp = tempfile.TemporaryFile()
          # fp.write(ar_analyze_window)

          # ps = canvas..postscript(colormode='color')
          # img = Image.open(io.BytesIO(ps.encode('utf-8')))
          # img.save('filename.jpg', 'jpeg')

          # postscript_file = "tmp_snARshot.ps"
          # subprocess.call(["impg", "-window", "td", postscript_file])
        
        #WIP Button for Printing 'Canvas' and button position
        btn_print = Button(ar_analyze_window, text="Print as PDF", command=pdf_Print(),
                           width=20, height=3, fg='green')

        btn_print.place(x=100, y=550)
        
        #WIP Button for saving file as 'PDF' and button position
        btn_pdf = Button(ar_analyze_window, text="Save as PDF", command=lambda: [ar_summary_window.destroy(), main_window.deiconify()],
                         width=20, height=3, fg='green')
        btn_pdf.place(x=500, y=550)
        
        #Display's 'AR Analyze' GUI frame
        ar_analyze_window.mainloop()

    #Button event handler for calling 'analyze_AR_summary' function and button position
    btn_analyze = Button(ar_summary_window, text="Analyze File", command=anaylze_AR_summary,
                         width=20, height=3, fg='green')
    btn_analyze.place(x=700, y=550)
    
    #Displays 'AR Summary' GUI    
    ar_summary_window.mainloop()

# Creates Button on main window for 'AR Summary' Option
btn_AR_Summary = Button(main_window, text="AR Summary", command=create_AR_Summary_frame,
                        width=40, height=3, fg='green')
btn_AR_Summary.place(x=50, y=200)

#Creates 'Assets Liabilities' window
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

# Creates Button on main window for 'Statement of Assets-Liabilities' Option
btn_Assets_Liabilities = Button(main_window, text="Statement of Assets-Liabilities",
                                command=create_Assets_Liabilities_frame, width=40, height=3, fg='green')
btn_Assets_Liabilities.place(x=50, y=300)

#Creates 'Revenue and Expenses' window
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

# Creates Button on main window for 'Statement Revenue and Expenses' Option
btn_Revenue_Expenses = Button(main_window, text="Statement of Revenue and Expenses",
                              command=create_Revenue_Expenses_frame, width=40, height=3, fg='green')
btn_Revenue_Expenses.place(x=50, y=400)

#Creates 'Revenue and Expenses Comparison' window
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

# Creates Button on main window for 'Statement of Assets-Liabilities' Option
btn_Revenue_Expenses_Comparison = Button(main_window, text="Statement of Revenue and Expenses Comparison",
                                         command=create_Revenue_Expenses_Comparison_frame, width=40, height=3, fg='green')
btn_Revenue_Expenses_Comparison.place(x=50, y=500)

#Creates 'Revenue and Expenses Compared to Budget' window
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

# Creates Button on main window for 'Statement of Revenue and Expense Budget' Option
btn_Revenue_Expenses_Budget = Button(main_window, text="Statement of Revenue and Expenses Budget",
                                     command=create_Revenue_Expenses_Budget_frame, width=40, height=3, fg='green')
btn_Revenue_Expenses_Budget.place(x=50, y=600)

#Creates 'Revenue and Expenses by Program' window
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

# Creates Button on main window for 'Statement oof Revenue and Expenses by Program' Option
btn_Revenue_Expenses_Program = Button(main_window, text="Statement of Revenue and Expenses by Program",
                                      command=create_Revenue_Expenses_Program_frame, width=40, height=3, fg='green')
btn_Revenue_Expenses_Program.place(x=50, y=700)

# Exit button to close program
btn_exit = Button(main_window, text="Exit",
                  command=main_window.destroy, width=40, height=3, fg='green')
btn_exit.place(x=1150, y=900)

# loads Pueblo Cooperative Care image
#Will need to be changed so can be accessed globally
load = Image.open(
    "/Users/terrydennison/Desktop/Python/Spyder/Pueblo IT/Pueblo-Coop-Center-History.gif")
render = ImageTk.PhotoImage((load))
photo = Label(main_window, image=render)
photo.place(x=600, y=100)


main_window.mainloop()
