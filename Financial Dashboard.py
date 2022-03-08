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

# @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ Main Window @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
# Creates main window of program
main_window = Tk()
main_window.geometry('1600x1200')
main_window.title("Pueblo Cooperative Care")


# Open file explorer method
def browseFiles():
    global filepath

    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Select Excel File",
                                          filetypes=(("Excel",
                                                      "*.xlsx"),
                                                     ("all files",
                                                      "*.*")))

    filepath = os.path.abspath(filename)

# @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ AP Summary Window @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
# Creates AP Summmary window


def create_AP_Summary_frame():
    # Creates GUI object and sets size of GUI
    ap_summary_window = Tk()
    ap_summary_window.geometry('1600x1200')

    # Hides 'Main Screen' GUI
    main_window.withdraw()

    # Sets title of 'AP Summary File Upload' GUI
    ap_summary_window.title('AP Summary File Upload')

    # Creates textbox for 'AP Summary Frame' GUI and sets location on GUI
    txt_file = Text(ap_summary_window, height=4,
                    width=50, bg='white', fg='black', font=('Sans Serif', 23, 'italic bold'))
    txt_file.place(x=100, y=350)

    # Button for going back to 'Pueblo Cooperative Care (Home)' GUI
    btn_home = Button(ap_summary_window, text="Back to Main", command=lambda: [ap_summary_window.destroy(), main_window.deiconify()],
                      width=20, height=3, fg='green')

    # Button for selecting 'AP Summary' file for upload and location on 'AP Summary Frame' GUI
    btn_home.place(x=100, y=550)
    btn_file_upload = Button(ap_summary_window, text="Find File", command=lambda: [txt_file.configure(state='normal'), txt_file.delete('1.0', "end"),
                                                                                   browseFiles(), txt_file.insert(INSERT, filepath),
                                                                                   txt_file.configure(state='disabled')], width=20, height=3, fg='green')
    btn_file_upload.place(x=400, y=550)

    # Nested fuction for ETL of 'AP Summary' .xslx file
    def anaylze_AP_summary():
        df = pd.read_excel(
            filepath, sheet_name=1)

        df['Unnamed: 0'] = np.nan
        df.dropna(how='all', axis=1, inplace=True)

        df = pd.melt(df, id_vars='Unnamed: 1',
                     var_name='Date Ranges', value_name='Amount')

        # Nested function for creating figure for 'grouped bar chart'
        def create_plot():
            axs = sns.catplot(x='Date Ranges', y='Amount',
                              hue='Unnamed: 1', data=df, kind='bar')
            axs.fig.suptitle('AP Summary')
            axs.fig.set_size_inches(8.5, 7)
            return axs.fig
        # Creates window for graphical display of data and size of GUI frame
        ap_analyze_window = Tk()
        ap_analyze_window.geometry('1900x1200')
        ap_analyze_window.state('zoomed')
        # Calls create_plot() function to put 'grouped bar chart' in 'Canvas' frame
        figure = create_plot()
        canvas = FigureCanvasTkAgg(figure, master=ap_analyze_window)
        canvas.draw()
        # Sets position of created 'Canvas'
        canvas.get_tk_widget().pack(side='top')
        canvas.get_tk_widget().place(bordermode=OUTSIDE)

        fig = plt.figure(figsize=(6, 6))
        fig.set_size_inches(8.5, 7)
        plt.axis('equal')
        pie_labels = (df['Unnamed: 1'][25], df['Unnamed: 1'][26],
                      df['Unnamed: 1'][27], df['Unnamed: 1'][28])
        pie_data = (df['Amount'][25], df['Amount'][26],
                    df['Amount'][27], df['Amount'][28])
        colors = sns.color_palette('bright')[0:4]

        plt.title("AP Summary Totals")
        pie_graph = plt.pie(pie_data, labels=pie_labels, colors=colors,
                            autopct='%.0f%%', pctdistance=.5, explode=[0.1]*4, textprops={"fontsize": 15}, shadow=True)
        plt.legend()

        canvas = FigureCanvasTkAgg(fig, master=ap_analyze_window)
        canvas.draw()
        canvas.get_tk_widget().place(relx=.6, rely=.78, anchor=CENTER)
        canvas.get_tk_widget().pack(side='top')

        # Hides 'AP Summary GUI'
        ap_summary_window.withdraw()

        # WIP

        def pdf_Print():
            f = 5
          # fp = tempfile.TemporaryFile()
          # fp.write(ap_analyze_window)

          # ps = canvas..postscript(colormode='color')
          # img = Image.open(io.BytesIO(ps.encode('utf-8')))
          # img.save('filename.jpg', 'jpeg')

          # postscript_file = "tmp_snapshot.ps"
          # subprocess.call(["impg", "-window", "td", postscript_file])

        # WIP Button for Printing 'Canvas' and button position
        btn_print = Button(ap_analyze_window, text="Print as PDF", command=pdf_Print(),
                           width=20, height=3, fg='green')

        btn_print.place(x=100, y=550)

        # WIP Button for saving file as 'PDF' and button position
        btn_pdf = Button(ap_analyze_window, text="Save as PDF", command=lambda: [ap_summary_window.destroy(), main_window.deiconify()],
                         width=20, height=3, fg='green')
        btn_pdf.place(x=500, y=550)

        # Display's 'AP Analyze' GUI frame
        ap_analyze_window.mainloop()

    # Button event handler for calling 'analyze_ap_summary' function and button position
    btn_analyze = Button(ap_summary_window, text="Analyze File", command=anaylze_AP_summary,
                         width=20, height=3, fg='green')
    btn_analyze.place(x=700, y=550)

    # Displays 'AP Summary' GUI
    ap_summary_window.mainloop()


# Creates Button on main window for 'AP Summary' Option
btn_AP_Summary = Button(main_window, text="AP Summary", command=create_AP_Summary_frame,
                        width=40, height=3, fg='green')
btn_AP_Summary.place(x=50, y=100)

# @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ AR Summary Window @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
# Creates AR Summmary window


def create_AR_Summary_frame():

    # Creates GUI object and sets size of GUI
    ar_summary_window = Tk()
    ar_summary_window.geometry('1600x1200')

    # Hides 'Main Screen' GUI
    main_window.withdraw()

    # Sets title of 'AR Summary File Upload' GUI
    ar_summary_window.title('AR Summary File Upload')

    # Creates textbox for 'AR Summary Frame' GUI and sets location on GUI
    ar_txt_file = Text(ar_summary_window, height=4,
                       width=50, bg='white', fg='black', font=('Sans Serif', 23, 'italic bold'))
    ar_txt_file.place(x=100, y=350)

    # Button for going back to 'Pueblo Cooperative Care (Home)' GUI
    btn_home = Button(ar_summary_window, text="Back to Main", command=lambda: [ar_summary_window.destroy(), main_window.deiconify()],
                      width=20, height=3, fg='green')

    # Button for selecting 'AR Summary' file for upload and location on 'AR Summary Frame' GUI
    btn_home.place(x=100, y=550)
    btn_file_upload = Button(ar_summary_window, text="Find File", command=lambda: [ar_txt_file.configure(state='normal'), ar_txt_file.delete('1.0', "end"),
                                                                                   browseFiles(), ar_txt_file.insert(INSERT, filepath),
                                                                                   ar_txt_file.configure(state='disabled')], width=20, height=3, fg='green')
    btn_file_upload.place(x=400, y=550)

    # Nested fuction for ETL of 'AR Summary' .xslx file
    def anaylze_AR_summary():
        df = pd.read_excel(
            filepath, sheet_name=1)

        df['Unnamed: 0'] = np.nan
        df.dropna(how='all', axis=1, inplace=True)

        df = pd.melt(df, id_vars='Unnamed: 1',
                     var_name='Date Ranges', value_name='Amount')

        # Nested function for creating figure for 'grouped bar chart'
        def create_plot():
            axs = sns.catplot(x='Date Ranges', y='Amount',
                              hue='Unnamed: 1', data=df, kind='bar')
            axs.fig.suptitle('AR Summary')
            return axs.fig
        # Creates window for grARhical display of data and size of GUI frame
        ar_analyze_window = Tk()
        ar_analyze_window.geometry('1900x1200')
        ar_analyze_window.state('zoomed')
        ar_analyze_window.title("AR Summary Calculations")
        # Calls create_plot() function to put 'grouped bar chart' in 'Canvas' frame
        figure = create_plot()
        canvas = FigureCanvasTkAgg(figure, master=ar_analyze_window)
        canvas.draw()
        # Sets position of created 'Canvas'
        canvas.get_tk_widget().pack(side='top')
        canvas.get_tk_widget().place(bordermode=OUTSIDE)

        fig = plt.figure(figsize=(6, 6))
        fig.set_size_inches(9, 7)
        plt.axis('equal')

        pie_labels = (df['Unnamed: 1'][20:23])
        pie_data = (df['Amount'][20:23])

        colors = sns.color_palette('bright')[0:4]

        plt.title("AR Summary Totals")

        pie_graph = plt.pie(pie_data, labels=pie_labels, colors=colors,
                            autopct='%.0f%%', pctdistance=.5, explode=[0.1]*3, textprops={"fontsize": 11.9}, shadow=True)
        plt.legend()

        canvas = FigureCanvasTkAgg(fig, master=ar_analyze_window)
        canvas.draw()

        canvas.get_tk_widget().place(relx=.6, rely=.78, anchor=CENTER)
        canvas.get_tk_widget().pack(side='top')

        # Hides 'AR Summary GUI'
        ar_summary_window.withdraw()

        # WIP
        def pdf_Print():
            f = 5
          # fp = tempfile.TemporaryFile()
          # fp.write(ar_analyze_window)

          # ps = canvas..postscript(colormode='color')
          # img = Image.open(io.BytesIO(ps.encode('utf-8')))
          # img.save('filename.jpg', 'jpeg')

          # postscript_file = "tmp_snARshot.ps"
          # subprocess.call(["impg", "-window", "td", postscript_file])

        # WIP Button for Printing 'Canvas' and button position
        btn_print = Button(ar_analyze_window, text="Print as PDF", command=pdf_Print(),
                           width=20, height=3, fg='green')

        btn_print.place(x=100, y=550)

        # WIP Button for saving file as 'PDF' and button position
        btn_pdf = Button(ar_analyze_window, text="Save as PDF", command=lambda: [ar_summary_window.destroy(), main_window.deiconify()],
                         width=20, height=3, fg='green')
        btn_pdf.place(x=500, y=550)

        # Display's 'AR Analyze' GUI frame
        ar_analyze_window.mainloop()

    # Button event handler for calling 'analyze_AR_summary' function and button position
    btn_analyze = Button(ar_summary_window, text="Analyze File", command=anaylze_AR_summary,
                         width=20, height=3, fg='green')
    btn_analyze.place(x=700, y=550)

    # Displays 'AR Summary' GUI
    ar_summary_window.mainloop()


# Creates Button on main window for 'AR Summary' Option
btn_AR_Summary = Button(main_window, text="AR Summary", command=create_AR_Summary_frame,
                        width=40, height=3, fg='green')
btn_AR_Summary.place(x=50, y=200)

# @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ Assets Liabilites Window @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
# Creates 'Assets Liabilities' window


def create_assets_liabilities_Summary_frame():
    # Creates GUI object and sets size of GUI
    al_summary_window = Tk()
    al_summary_window.geometry('1600x1200')

    # Hides 'Main Screen' GUI
    main_window.withdraw()

    # Sets title of Assets Liabilities' File Upload' GUI
    al_summary_window.title('AL Summary File Upload')

    # Creates textbox for Assets Liabilities' GUI and sets location on GUI
    al_txt_file = Text(al_summary_window, height=4,
                       width=50, bg='white', fg='black', font=('Sans Serif', 23, 'italic bold'))
    al_txt_file.place(x=100, y=350)

    # Button for going back to 'Pueblo Cooperative Care (Home)' GUI
    btn_home = Button(al_summary_window, text="Back to Main", command=lambda: [al_summary_window.destroy(), main_window.deiconify()],
                      width=20, height=3, fg='green')

    btn_home.place(x=100, y=550)

    # Button for selecting Assets Liabilities' file for upload and location on Assets Liabilities' GUI
    btn_file_upload = Button(al_summary_window, text="Find File", command=lambda: [al_txt_file.configure(state='normal'), al_txt_file.delete('1.0', "end"),
                                                                                   browseFiles(), al_txt_file.insert(INSERT, filepath),
                                                                                   al_txt_file.configure(state='disabled')], width=20, height=3, fg='green')
    btn_file_upload.place(x=400, y=550)

    def analyze_AL_summary():
        df = pd.read_excel(
            filepath, sheet_name=1)

        # captured data points for total assets
        prev_year_total_assets = (df['Unnamed: 8'][47])
        current_year_total_assets = (df['Unnamed: 6'][47])

        # captured data points for total liabilities
        prev_year_total_liabilities = (df['Unnamed: 8'][68])
        current_year_total_liabilities = (df['Unnamed: 6'][68])

        # captured data points for total equity
        prev_year_total_equity = (df['Unnamed: 8'][78])
        current_year_total_equity = (df['Unnamed: 6'][78])

        # Creates GUI for 'Assets Liabilites' data
        al_analyze_window = Tk()
        al_analyze_window.geometry('1900x1200')
        al_analyze_window.title("Test")
        al_analyze_window.state('zoomed')

        # increases size of entire seaborn contents
        sns.set_theme(font_scale=1.5)

        # Create total assets bar graph
        def create_total_assets_bar():

            sns.set_palette(['orange', 'purple'])
            ta_axs = sns.catplot(x=['Jan 1, 22', 'Jan 1, 21'], y=[current_year_total_assets, prev_year_total_assets],
                                 data=df, kind='bar')
            ta_axs.fig.suptitle('Total Assets Summary')
            plt.xlabel("Year")
            plt.ylabel("$ Amount (In Millions) ")

            ax = ta_axs.axes[0, 0]
            plt.bar_label(ax.containers[0], fmt='$%.2f', padding=.01)
            ta_axs.fig.set_size_inches(8.5, 7)
            return ta_axs.fig

        ta_fig = create_total_assets_bar()
        canvas = FigureCanvasTkAgg(ta_fig, master=al_analyze_window)
        canvas.draw()

        canvas.get_tk_widget().place(relx=.17, rely=.5, anchor=CENTER)

        # creates total liabilities bar graph
        def create_total_liabilities_bar():
            sns.set_palette(['red', 'green'])
            tl_axs = sns.catplot(x=['Jan 1, 22', 'Jan 1, 21'], y=[current_year_total_liabilities, prev_year_total_liabilities],
                                 data=df, kind='bar')
            tl_axs.fig.suptitle('Total Liabilities Summary')
            plt.xlabel("Year")
            plt.ylabel("$ Amount")

            ax = tl_axs.axes[0, 0]
            plt.bar_label(ax.containers[0], fmt='$%.2f', padding=.01)
            tl_axs.fig.set_size_inches(8.5, 7)
            return tl_axs.fig

        tl_fig = create_total_liabilities_bar()
        canvas = FigureCanvasTkAgg(tl_fig, master=al_analyze_window)
        canvas.draw()

        canvas.get_tk_widget().place(relx=.5, rely=.5, anchor=CENTER)

        # create total equity bar graph
        def create_total_equity_bar():
            sns.set_palette(['teal', 'grey'])
            te_axs = sns.catplot(x=['Jan 1, 22', 'Jan 1, 21'], y=[current_year_total_equity, prev_year_total_equity],
                                 data=df, kind='bar')
            te_axs.fig.suptitle('Total Equity Summary')
            plt.xlabel("Year")
            plt.ylabel("$ Amount (In Millions) ")

            ax = te_axs.axes[0, 0]
            plt.bar_label(ax.containers[0], fmt='$%.2f', padding=.01)
            te_axs.fig.set_size_inches(8.5, 7)
            return te_axs.fig

        te_fig = create_total_equity_bar()
        canvas = FigureCanvasTkAgg(te_fig, master=al_analyze_window)
        canvas.draw()
        canvas.get_tk_widget().place(relx=.83, rely=.5, anchor=CENTER)

        al_analyze_window.mainloop()

    btn_analyze = Button(al_summary_window, text="Analyze File", command=analyze_AL_summary,
                         width=20, height=3, fg='green')
    btn_analyze.place(x=700, y=550)

    al_summary_window.mainloop()


# Creates Button on main window for 'Statement of Assets-Liabilities' Option
btn_Assets_Liabilities = Button(main_window, text="Statement of Assets-Liabilities",
                                command=create_assets_liabilities_Summary_frame, width=40, height=3, fg='green')
btn_Assets_Liabilities.place(x=50, y=300)

# Creates 'Revenue and Expenses' window

# @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ Revenue Expenses Window @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


def create_Revenue_Expenses_frame():

    revenues_expsenses_window = Tk()
    revenues_expsenses_window.geometry('1600x1200')

    # Hides 'Main Screen' GUI
    main_window.withdraw()

    # Sets title of 'Revenue and Expenses' GUI
    revenues_expsenses_window.title('AR Summary File Upload')

    # Creates textbox for 'Revenue and Expenses' GUI and sets location on GUI
    re_txt_file = Text(revenues_expsenses_window, height=4,
                       width=50, bg='white', fg='black', font=('Sans Serif', 23, 'italic bold'))
    re_txt_file.place(x=100, y=350)

    # Button for going back to 'Pueblo Cooperative Care (Home)' GUI
    btn_home = Button(revenues_expsenses_window, text="Back to Main", command=lambda: [revenues_expsenses_window.destroy(), main_window.deiconify()],
                      width=20, height=3, fg='green')

    # Button for selecting 'Revenue and Expenses' file for upload and location on ''Revenue and Expenses' GUI
    btn_home.place(x=100, y=550)
    btn_file_upload = Button(revenues_expsenses_window, text="Find File", command=lambda: [re_txt_file.configure(state='normal'), re_txt_file.delete('1.0', "end"),
                                                                                           browseFiles(), re_txt_file.insert(INSERT, filepath),
                                                                                           re_txt_file.configure(state='disabled')], width=20, height=3, fg='green')
    btn_file_upload.place(x=400, y=550)

    revenues_expsenses_window.mainloop()


# Creates Button on main window for 'Statement Revenue and Expenses' Option
btn_Revenue_Expenses = Button(main_window, text="Statement of Revenue and Expenses",
                              command=create_Revenue_Expenses_frame, width=40, height=3, fg='green')
btn_Revenue_Expenses.place(x=50, y=400)


# @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ Revenue and Expenses Comparison Window @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
# Creates 'Revenue and Expenses Comparison' window

def create_Revenue_Expenses_Comparison_frame():

    revenues_expsenses_comparison_window = Tk()
    revenues_expsenses_comparison_window .geometry('1600x1200')

    # Hides 'Main Screen' GUI
    main_window.withdraw()

    # Sets title of 'Revenue and Expenses Comparison' GUI
    revenues_expsenses_comparison_window.title('AR Summary File Upload')

    # Creates textbox for 'Revenue and Expenses Comparison' GUI and sets location on GUI
    re_C_txt_file = Text(revenues_expsenses_comparison_window, height=4,
                         width=50, bg='white', fg='black', font=('Sans Serif', 23, 'italic bold'))
    re_C_txt_file.place(x=100, y=350)

    # Button for going back to 'Pueblo Cooperative Care (Home)' GUI
    btn_home = Button(revenues_expsenses_comparison_window, text="Back to Main", command=lambda: [revenues_expsenses_comparison_window.destroy(), main_window.deiconify()],
                      width=20, height=3, fg='green')

    # Button for selecting 'Revenue and Expenses Comparison' file for upload and location on 'Revenue and Expenses Comparison' GUI
    btn_home.place(x=100, y=550)
    btn_file_upload = Button(revenues_expsenses_comparison_window, text="Find File", command=lambda: [re_C_txt_file.configure(state='normal'), re_C_txt_file.delete('1.0', "end"),
                                                                                                      browseFiles(), re_C_txt_file.insert(INSERT, filepath),
                                                                                                      re_C_txt_file.configure(state='disabled')], width=20, height=3, fg='green')
    btn_file_upload.place(x=400, y=550)

    revenues_expsenses_comparison_window.mainloop()


# Creates Button on main window for 'Statement of Assets-Liabilities' Option
btn_Revenue_Expenses_Comparison = Button(main_window, text="Statement of Revenue and Expenses Comparison",
                                         command=create_Revenue_Expenses_Comparison_frame, width=40, height=3, fg='green')
btn_Revenue_Expenses_Comparison.place(x=50, y=500)

# Creates 'Revenue and Expenses Compared to Budget' window

# @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ Revenue Expenses Budget Window @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


def create_Revenue_Expenses_Budget_frame():

    revenues_expsenses_budget_window = Tk()
    revenues_expsenses_budget_window .geometry('1600x1200')

   # Hides 'Main Screen' GUI
    main_window.withdraw()

    # Sets title of 'Revenue and Expenses Compared to Budget' GUI
    revenues_expsenses_budget_window.title('AR Summary File Upload')

    # Creates textbox for 'Revenue and Expenses Compared to Budget' GUI and sets location on GUI
    re_B_txt_file = Text(revenues_expsenses_budget_window, height=4,
                         width=50, bg='white', fg='black', font=('Sans Serif', 23, 'italic bold'))
    re_B_txt_file.place(x=100, y=350)

    # Button for going back to 'Pueblo Cooperative Care (Home)' GUI
    btn_home = Button(revenues_expsenses_budget_window, text="Back to Main", command=lambda: [revenues_expsenses_budget_window.destroy(), main_window.deiconify()],
                      width=20, height=3, fg='green')

    # Button for selecting 'Revenue and Expenses Compared to Budget' file for upload and location on 'Revenue and Expenses Compared to Budget' GUI
    btn_home.place(x=100, y=550)
    btn_file_upload = Button(revenues_expsenses_budget_window, text="Find File", command=lambda: [re_B_txt_file.configure(state='normal'), re_B_txt_file.delete('1.0', "end"),
                                                                                                  browseFiles(), re_B_txt_file.insert(INSERT, filepath),
                                                                                                  re_B_txt_file.configure(state='disabled')], width=20, height=3, fg='green')
    btn_file_upload.place(x=400, y=550)

    revenues_expsenses_budget_window.mainloop()


# Creates Button on main window for 'Statement of Revenue and Expense Budget' Option
btn_Revenue_Expenses_Budget = Button(main_window, text="Statement of Revenue and Expenses Budget",
                                     command=create_Revenue_Expenses_Budget_frame, width=40, height=3, fg='green')
btn_Revenue_Expenses_Budget.place(x=50, y=600)


# @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ Revenue Expnese by Program Window @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
# Creates 'Revenue and Expenses by Program' window

def create_Revenue_Expenses_Program_frame():

    revenues_expsenses_program_window = Tk()
    revenues_expsenses_program_window .geometry('1600x1200')

    # Hides 'Main Screen' GUI
    main_window.withdraw()

    # Sets title of 'Revenue and Expenses by Program' GUI
    revenues_expsenses_program_window.title('AR Summary File Upload')

    # Creates textbox for 'Revenue and Expenses by Program' GUI and sets location on GUI
    re_P_txt_file = Text(revenues_expsenses_program_window, height=4,
                         width=50, bg='white', fg='black', font=('Sans Serif', 23, 'italic bold'))
    re_P_txt_file.place(x=100, y=350)

    # Button for going back to 'Pueblo Cooperative Care (Home)' GUI
    btn_home = Button(revenues_expsenses_program_window, text="Back to Main", command=lambda: [revenues_expsenses_program_window.destroy(), main_window.deiconify()],
                      width=20, height=3, fg='green')

    # Button for selecting 'Revenue and Expenses by Program' file for upload and location on 'AR Summary Frame' GUI
    btn_home.place(x=100, y=550)
    btn_file_upload = Button(revenues_expsenses_program_window, text="Find File", command=lambda: [re_P_txt_file.configure(state='normal'), re_P_txt_file.delete('1.0', "end"),
                                                                                                   browseFiles(), re_P_txt_file.insert(INSERT, filepath),
                                                                                                   re_P_txt_file.configure(state='disabled')], width=20, height=3, fg='green')
    btn_file_upload.place(x=400, y=550)

    revenues_expsenses_program_window.mainloop()


# Creates Button on main window for 'Statement oof Revenue and Expenses by Program' Option
btn_Revenue_Expenses_Program = Button(main_window, text="Statement of Revenue and Expenses by Program",
                                      command=create_Revenue_Expenses_Program_frame, width=40, height=3, fg='green')
btn_Revenue_Expenses_Program.place(x=50, y=700)


# @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ Main Window @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
# Exit button to close program
btn_exit = Button(main_window, text="Exit",
                  command=main_window.destroy, width=40, height=3, fg='green')
btn_exit.place(x=1150, y=900)

# loads Pueblo Cooperative Care image
# Will need to be changed so can be accessed globally
load = Image.open(
    "Pueblo-Coop-Center-History.gif")
render = ImageTk.PhotoImage((load))
photo = Label(main_window, image=render)
photo.place(x=600, y=100)


main_window.mainloop()
