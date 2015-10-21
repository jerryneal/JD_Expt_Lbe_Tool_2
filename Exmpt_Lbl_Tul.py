'''
    Title : Exmpt_Lbl_Tul
    Version: 1.0
    Author: Neal Bazirake
    Description: Application Framework 
'''
import os
from tkinter import *
from tkinter import filedialog #fix an error with cx_Freeze, requires import of this tkinter function specifically
from tkinter import ttk #needed to make combobox
import pyodbc #used to establish a connetion to the SQL server and make queries
import sys
from Weekly_Report import * # Contains the weekly report functions.
from Monthly_Report import *
import subprocess #sys and subprocess allow me to send arguments to programs I open
from tkinter import messagebox #used to allow tkinter pop up windows
from datetime import datetime, timedelta

class Application(Frame):
    def __init__(self, master=None):
        #this sets up the main window/frame and names it "master"
        Frame.__init__(self, master)
        self.grid()
        self.master.title("Exempt Label Tool")

        def WeeklyButton(text):
            '''
            Weekly Report
            '''

            def weekly_report():
                fromdateText = ""+fromdateYear.get()+"-"+fromdateMonth.get()+"-"+fromdateDay.get()
                todateText = ""+todateYear.get()+"-"+todateMonth.get()+"-"+todateDay.get()
                weekly_generator(fromdateText,todateText)
                
                
            text.set("")
            #This "Clears" the last frame but putting a new frame on top of the last. covering everything from the last frame
            FrameClear = Frame(master, bg = "#F5F5DC").grid(row = 1, column = 0, columnspan=10, rowspan=11)
            FrameClearText = Label(FrameClear, text=" "*240+"\n"*25, bg = "#F5F5DC").grid(row = 1, column = 0, columnspan=10, rowspan=11)
                
            #Once the last frame is covered, I create a new one that I want displayed overtop of the "Clear" frame.

            requiredfromdateformat = Label(Frame2, text="MM-DD-YYYY", bg='#F5F5DC', font=("Helvetica",10), fg='#000000', justify=CENTER).grid(row=2,column=3,columnspan=4)
            requiredtodateformat = Label(Frame2, text="MM-DD-YYYY", bg='#F5F5DC', font=("Helvetica",10), fg='#000000', justify=CENTER).grid(row=3,column=3,columnspan=4)

     
   
            Heading = Label(Frame2, text = "\nWeekly Report Inputs\n", bg='#F5F5DC', font=("Helvetica", 16), fg='#000000', justify=CENTER).grid(row=1,columnspan=10,sticky=N+W+E)
            fromdate = Label(Frame2, text = "Start Date", bg = "#F5F5DC", fg='#000000', font=("Helvetica", 12)).grid(row = 2, column = 1)
            todate = Label(Frame2, text = "End Date", bg = "#F5F5DC", fg='#000000', font=("Helvetica", 12)).grid(row = 3, column = 1)
             
            #Drop downs
            fromdateMonth = ttk.Combobox()
            fromdateMonth.grid(row = 2, column = 2, columnspan =1)
            fromdateMonth['values'] = [' ','01','02','03','04','05','06','07','08','09','10','11','12']
            fromdateMonth['width'] = ['4']
            fromdateDay = ttk.Combobox()
            fromdateDay.grid(row = 2, column = 2, columnspan =2)
            fromdateDay['values'] = [' ', '01','02','03','04','05','06','07','08','09','10','11','12','13','14','15','16','17','18','19','20','21','22','23','24','25','26','27','28','29','30','31']
            fromdateDay['width'] = ['2']
            fromdateYear = ttk.Combobox()
            fromdateYear.grid(row = 2, column = 2, columnspan =3)
            fromdateYear['values'] = [' ', '2004', '2005', '2006','2007','2008','2009','2010','2011','2012','2013','2014','2015','2016','2017']
            fromdateYear['width'] = ['4']
            
            todateMonth = ttk.Combobox()
            todateMonth.grid(row = 3, column = 2, columnspan =1)
            todateMonth['values'] = [' ', '02','03','04','05','06','07','08','09','10','11','12']
            todateMonth['width'] = ['4']
            todateDay = ttk.Combobox()
            todateDay.grid(row = 3, column = 2, columnspan =2)
            todateDay['values'] = [' ', '01','02','03','04','05','06','07','08','09','10','11','12','13','14','15','16','17','18','19','20','21','22','23','24','25','26','27','28','29','30','31']
            todateDay['width'] = ['2']
            todateYear = ttk.Combobox()
            todateYear.grid(row = 3, column = 2, columnspan =3)
            todateYear['values'] = [' ', '2004', '2005', '2006','2007','2008','2009','2010','2011','2012','2013','2014','2015','2016','2017']
            todateYear['width'] = ['4']

            
            #The run button is supposed to push the weekly_report query
            run = Button(Frame2, text="Run Weekly", width=15, height=1, command = weekly_report).grid(row= 4, column=0, columnspan=4)

            

        def MonthlyButton(text):
            '''
            Monthly Report
            '''

            def monthly_report():
                m = str(combinedMonth['values'].index(combinedMonth.get()))
                
                if combinedMonth.get() == "February":
                    startdate = ""+combinedYear.get()+"-"+m+"-01"
                    endate = ""+combinedYear.get()+"-"+m+"-28"
                    
                elif combinedMonth.get() == "Septempber" or combinedMonth.get() == "April" or combinedMonth.get() == "June" or combinedMonth.get() == "November":
                    startdate = ""+combinedYear.get()+"-"+m+"-01"
                    endate = ""+combinedYear.get()+"-"+m+"-30"
                    
                else:
                    startdate = ""+combinedYear.get()+"-"+m+"-01"
                    endate = ""+combinedYear.get()+"-"+m+"-31"
                        
                    
                #print(combinedYear.get(),combinedMonth.get())
                caty = Monthly()
                caty.monthly_queries(startdate,endate)
            
            text.set("")
            #This "Clears" the last frame but putting a new frame on top of the last. covering everything from the last frame
            FrameClear = Frame(master, bg = "#F5F5DC").grid(row = 1, column = 0, columnspan=10, rowspan=11)
            FrameClearText = Label(FrameClear, text=" "*240+"\n"*25, bg = "#F5F5DC").grid(row = 1, column = 0, columnspan=10, rowspan=11)
                
            #Once the last frame is covered, I create a new one that I want displayed overtop of the "Clear" frame.
                
            Heading = Label(Frame2, text = "\nMonthly Report Maker\n", bg='#F5F5DC', font=("Helvetica", 16), fg='#000000', justify=CENTER).grid(row=1,columnspan=10,sticky=N+W+E)

            Year = Label(Frame2, text=" " + "Select Year", bg = "#F5F5DC", fg='#000000', font=("Helvetica", 12)).grid(row = 2, column = 1)
            combinedYear = ttk.Combobox()
            combinedYear.grid(row=2, column=1, columnspan=3)
            combinedYear['values'] = [' ', '2004', '2005', '2006','2007','2008','2009','2010','2011','2012','2013','2014','2015','2016']
            combinedYear['width'] = ['5'] 
            #combinedYear.bind('<<ComboboxSelected>>', year)

            Month = Label(Frame2, text=" " + "Select Month", bg = "#F5F5DC", fg='#000000', font=("Helvetica", 12)).grid(row = 3, column = 1)
            combinedMonth = ttk.Combobox()
            combinedMonth.grid(row=3, column=1, columnspan=3)
            combinedMonth['values'] = [' ', 'January','February','March','April','May','June','July','August','September','October','November','December']
            combinedMonth['width'] = ['10'] 
            #combinedMonth.bind('<<ComboboxSelected>>', month)
            
            run = Button(Frame2, text="Run Monthly", width=15, height=1, command = monthly_report).grid(row= 4, column=0, columnspan=4)

            
        def QuarterlyButton(text):
            '''
            Quarterly Report
            '''

            def quarterly_report():
                #Under maintenance for now so notification feature to be included here.
                pass
            
            text.set("")
            #This "Clears" the last frame but putting a new frame on top of the last. covering everything from the last frame
            FrameClear = Frame(master, bg = "#F5F5DC").grid(row = 1, column = 0, columnspan=10, rowspan=11)
            FrameClearText = Label(FrameClear, text=" "*240+"\n"*25, bg = "#F5F5DC").grid(row = 1, column = 0, columnspan=10, rowspan=11)
                
            #Once the last frame is covered, I create a new one that I want displayed overtop of the "Clear" frame.
                
            Heading = Label(Frame2, text = "\n Quarterly Report\n", bg='#F5F5DC', font=("Helvetica", 16), fg='#000000', justify=CENTER).grid(row=1,columnspan=10,sticky=N+W+E)
            Year = Label(Frame2, text=" " + "Select Year", bg = "#F5F5DC", fg='#000000', font=("Helvetica", 12)).grid(row = 2, column = 1)
            combinedYear = ttk.Combobox()
            combinedYear.grid(row=2, column=1, columnspan=3)
            combinedYear['values'] = [' ', '2004', '2005', '2006','2007','2008','2009','2010','2011','2012','2013','2014','2015','2016']
            combinedYear['width'] = ['5'] 
            #cmbYear.bind('<<ComboboxSelected>>', year)

            Quarter = Label(Frame2, text=" " + "Select Quarter", bg = "#F5F5DC", fg='#000000', font=("Helvetica", 12)).grid(row = 3, column = 1)
            cmbQuarter = ttk.Combobox()
            cmbQuarter.grid(row=3, column=1, columnspan=3)
            cmbQuarter['values'] = [' ', '1','2','3','4']
            cmbQuarter['width'] = ['2'] 
            #cmbMonth.bind('<<ComboboxSelected>>', month)

            run = Button(Frame2, text="Run Quarterly", width=15, height=1, command = quarterly_report).grid(row= 4, column=0, columnspan=4)

       
        def UserReportButton(text):
            '''
            #Exempt Report per Owner
            '''

            def user_report():
                #Under maintenance for now and notification feature to included here.
                pass
            
            text.set("")
            #This "Clears" the last frame but putting a new frame on top of the last. covering everything from the last frame
            FrameClear = Frame(master, bg = "#F5F5DC").grid(row = 1, column = 0, columnspan=10, rowspan=11)
            FrameClearText = Label(FrameClear, text=" "*240+"\n"*25, bg = "#F5F5DC").grid(row = 1, column = 0, columnspan=10, rowspan=11)

            #Once the last frame is covered, I create a new one that I want displayed overtop of the "Clear" frame.
            
            Heading = Label(Frame2, text = "\nStats per User\n", bg='#F5F5DC', font=("Helvetica", 16), fg='#000000', justify=CENTER).grid(row=1,columnspan=10,sticky=N+W+E)
            
            user = Label(Frame2, text = "Enter Racf ID", bg = "#F5F5DC", fg='#000000', font=("Helvetica", 12)).grid(row = 2, column = 1)
            user = Text(Frame1, width=11, height=1)
            user.grid(row = 2, column = 1, columnspan =3)
            #fromdateText.bind("<Return>", user_report)

            run = Button(Frame2, text="Run Report", width=15, height=1, command = user_report).grid(row= 3, column=0, columnspan=3)

        def QuickStatsButton(text):
            '''
            #Quick Stats
            1. Displays number of EPA 1068.210 engines
            2. Displays number of EPA 1068.215 engines
            3. Displays number of # Exemption Country1  Engines
            4. Displays number of # Exemption Country2
            '''
            pass
        
        def ResetButton(text):
            #Reset button
            pass
        
        Frame0 = Frame(master, bg = "#F5F5DC").grid(row = 1, column = 0, columnspan=10, sticky=N+W)
        Frame1 = Frame(master, bg = "#F5F5DC").grid(row = 3, column = 0, columnspan=10, sticky=E+W)
        Frame2 = Frame(master, bg = "#F5F5DC").grid(row = 3, column = 0, columnspan=10, sticky=E+W)

        Button(Frame0, text="Weekly Report", width = 20, command=lambda: WeeklyButton(text)).grid(row=0,column=0)
        Button(Frame0, text="Monthly Report", width = 20, command=lambda: MonthlyButton(text)).grid(row=0,column=1)
        Button(Frame0, text="Quarterly Rport", width = 20, command=lambda: QuarterlyButton(text)).grid(row=0,column=2)
        Button(Frame0, text="Exempt Report per Holder", width = 20, command=lambda: UserReportButton(text)).grid(row=0,column=3)
        Button(Frame0, text="Quick Starts", width = 20, command=lambda: QuickStatsButton(text)).grid(row=0,column=4)
        Button(Frame0, text="Reset Button", width = 20, command=lambda: ResetButton(text)).grid(row=0,column=5)
        
        
        text = StringVar()
        text.set("\nWelcome to The Exempt Label tool!\n\nPlease select which tool you would like to use at the top of the screen. " +
                 "\nThis is an administrator tool to make report generation much easier\n")
        
        Label(Frame1, textvariable= text, justify = LEFT, bg = "#F5F5DC", fg='#000000', font=("Helvetica", 10)).grid(row=2,column=0, columnspan=4, sticky=W)
        
        Label(Frame1, text= "\n\n", justify = LEFT, bg = "#F5F5DC", fg='#000000').grid(row=3,column=0)

root = Tk()
root.geometry("735x320")
root.config(bg='#F5F5DC')
app = Application(master=root)
app.mainloop()

