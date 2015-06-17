# JD_Expt_Lbe_Tool_2
Tool to pull data from EM database and provide user friendly GUI to automatically get results both in 
excel form and statistically.
#Consists of 4 files
  #Exempt_Lable_Tool
  +This hold the framework of the applicaiton calling classes such as the Weekly, Monthly and Quarterly generators.
    Written with the tkinter framework and easy to mantain features, you can edit this at leisure
  #Monthly_Report
    Written with pyodbc to automatically querry data from a SQL server and carry out individualized querrying for a specific
    field for monthly data. Openpyxl and pyodbc were used in this process
  #Weekly Report
    Written with similar functionality to the monthly class but simply in nature to querry for weekly data.
    

