"""
Initialization of directory path for files
"""
directory = r"C:\Users\ryana\Desktop\FolderForCode\TEST"

"""
Allows for iteration of multiple files within folder. References the prior named directory.
"""
name = input("Name: ")
tmonth = input("Month: ")
t1month = input("Month + Year: ")

for filename in os.listdir(directory):
    if filename.endswith(".xlsx"):
        timesheet_wb = load_workbook(f"./TEST/TS-2022_{name}.xlsx")
        ws = timesheet_wb.active   # loading excel sheet
        
        timesheet_month = tmonth    # Haley will enter the timesheet month here
        outputsheet_month = t1month    # Haley must enter this too, format is: {MONTH 22}
        month = timesheet_wb[timesheet_month]     # determines what sheet to work on 
        
        output_file = f'OT_template{name}.xlsx'
        output_workbook = load_workbook(f"./TEST/OT_template{name}.xlsx")
        output_ws = output_workbook[outputsheet_month]    # loading in the month in the output file
        output_ws = output_workbook[outputsheet_month]
print("Done")