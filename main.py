import xlwings as xw
import tkinter as tk
import pandas as pd

def tkinter_window():
    root = tk.Tk()
    root.title("Part Matcher")
    root.geometry("300x200")

    root.mainloop()

def match_descriptions():
    #wb = xw.Book.caller()
    wb = xw.Book('demo.xlsm')
    df = pd.read_excel(wb.fullname, sheet_name='Sheet1',skiprows=11, header=0)
    customer_description = df['Customer Description'].tolist()
    
    #for desc in customer_description:
    

if __name__ == "__main__":
    match_descriptions()