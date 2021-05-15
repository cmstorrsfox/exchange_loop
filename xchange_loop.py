from os import error
import pandas as pd
import numpy as np
import yfinance as yf
from tkinter import *
from tkinter import ttk
from tkinter import filedialog 
from tkinter import messagebox
import traceback


def stock_looper(stocks, period, interval):
  stock_dfs = []
  i = 0
  for ticker in stocks:
    try:
      i +=1
      print("processing {} stock data (stock {}/{})".format(ticker, i, len(stocks)))
      ticker_data = yf.Ticker(ticker)
      df = ticker_data.history(period=period, interval=interval)

      #drop unnecessary columns
      df.drop(columns=["Dividends", "Stock Splits"], inplace=True)
      
      #add cols to calculate Nick rules
      df["H day -1"] = df["High"].shift(periods=1)
      df["H day +1"] = df["High"].shift(periods=-1)
      df["H day -2"] = df["High"].shift(periods=2)
      df["L day -1"] = df["Low"].shift(periods=1)
      df["L day +1"] = df["Low"].shift(periods=-1)
      df["L day -2"] = df["Low"].shift(periods=2)

      #function to get higher high
      def hh(row):
        if (row["High"]>row["H day -1"]) and (row["High"] > row["H day +1"]) and (row["H day -1"] > row["H day -2"]):
          val="Higher High"
        else:
          val=""
      
        return val

      #function to get lower low
      def ll(row):
        if (row["Low"] < row["L day -1"]) and (row["Low"] < row["L day +1"]) and (row["L day -1"] < row["L day -2"]):
          val="Lower Low"
        else:
          val=""
      
        return val

      #create columns that shows whether or not a day is a match
      df["HH"] = df.apply(hh, axis=1)
      df["LL"] = df.apply(ll, axis=1)

      #golden goose
      df["GOLDEN GOOSE"] = df.apply(lambda row: "Golden Goose" if (len(row["HH"]) >1) and (len(row["LL"]) > 1) else "", axis=1)
      

      #drop the calculator columns
      df.drop(columns=["H day -1", "H day +1", "H day -2", "L day -1", "L day +1", "L day -2"], inplace=True)
      
      #save df and ticker to tuple
      ticker_df_tuple = (ticker, df)

      #append to array of dfs for later loop
      stock_dfs.append(ticker_df_tuple)
      if i < len(stocks):
        print("done - moving on to next stock")
      else:
        pass
    
    except(error):
      print("There is a problem with this ticker symbol. Moving on to the next stock")
      print(error)
      pass
  

  
  writer = pd.ExcelWriter(save_var.get()+'/Stock overview - {} - {}.xlsx'.format(period, interval), engine="xlsxwriter")
  
  if (interval=="1m" or interval=="5m" or interval=="15m" or interval=="60m"):
    df.reset_index(level=0, inplace=True)
    df["Datetime"] = df["Datetime"].dt.tz_localize(tz=None)
    df.set_index('Datetime', drop=True, inplace=True)
  
  for ticker, df in stock_dfs:
    try:
      df.to_excel(writer, sheet_name=ticker)

      #Get the xlsxwriter workbook and worksheet object
      workbook = writer.book
      worksheet = writer.sheets[ticker]
      columns = df.columns.insert(0, "Date")

      column_settings = [{'header': column} for column in columns]

      (max_row, max_col) = df.shape

      worksheet.add_table(0, 0, max_row, max_col, {'columns': column_settings})

      #formats for highlighting
      green = workbook.add_format({'bold': 1, "bg_color": '#C6EFCE', "font_color": '#006100'})

      yellow = workbook.add_format({'bold': 1, "bg_color": '#FFEB9C', "font_color": '#9C6500'})

      red = workbook.add_format({'bold': 1, "bg_color": '#FFC7CE', "font_color": '#9C0006'})

      title = workbook.add_format({"bold": 1, "font_size": 18})

      #apply formatting
      worksheet.set_column('A:A', 18)
      worksheet.set_column('B:O', 15)

      worksheet.conditional_format("G2:M{}".format(max_row), {"type": "cell", "criteria": "equal to", "value": '"Higher High"', "format": green})

      worksheet.conditional_format("H2:N{}".format(max_row), {"type": "cell", "criteria": "equal to", "value": '"Lower Low"', "format": red})

      worksheet.conditional_format("I2:O{}".format(max_row), {"type": "cell", "criteria": "equal to", "value": '"GOLDEN GOOSE"', "format": yellow})

    except(error):
      print("an error occurred when writing the worksheet. Moving on to next stock")
      print(error)
      pass
  writer.save()
  print("all done!")
  

def get_day_week_month():
  stocks = stock_var.get().replace(" ", "").split(',')
  print(stocks)
  stock_looper(stocks, "7d", "1m")
  stock_looper(stocks, "1mo", "5m")
  stock_looper(stocks, "1mo", "15m")
  stock_looper(stocks, "2y", "60m")
  stock_looper(stocks, "max", "1d")
  stock_looper(stocks, "max", "1wk")
  stock_looper(stocks, "max", "1mo")

def browse():
    save_loc = filedialog.askdirectory()
    save_var.set(save_loc)

#error callback
def show_error(self, *args):
  err = traceback.format_exception(*args)
  messagebox.showerror("Exception", err)

Tk.report_callback_exception = show_error

#start tkinter
root = Tk()
root.title("Stock Checker")
root.geometry("500x500")

#tkinter variables
stock_var = StringVar()
save_var = StringVar()

#elements
input = ttk.Frame(root)
instructions = ttk.Label(input, text="Enter the stocks you want daily, weekly and monthly data on in the box below separating them with a ','", wraplength=500, font=("Arial", 14, "bold"), justify="center")
stock_entry = ttk.Entry(input, textvariable=stock_var, width=50)
save_location = ttk.Entry(input, textvariable=save_var, width=50)
browse_btn = ttk.Button(input, text="Browse", command=browse)
go_btn = ttk.Button(input, text="Go", command=get_day_week_month)


#layout
input.grid(column=0, row=1)
instructions.grid(column=0, row=0, columnspan=3, padx=10, pady=5)
stock_entry.grid(column=1, row=2, pady=5)
save_location.grid(column=1, row=3, pady=5)
browse_btn.grid(column=1, row=4, pady=5)
go_btn.grid(column=1, row=5, pady=5)



root.mainloop()




  



