from os import error
import os
import pandas as pd
import numpy as np
import yfinance as yf
from tkinter import *
from tkinter import ttk
from tkinter import filedialog 
from tkinter import messagebox
import traceback
import warnings

warnings.filterwarnings("ignore")


def stock_looper(stocks, period, interval):
  stock_dfs = []
  i = 0
  for ticker in stocks:
    try:
      #create path for saving data
      stock_dir = "/"+ticker
      save_path = save_var.get()+stock_dir

      if os.path.exists(save_path):
        pass
      else:
        os.makedirs(save_path)

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

      #get data for the wicked wango
      df["pre HH"] = df["HH"].shift(periods=-1)
      df["pre LL"] = df["LL"].shift(periods=-1)

   
      #golden goose
      df["GOLDEN GOOSE"] = df.apply(lambda row: "GOLDEN GOOSE" if (len(row["HH"]) >1) and (len(row["LL"]) > 1) else "", axis=1)

      #wicked wango
      df["WICKED WANGO"] = df.apply(lambda row: "WICKED WANGO" if ((row["HH"] == "Higher High") and (row["pre LL"] == "Lower Low")) or ((row["LL"] == "Lower Low") and (row["pre HH"] == "Higher High")) else "", axis=1)

      #get average time between wicked wangos
      df_ww = df[df["WICKED WANGO"] == "WICKED WANGO"]
      df_ww.reset_index(level=0, inplace=True)
      if (interval=="1m" or interval=="5m" or interval=="15m" or interval=="60m"):
        df_ww["date_prev_ww"] = df_ww["Datetime"].shift(periods=1)
        df_ww["time_diff"] = (df_ww["Datetime"] - df_ww["date_prev_ww"]).dt.days
      else:
        df_ww["date_prev_ww"] = df_ww["Date"].shift(periods=1)
        df_ww["time_diff"] = (df_ww["Date"] - df_ww["date_prev_ww"]).dt.days
      
      
      average_time_ww = round(np.mean(df_ww["time_diff"]), 2)

      #get average time between golden goose
      df_gg = df[df["GOLDEN GOOSE"] == "GOLDEN GOOSE"]
      df_gg.reset_index(level=0, inplace=True)
      if (interval=="1m" or interval=="5m" or interval=="15m" or interval=="60m"):
        df_gg["date_prev_gg"] = df_gg["Datetime"].shift(periods=1)
        df_gg["time_diff"] = (df_gg["Datetime"] - df_gg["date_prev_gg"]).dt.days
      else:
        df_gg["date_prev_gg"] = df_gg["Date"].shift(periods=1)
        df_gg["time_diff"] = (df_gg["Date"] - df_gg["date_prev_gg"]).dt.days
      
      
      average_time_gg = round(np.mean(df_gg["time_diff"]), 2)
      

      
    
      
      


      


      #drop the calculator columns
      df.drop(columns=["H day -1", "H day +1", "H day -2", "L day -1", "L day +1", "L day -2", "pre HH", "pre LL"], inplace=True)
  
      #reset index to avoid timezone error
      if (interval=="1m" or interval=="5m" or interval=="15m" or interval=="60m"):
        df.reset_index(level=0, inplace=True)
        df["Datetime"] = df["Datetime"].dt.tz_localize(tz=None)
        df.set_index('Datetime', drop=True, inplace=True)
      
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
  


  

    writer = pd.ExcelWriter(save_path+'/Stock overview {} - {} - {}.xlsx'.format(ticker, period, interval), engine="xlsxwriter")
    

    
    try:
      
      df.to_excel(writer, sheet_name=ticker)

      #Get the xlsxwriter workbook and worksheet object
      workbook = writer.book
      worksheet = writer.sheets[ticker]
      columns = df.columns.insert(0, "Date")

      column_settings = [{'header': column} for column in columns]

      (max_row, max_col) = df.shape

      worksheet.add_table(0, 0, max_row, max_col, {'columns': column_settings})

      worksheet.write("L1", "Average time between Wicked Wangos = {} days".format(average_time_ww))
      worksheet.write("L2", "Average time between Golden Geese = {} days".format(average_time_gg))

      #formats for highlighting
      green = workbook.add_format({'bold': 1, "bg_color": '#C6EFCE', "font_color": '#006100'})

      yellow = workbook.add_format({'bold': 1, "bg_color": '#FFEB9C', "font_color": '#9C6500'})

      red = workbook.add_format({'bold': 1, "bg_color": '#FFC7CE', "font_color": '#9C0006'})

      orange = workbook.add_format({'bold': 1, 'bg_color': '#ffcc66', 'font_color': '#cc3300'})

      #apply formatting
      worksheet.set_column('A:A', 18)
      worksheet.set_column('B:O', 15)

      worksheet.conditional_format("G2:G{}".format(max_row), {"type": "cell", "criteria": "equal to", "value": '"Higher High"', "format": green})

      worksheet.conditional_format("H2:H{}".format(max_row), {"type": "cell", "criteria": "equal to", "value": '"Lower Low"', "format": red})

      worksheet.conditional_format("I2:I{}".format(max_row), {"type": "cell", "criteria": "equal to", "value": '"GOLDEN GOOSE"', "format": yellow})

      worksheet.conditional_format("J2:J{}".format(max_row), {"type": "cell", "criteria": "equal to", "value": '"WICKED WANGO"', "format": orange})

      writer.save()

    except(error):
      print("an error occurred when writing the worksheet. Moving on to next stock")
      print(error)
      pass
  
  
  

def get_day_week_month():
  stocks = stock_var.get().upper().replace(" ", "").split(',')
  print("Getting 1 minute data")
  stock_looper(stocks, "7d", "1m")
  print("Getting 5 minute data")
  stock_looper(stocks, "1mo", "5m")
  print("Getting 15 minute data")
  stock_looper(stocks, "1mo", "15m")
  print("Getting 60 minute data")
  stock_looper(stocks, "2y", "60m")
  print("Getting 1 day data")
  stock_looper(stocks, "max", "1d")
  print("Getting 1 week data")
  stock_looper(stocks, "max", "1wk")
  print("Getting 1 month data")
  stock_looper(stocks, "max", "1mo")
  print("all done!")

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




  



