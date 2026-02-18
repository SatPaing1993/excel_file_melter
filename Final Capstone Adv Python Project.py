#import library
import pandas as pd
pd.set_option('display.max_columns', None)
import os

#input dir
input_dir = r"G:\My Drive\Data"

#file list
files = os.listdir(input_dir)

all_df = []

for file in files:
    #read excel file
    df = pd.read_excel(input_dir + "/" + file, header=None)
    #find file name
    file_name = str(df.iloc[1, 0]).split(":")[0].strip()
    df["File Name"] = file_name
    #select data rows
    data = df.iloc[6:, :]
    #rename two columns
    data = data.rename(columns={0: "No", 1: "Label"})
    #select amount rows
    amount = data.iloc[:, 2:]
    #select date row
    date = df.iloc[5, 2:]
    #put date as col header
    amount.columns = date
    #combine needed columns
    df2 = pd.concat([data[["No", "Label", "File Name"]], amount], axis=1)
    #melt rows into col
    df3 = df2.melt(id_vars=["No", "Label","File Name"],
                   var_name="Date",value_name="Amount")
    #drop missing data rows
    df3 = df3.dropna(subset=["Amount"])
    #Split date column
    df3 = df3.copy()
    df3[["Month","Year"]] = df3["Date"].astype(str).str.split(r"\s+" ,n=1, expand=True)
    #drop date col
    df3 = df3.drop(columns=["Date"])
    #change month type
    month_map = {"JAN": 1, "FEB": 2, "MAR": 3, "APR": 4,"MAY": 5, "JUN": 6, "JUL": 7, "AUG": 8,"SEP": 9, "OCT": 10, "NOV": 11, "DEC": 12}
    df3["Month"] = (df3["Month"].astype(str).str.upper().str.strip().map(month_map))
    #change data types
    df3["No"] = pd.to_numeric(df3["No"], errors="coerce").astype("Int64")
    df3["Label"] = df3["Label"].astype(str).str.strip()
    df3["File Name"] = df3["File Name"].astype(str).str.strip()
    df3["Amount"] = pd.to_numeric(df3["Amount"], errors="coerce").astype("float64")
    df3["Month"] = pd.to_numeric(df3["Month"], errors="coerce").astype("Int64")
    df3["Year"] = pd.to_numeric(df3["Year"], errors="coerce").astype("Int64")
    #arrange col names
    df3 = df3[["No", "Label", "Amount", "Month", "Year", "File Name"]]
    #add to all file list
    all_df.append(df3)
#concat all files    
final_df = pd.concat(all_df, ignore_index=True)

#extract excel file
final_df.to_excel("Final_Commercial_Bank_Consolidated.xlsx",index=False)

print("Completed")
