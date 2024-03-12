from datetime import timedelta
import pandas as pd


import_file = input("File path\n").replace("& ","").replace("'","")

try:
    df = pd.read_excel(import_file, sheet_name="90 Day", keep_default_na=False)
    df_dup = pd.read_excel(import_file, sheet_name="90 Day", keep_default_na=False)

    # add premium
    df["Premium"] = 20

    # add stamp
    df["Stamp"] = round(df["Premium"] * 0.004,2)
    # add vat
    df["Vat"] = round((df["Premium"] + df["Stamp"]) * 7/100,2)
    # add total
    df["Total"] = round(df["Premium"] + df["Stamp"] + df["Vat"],2)
    # delete duplicate
    df = df.drop_duplicates(subset=["IMEI สินค้า"], keep='last')
    # convert str to date
    str_to_date = pd.to_datetime(df["วันลงทะเบียน"],dayfirst=True)
    df["วันลงทะเบียน"] = str_to_date
    # calculate date
    date = pd.to_datetime(str_to_date, format="%d%m%Y", errors="coerce")
    ninety_day = timedelta(days=90)
    end_day = date + ninety_day

    # insert colum date
    df.insert(13, "วันสิ้นสุด", end_day)

    # add sheet dup with duplicate value
    df1 = df_dup.sort_index(ascending=False)
    df_dup = df_dup.drop(df1[df1["IMEI สินค้า"].duplicated() == False].index, inplace=False)


    # export file output
    with pd.ExcelWriter("output.xlsx") as writer:
        df.to_excel(writer, sheet_name="90D")
        df_dup.to_excel(writer, sheet_name="Dup")

    print("Success")

except Exception as e:
    print("Error! Please Check.", e)