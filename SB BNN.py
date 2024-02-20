import pandas as pd
import os
import sys
from datetime import timedelta

# input without specific and space
o = input("File path\n").replace("& ","").replace("'","")

try: 
    df = pd.read_excel (o, keep_default_na=False)
    df_dup = pd.read_excel (o, keep_default_na=False)

    # add premium
    df["Premium"] = 0

    for index,row in df.iterrows(): 
        if row["Code"] == 70000090:
            df.at[index,"Premium"] = 20
        elif row["Code"] == 70000180:
            df.at[index,"Premium"] = 40
        elif row["Code"] == 70000365:
            if row["ราคาสินค้า"] < 6000:
                df.at[index,"Premium"] = 50
            elif row["ราคาสินค้า"] > 6000 and row["ราคาสินค้า"] < 20000:
                df.at[index,"Premium"] = 90
            elif row["ราคาสินค้า"] > 20000:
                df.at[index, "Premium"] = 110
        else:
            df.at[index,"Premium"] = 0

    # add stamp
    df["Stamp"] = round(df["Premium"] * 0.004,2)
    # add vat
    df["Vat"] = round((df["Premium"] + df["Stamp"]) * 7/100,2)
    # add total
    df["Total"] = round(df["Premium"] + df["Stamp"] + df["Vat"],2)
    # delte duplicate
    df = df.drop_duplicates(subset=['IMEI สินค้า'], keep='last')
    # convert date
    three_month = timedelta(days=90)
    six_month = timedelta(days=180)
    one_year = timedelta(days=1* 365)

    # insert column date
    df.insert(17, "วันสิ้นสุด", df["วันลงทะเบียน"].apply(pd.to_datetime, format="%d/%m/%Y", errors="coerce"))


    df.loc[df["Code"] == 70000090,"วันสิ้นสุด"] += three_month
    df.loc[df["Code"] == 70000180,"วันสิ้นสุด"] += six_month
    df.loc[df["Code"] == 70000365,"วันสิ้นสุด"] += one_year

    # add sheet dup with duplicate value
    df1 = df_dup.sort_index(ascending=False)
    df_dup = df_dup.drop(df1[df1["IMEI สินค้า"].duplicated() == False].index, inplace=False)

    # export
    with pd.ExcelWriter('output.xlsx') as writer:  # doctest: +SKIP
        df.to_excel(writer, sheet_name='SB_BNN')
        df_dup.to_excel(writer, sheet_name='Dup')

    print("Success")

except Exception as e:
    print("Some Mistake, Check your file.", e)