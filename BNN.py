from datetime import timedelta
import pandas as pd


import_file = input("File path\n").replace("& ","").replace("'","")
print("Enter Sheet Name")
sheet_name = input()

try:
    df = pd.read_excel(import_file, sheet_name=sheet_name, keep_default_na=False)
    df_dup = pd.read_excel(import_file, sheet_name=sheet_name, keep_default_na=False)

    # add premium
    if (sheet_name == "90 Day"):
        df["Premium"] = 20
    elif (sheet_name == "180 Day"):
        df["Premium"] = 40
    elif (sheet_name == "365 Day"):
        df["Premium"] = 0
        for index,row in df.iterrows():
            if row["ราคาสินค้า"] < 6000:
                df.at[index, "Premium"] = 50
            elif row["ราคาสินค้า"] > 6000 and row["ราคาสินค้า"] < 20000:
                df.at[index, "Premium"] = 90
            else:
                df.at[index, "Premium"] = 110
    else:
        df["Premium"] = 0
    
    # add column for sheet 365D
    if (sheet_name == "365 Day"):
        df["Com"] = df["Premium"] * 18/100 # add commission
        df["Pre-Com"] = (df["Premium"] - df["Com"]) # add Pre-Com
        df["Stamp"] = df["Pre-Com"] * 0.004
        df["Vat"] = (df["Pre-Com"] + df["Stamp"]) * 7/100
        df["Total"] = df["Pre-Com"] + df["Stamp"] + df["Vat"]
    else:
        df["Stamp"] = df["Premium"] * 0.004 # add stamp
        df["Vat"] = (df["Premium"] + df["Stamp"]) * 7/100 # add vat
        df["Total"] = df["Premium"] + df["Stamp"] + df["Vat"] # add total

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
    df.insert(17, "วันสิ้นสุด", end_day)

    # week
    print("จำนวนสัปดาห์")
    inp = input()

    # insert column policy
    df["AR_Policy"] = "" 
            
    print("Format Date (dd/mm/yyyy)")
    for x in range(int(inp)): # type: ignore
        print("วันเริ่มคุ้มครองสัปดาห์ที่", x + 1)
        week_start = input()
        date_start = pd.to_datetime(week_start, dayfirst=True, errors="coerce")
        print("วันสิ้นสุดกรมธรรม์สัปดาห์ที่", x + 1)
        week_end = input()
        date_end = pd.to_datetime(week_end, dayfirst=True, errors="coerce")
        print("กรมธรรม์สัปดาห์ที่", x + 1)
        week_policy = input()
    for index, row in df.iterrows():
        if row["วันลงทะเบียน"] >= date_start and row["วันลงทะเบียน"] <= date_end:
            df.at[index, "AR_Policy"] = week_policy

    # add sheet dup with duplicate value
    df1 = df_dup.sort_index(ascending=False)
    df_dup = df_dup.drop(df1[df1["IMEI สินค้า"].duplicated() == False].index, inplace=False)


    # export file output
    with pd.ExcelWriter("output.xlsx") as writer:
        df.to_excel(writer, sheet_name=sheet_name)
        df_dup.to_excel(writer, sheet_name="Dup")

    print("Success")

except Exception as e:
    print("Error! Please Check.", e)