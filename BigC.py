import pandas as pd
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
            df.at[index, "Premium"] = 80
        else:
            df.at[index,"Premium"] = 0

    # add stamp
    df["Stamp"] = df["Premium"] * 0.004
    # add vat
    df["Vat"] = (df["Premium"] + df["Stamp"]) * 7/100
    # add total
    df["Total"] = df["Premium"] + df["Stamp"] + df["Vat"]
    # delte duplicate
    df = df.drop_duplicates(subset=['IMEI สินค้า'], keep='last')
    # convert date
    three_month = timedelta(days=90)
    six_month = timedelta(days=180)
    one_year = timedelta(days=1* 365)

    # insert column date
    df.insert(11, "วันสิ้นสุด", df["วันลงทะเบียน"].apply(pd.to_datetime, format="%d/%m/%Y", errors="coerce"))


    df.loc[df["Code"] == 70000090,"วันสิ้นสุด"] += three_month
    df.loc[df["Code"] == 70000180,"วันสิ้นสุด"] += six_month
    df.loc[df["Code"] == 70000365,"วันสิ้นสุด"] += one_year

    # add sheet dup with duplicate value
    df1 = df_dup.sort_index(ascending=False)
    df_dup = df_dup.drop(df1[df1["IMEI สินค้า"].duplicated() == False].index, inplace=False)

    # export
    with pd.ExcelWriter('output.xlsx') as writer:  # doctest: +SKIP
        df.to_excel(writer, sheet_name="All")
        df_dup.to_excel(writer, sheet_name="Dup")
        if (df["Code"].any() == 70000090):
            df.to_excel(writer, sheet_name="90")
        elif (df["Code"].any() == 70000180):
            df.to_excel(writer, sheet_name="180")
        else:
            df.to_excel(writer, sheet_name="365")

    print("Success")

except Exception as e:
    print("Some Mistake, Check your file.", e)