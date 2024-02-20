import pandas as pd
from datetime import timedelta

o = input("File path\n")

try: 
    df = pd.read_excel (o, keep_default_na=False)
    df_dup = pd.read_excel (o, keep_default_na=False)

    # add premium
    df["Premium"] = 0

    for index,row in df.iterrows(): 
        if row["Code"] == "SUPER7CAREPLUS08":
            df.at[index,"Premium"] = 20
        elif row["Code"] == "SUPER7CAREPLUS15":
            df.at[index,"Premium"] = 30
        elif row["Code"] == "SUPER7CAREPLUS30":
            df.at[index,"Premium"] = 50
        elif row["Code"] == "SUPER7CAREPLUS50":
            df.at[index,"Premium"] = 80
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
    date = pd.to_datetime(df["วันลงทะเบียน"], format="%d/%m/%Y", errors="coerce")
    seven_year = timedelta(days=7 * 365)
    end_year = date + seven_year
    # insert column date
    df.insert(17, "วันสิ้นสุด", end_year)

    # add sheet dup with duplicate value
    df1 = df_dup.sort_index(ascending=False)
    df_dup = df_dup.drop(df1[df1["IMEI สินค้า"].duplicated() == False].index, inplace=False)

    # export
    with pd.ExcelWriter('output.xlsx') as writer:  # doctest: +SKIP
        df.to_excel(writer, sheet_name='7Y')
        df_dup.to_excel(writer, sheet_name='Dup')

    print("Success")

except:
    print("Some Mistake")