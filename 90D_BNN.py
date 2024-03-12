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

    # ask date of policy week 1
    print("วันเริ่มต้นของกรมธรรม์ สัปดาห์ที่ 1 คือ")
    value_start_week_one = input()
    start_week_one = pd.to_datetime(value_start_week_one, dayfirst=True, errors="coerce")
    print("วันสิ้นสุดของกรมธรรม์ สัปดาห์ที่ 1 คือ")
    value_end_week_one = input()
    end_week_one = pd.to_datetime(value_end_week_one, dayfirst=True, errors="coerce")

    # policy number
    print("กรมธรรม์ของ สัปดาห์ที่ 1 ของเดือน คือ")
    policy_week_one = input()

    # ask date of policy week 2
    print("วันเริ่มต้นของกรมธรรม์ สัปดาห์ที่ 2 คือ")
    value_start_week_two = input()
    start_week_two = pd.to_datetime(value_start_week_two, dayfirst=True, errors="coerce")
    print("วันสิ้นสุดของกรมธรรม์ สัปดาห์ที่ 2 คือ")
    value_end_week_two = input()
    end_week_two = pd.to_datetime(value_end_week_two, dayfirst=True, errors="coerce")
    
    # policy number
    print("กรมธรรม์ของ สัปดาห์ที่ 2 ของเดือน คือ")
    policy_week_two = input()

    # ask date of policy week 3
    print("วันเริ่มต้นของกรมธรรม์ สัปดาห์ที่ 3 คือ")
    value_start_week_three = input()
    start_week_three = pd.to_datetime(value_start_week_three, dayfirst=True, errors="coerce")
    print("วันสิ้นสุดของกรมธรรม์ สัปดาห์ที่ 3 คือ")
    value_end_week_three = input()
    end_week_three = pd.to_datetime(value_end_week_three, dayfirst=True, errors="coerce")

    # policy number
    print("กรมธรรม์ของ สัปดาห์ที่ 3 ของเดือน คือ")
    policy_week_three = input()

    # ask date of policy week 4
    print("วันเริ่มต้นของกรมธรรม์ สัปดาห์ที่ 4 คือ")
    value_start_week_four = input()
    start_week_four = pd.to_datetime(value_start_week_four, dayfirst=True, errors="coerce")
    print("วันสิ้นสุดของกรมธรรม์ สัปดาห์ที่ 4 คือ")
    value_end_week_four = input()
    end_week_four = pd.to_datetime(value_end_week_four, dayfirst=True, errors="coerce")

    # policy number
    print("กรมธรรม์ของ สัปดาห์ที่ 4 ของเดือน คือ")
    policy_week_four = input()

    # policy number week 5
    print("กรมธรรม์ของ สัปดาห์ที่ 5 ของเดือน คือ")
    policy_week_five = input()

    
    # insert column policy
    df["AR_Policy"] = "" 

    for index, row in df.iterrows():
        if row["วันลงทะเบียน"] >= start_week_one and row["วันลงทะเบียน"] <= end_week_one:
            df.at[index, "AR_Policy"] = policy_week_one
        elif row["วันลงทะเบียน"] >= start_week_two and row["วันลงทะเบียน"] <= end_week_two:
            df.at[index, "AR_Policy"] = policy_week_two
        elif row["วันลงทะเบียน"] >= start_week_three and row["วันลงทะเบียน"] <= policy_week_three:
            df.at[index, "AR_Policy"] = policy_week_three
        elif row["วันลงทะเบียน"] >= start_week_four and row["วันลงทะเบียน"] <= policy_week_four:
            df.at[index, "AR_Policy"] = policy_week_four
        else:
            df.at[index, "AR_Policy"] = policy_week_five


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