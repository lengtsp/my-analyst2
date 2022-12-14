##11111

def cnvlist2str(val):
    listToStr = ''.join(map(str, val))
    return listToStr

def connect_google_drive_api():
    from pydrive2.auth import GoogleAuth
    from pydrive2.drive import GoogleDrive
    
    gauth = GoogleAuth()
    gauth.LocalWebserverAuth() # client_secrets.json need to be in the same directory as the script    
    drive = GoogleDrive(gauth)
    
    return drive

def covert_sheet_to_df(gc, url_mainfile, sheetname):
    import pandas as pd
    
    sh_link = gc.open_by_key(url_mainfile)
    s_range = sh_link.worksheet(sheetname).get_all_values()
    df = pd.DataFrame(s_range)
    df = df.rename(columns=df.iloc[0]).drop(df.index[0])
    return df




def my_date():
    # https://www.programiz.com/python-programming/datetime/current-datetime
    from datetime import date
    from datetime import datetime

    today = date.today()

    # dd/mm/YY
    d1 = today.strftime("%d/%m/%Y")
    # print("d1 =", d1)

    # dd
    dd = today.strftime("%d")
    # print("dd =", dd)

    # dd/mm/YY
    mm = today.strftime("%m")
    # print("mm =", mm)

    # dd/mm/YY
    Y = today.strftime("%Y")
    # print("Y =", Y)

    # dd/mm/YY
    time = today.strftime("%H:%M:%S")
    # print("d1 =", d1)

    # print(Y[0:2])
    today = d1
    start_run = time

    # datetime object containing current date and time
    now1 = datetime.now()

    # dd/mm/YY H:M:S
    dt_string = now1.strftime("%d/%m/%Y %H:%M:%S")

    time_string = now1.strftime("%H:%M:%S")
    yyyymmdd = now1.strftime("%Y-%m-%d") + " " + time_string
    start_run = time_string
    
    print("start_run", start_run)
    print("time_string", time_string)
    print("yyyymmdd", yyyymmdd)
    print("dt_string", dt_string)
    print("now1", now1)
    print("today", today)
    print("time", time)
    return start_run, time_string, yyyymmdd, dt_string, now1, today, time

#แปลง link url google sheet ตัดให้เหลือ id อย่างเดียว
def cnv_urlgooglesheet_to_id(url1):
    url = url1
    url = url.replace("https://docs.google.com/spreadsheets/d/", "").split("/")[0]
    return url
