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
    
    url_mainfile = cnv_urlgooglesheet_to_id(url_mainfile)
    
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

def listfile_from_googledrive(drive, folder_id):
    file_list = drive.ListFile({'q': "'" + folder_id + "' in parents and trashed=false"}).GetList()
    return file_list



def fnc_loopfrom_googledrive(file_list, val_sheetname, val_column):
    
    import pandas as pd
    
    drive = connect_google_drive_api()
    import gspread
    gc = gspread.oauth() ##appdata roamin gspreadsheet
    
    #Create new folder
    import os
    if(os.path.exists('file attach') == False):
        os.mkdir('file attach')
    
    
    
    i = 1
    for file in file_list:

        extenson = file['title'].split('.')[-1].lower()

        if extenson == 'xlsx' or extenson == 'xlsx' : #กรณีเจอไฟล์ xlsx จะถูก download ข้อมูลลงมาเรียกใช้งาน
            filename = file['title']
#             print(filename)
            
            downloaded1 = drive.CreateFile({'id':  file['id']})
            downloaded1.GetContentFile(r'file attach/' + filename) 
            df_acc = pd.read_excel(
                                r'file attach/' + filename,
                                sheet_name=val_sheetname, 
                                skiprows =3,  
                                usecols = val_column,   
                                index_col = None).fillna('')
                                #nrows= 3

        #add file name when append data in sheet
        title = file['title']
        df_acc.insert(0,'name_of_column',title)
#         df_acc['GL Account No']=df_acc['GL Account No'].astype(str).str.replace('\.0',"")


        if(i == 1 ):
            df1 = df_acc.copy()
            del df_acc
        else:
            df1 = df1.append(df_acc)

        i += 1
        
    return df1



## call ตัวแปรจาก google sheet ( เรียกแบบ range) แล้วมาทำเป็น dicitonary
def feature_call_variable_fromgooglesheet(gc, sheet_lookup, sheet_name, use_col):
    sheet_lookup = sheet_lookup.replace("https://docs.google.com/spreadsheets/d/", "").split("/")[0]
    sh_link = gc.open_by_key(sheet_lookup)
    s_range = sh_link.worksheet(sheet_name)

    #https://stackoverflow.com/questions/60127455/python-gspread-range-in-format-a2a-a1-notation
    a = list(s_range.batch_get( (use_col,) )[0])
    
    #แปลงเป็น dictionary
    new_list = {}
    for k, v in a:
        new_list.setdefault(k, []).append(v)
    return new_list

