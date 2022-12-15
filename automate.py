##11111
from gspread_dataframe import set_with_dataframe
from gspread_dataframe import get_as_dataframe


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


#แปลง link url google sheet ตัดให้เหลือ id อย่างเดียว
def cnv_urlgoogledrive_to_id(url1):
    url = url1
    url = url.replace("https://drive.google.com/drive/folders/", "").split("/")[0]
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
#                                 skiprows =3,  
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


def append_tosheet_set_with_dataframe(df, sheet_lookup, sheet_name, my_range):
    
    import gspread
    gc = gspread.oauth() ##appdata roamin gspreadsheet
    
    sheet_lookup = sheet_lookup.replace("https://docs.google.com/spreadsheets/d/", "").split("/")[0]
    sh_link = gc.open_by_key(sheet_lookup)

    s_range = sh_link.worksheet(sheet_name)
    
    clear_value_sheet(sh_link, sheet_name , my_range)
    
    sheet_destination_sheet = sh_link.worksheet(sheet_name)
    
    set_with_dataframe(sheet_destination_sheet, 
                       df, 
                       row=1, 
                       include_column_header=True) 




def clear_value_sheet(sh_link, sheetname,my_range):
    sh_bot = sh_link.worksheet(sheetname)
    
    find_lastrow = next_available_row(sh_bot)
    selectrange = str(my_range) + str(find_lastrow)
    sh_link.values_clear("'{}'!{}".format(sheetname, selectrange ))
    
    
def next_available_row(sh_bot):
#     str_list = list(filter(None, worksheet.col_values(1)))
#     return str(len(str_list)+1)
    return len(sh_bot.get_all_values()) + 1




def feature_duplicatefile_and_changeowner(url, 
                                          val_destination_folder_backup,
                                          rename_file,
                                          val_delete_sheetlist):
                                          
                                          
    
    import pandas as pd
    
    drive = connect_google_drive_api()
    import gspread
    gc = gspread.oauth() ##appdata roamin gspreadsheet
    
    file_main = feature_duplicate_export_file(
                              drive,
                              url,
                              val_destination_folder_backup,
                              rename_file
                             )
    sh_link_export = gc.open_by_key(file_main['id'])
    delete_sheet_not_use(sh_link_export, val_delete_sheetlist)
    
    start_run, time_string, yyyymmdd, dt_string, now1, today, time = my_date()
    ##--------- Logfile
    list_add = [[
        yyyymmdd  ,
        start_run,
        time_string, #now.strftime("%H:%M:%S"),
#         str(list2D_month)
    ]]

    sh_link_template = gc.open_by_key(url)
    
    
    df_log_run = pd.DataFrame(list_add, columns =['Date','Start run time','End run time'])
#     append_df_to_sheet(sh_link_template, 'Log Run', df_log_run)

    return file_main, df_log_run
    
    

    
def feature_duplicate_export_file(drive, sheet_id, specific_folderid, rename_file):
    print("sheet_id", sheet_id)
    print("specific_folderid", specific_folderid)
    print("rename_file", rename_file)
    file1 = drive.CreateFile({'id': sheet_id})
    file1.Upload()                 # Upload new title.
    file_main = drive.auth.service.files().copy(fileId=sheet_id, body={"parents": [{"id": specific_folderid}], 'title': rename_file}).execute()
    return file_main

                     
def delete_sheet_not_use(sh_link, list_del_sheet):
    for list_sheet in list_del_sheet:
        obj_sheet_del = sh_link.worksheet(list_sheet)
        sh_link.del_worksheet(obj_sheet_del)
        
def append_df_to_sheet(sh_link, sheetname, df):
    sh_bot = sh_link.worksheet(sheetname)
    data_list = df.values.tolist()
    sh_bot.append_rows(data_list)    

    


def feature_changefile_owner(cell_changeowner_ID, file_main):
    from apiclient import discovery
    from pydrive2.auth import GoogleAuth
    from pydrive2.drive import GoogleDrive
    gauth = GoogleAuth()
    gauth.LocalWebserverAuth() # client_secrets.json need to be in the same directory as the script    

        
    change_mail_owner = cell_changeowner_ID
    

    drive_service = discovery.build('drive', 'v3', credentials=gauth.credentials) 
    permission = {
                "emailAddress": change_mail_owner,
                "role": 'owner',
                "type": 'user',
    }
    drive_service.permissions().create(fileId=file_main['id'], body=permission, transferOwnership=True).execute()
    print('complete')
    
    


def feature_sendmail_pattern1_compare(
    file_main, 
    rename_file, 
    val_array, 
    subject1, 
    subject2
):
    start_run, time_string, yyyymmdd, dt_string, now1, today, time = my_date()

    
    if(cnvlist2str(val_array['ต้องการเปิดใช้ Function ส่งเมลหรือไม่']) == "Y"): 
        
        
        email_subject = '[Automate' + '] ' + cnvlist2str(val_array[subject1]) + " " + cnvlist2str(val_array[subject2]) + ' .'

        mail_content = '''
        <u><b>Result</b></u><br>
        %s
        <br>
        <br>
        <u><b>Log</b></u><br>
        1. เริ่มรัน ณ เวลา %s สิ้นสุด ณ %s <br>
        2. Folder Backup Google Sheets %s <br>
        3. กรณีต้องการ Setting เงื่อนไขเพิ่มเติมด้วยตนเอง %s <br>
        <br>
        ''' % (    
                   '<a href="'+ file_main['alternateLink']+'">' + rename_file + '</a>',
                   start_run, 
                   datetime.now().strftime("%H:%M:%S"),
                   cnvlist2str(val_array['val_folderbackup - link แนบใน email ข้อ 2']),
                   cnvlist2str(val_array['val_templatesetting - link แนบใน email ข้อ 3']),


                   )

    #------------------------
    feature_sendmail(cnvlist2str(val_array['Sender address']), 
                     cnvlist2str(val_array['Sender pass']),
                     cnvlist2str(val_array['To (ผู้รับเมล : receiver_address)']),
                     email_subject,
                     mail_content
                    )
   



def feature_sendmail(sender_address, 
                 sender_pass,
                 receiver_address,
                 email_subject,
                 mail_content
                ):
    #python send gmail with smtp geekshare
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText

    #Setup the MIME
    message = MIMEMultipart()
    message['From'] = sender_address

    receiver_address = receiver_address.split(",")
    receiver_address = [i.lstrip() for i in receiver_address]
    message['To'] = ", ".join(receiver_address)

    message['Subject'] = email_subject   #The subject line
    #The body and the attachments for the mail
    message.attach(MIMEText(mail_content, 'html')) #ตัวเลือกว่าจะแสดงแบบ text, html
    #Create SMTP session for sending the mail
    session = smtplib.SMTP('smtp.gmail.com', 587) #use gmail with port
    session.starttls() #enable security
    session.login(sender_address, sender_pass) #login with mail_id and password
    text = message.as_string()
    session.sendmail(sender_address, receiver_address, text)
    session.quit()
