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
