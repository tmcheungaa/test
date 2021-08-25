from os import walk
import pandas as pd
import win32gui
from win32com.shell import shell, shellcon
import subprocess
import pyodbc
import re
# ========================================================================================================================================================================
def listToString(s):  
    
    # initialize an empty string 
    str1 = ""  
    
    # traverse in the string   
    for idx,ele in enumerate(s):
        if idx == len(s)-1:
            str1 += f'{ele}'
            
        else:
            str1 += f'{ele}, '
    
    # return string   
    return str1 

mydocs_pidl = shell.SHGetFolderLocation (0,shellcon.CSIDL_PERSONAL, 0, 0)
pidl, display_name, image_list = shell.SHBrowseForFolder (
  win32gui.GetDesktopWindow (),
  mydocs_pidl,
  "Select Tree Photo folder",
  shellcon.BIF_BROWSEINCLUDEFILES,
  None,
  None
)

if (pidl, display_name, image_list) == (None, None, None):
  print ("Nothing selected")
else:
  path = shell.SHGetPathFromIDList (pidl)
  print (path.decode("utf-8"))

# DATABASE CONNECTION
access_path = "C:/DB/ProcessDB.mdb"
con = pyodbc.connect("DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={};charset=65001;" \
                     .format(access_path))
# ========================================================================================================================================================================
mypath = path.decode("utf-8")
subprocess.call([r'C:/Users/david.cheung1/Documents/ITDC2/Raw Data/Delete thumbs.bat'])
dictionary ={}

# 遞迴列出所有檔案的絕對路徑
for _, dirs1, _ in walk(mypath):
    for i in dirs1:
        for _, dirs2, _ in walk(f'{mypath}\{i}'):
            for j in dirs2:
                for _, _, files in walk(f'{mypath}\{i}\{j}'):
                    folder = f'{mypath}\{i}\{j}'
                    slope_id = folder.split('\\', -1)[-2]
                    pos = slope_id.find('-') + 2
                    slope_id = f'{slope_id[:pos]}/{slope_id[pos:]}'
                    files.append(slope_id)
                    dictionary[j]=listToString(files)

new_df = pd.DataFrame.from_dict(dictionary, orient='index',columns=['PHOTO'])
new_df.index.name = 'TREE_ID'
new_df['slope_id'] = new_df['PHOTO'].str.split(', ',-1).str[-1]
new_df['PHOTO'] = new_df['PHOTO'].str.split(', ',-1).str[0:-1].apply(lambda x: ', '.join([str(i) for i in x]))
new_df['NUM_PHOTO'] = new_df['PHOTO'].str.split(',',-1).apply(lambda x: len(x))
#https://blog.csdn.net/qq_39885465/article/details/106650255

result = {}
for i in range(len(new_df)):
    str1 = new_df.index[i]
    str2 = new_df['PHOTO'].iloc[i]
    photo = str2.split(', ',-1)
    temp = []
    temp.append(not all([bool(re.search(str1, j)) for j in photo]))
    if any(temp):    
        result[str1] = temp


new_df.to_csv('C:/Users/david.cheung1/Documents\ITDC2\py_photo_count.csv')
# ========================================================================================================================================================================

strSQL = "SELECT * INTO PHOTO_COUNT FROM [text;HDR=Yes;FMT=Delimited(,);CharacterSet=65001;" + \
         "Database=C:/Users/david.cheung1/Documents/ITDC2].py_photo_count.csv;" 
         
cur = con.cursor()
cur.execute("DROP TABLE PHOTO_COUNT")
cur.execute(strSQL)
con.commit()
print(new_df.count(axis=0))
print("ALL PHOTOS ARE FOUND")
print("*"*99)
if len(result) > 0:
    print("Photo name hv some issues")
    print(result)
    print("*"*99)
else:
    print("Photo names are all correct")


        

