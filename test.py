"""
This code is meant to convert any .xls ledger into a .xlsx file within minutes

(Probably a one time use code tbh)
"""

import win32com.client
import os
import re

walk_path = ''
sports_path = ''
non_sports_path = ''
off_campus_path = ''

#create file pattern to use when finding excel sheets (anything with 9 straight digits followed by a "-")
file_pattern = r'\d{9} - '

omitted_clubs = []

#create a loop to search through all excel sheets in the folder
for folders,sub_folders,files in os.walk(walk_path):

    #searches through all files in directory
    for f in files:

        #conditional to make sure that its an excel file and that it fits the desired file pattern
        if (
            (f.endswith(".xls"))
            and (folders == sports_path or folders == non_sports_path or folders == off_campus_path)
            and re.search(file_pattern,f)
            and f not in omitted_clubs
            ):

                excel = win32com.client.Dispatch('Excel.Application')
                excel.Visible = False

                file = os.path.basename(walk_path + f)
                wb = excel.Workbooks.Open(walk_path + f)
                wb.ActiveSheet.SaveAs(walk_path + f + "x",51)
                wb.Close(True)
