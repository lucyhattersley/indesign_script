import win32com.client
import os
app = win32com.client.Dispatch('InDesign.Application')

myFile =r'C:\Users\lucyh\Scripting_InDesign\032-039_MagPi#91_FEATURE_MonthOfMaking_NK.indd'
directory = os.path.dirname(myFile)



myDocument = app.Open(myFile)
