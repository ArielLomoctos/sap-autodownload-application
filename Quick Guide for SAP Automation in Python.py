
import sys
import win32com.client
import os as opsys

currentUsername = opsys.getlogin()
currentTable = "T001W"

# Create instance
SapGuiAuto = win32com.client.GetObject("SAPGUI")
Application = SapGuiAuto.GetScriptingEngine
Connection = Application.Children(0)
Session = Connection.Children(0)

# Transaction Code
Session.findById("wnd[0]/tbar[0]/okcd").text = "/n/ds1/yse16n"
Session.findById("wnd[0]").sendVKey(0)

# Transaction Table
Session.findById("wnd[0]/usr/ctxtP_TABLE").text = "T001W"
Session.findById("wnd[0]/tbar[1]/btn[8]").press()
Session.findById("wnd[1]/tbar[0]/btn[25]").press()

# Provide Plant No.
Session.findById("wnd[0]/usr/tbl/DS1/SAPLUTL_SE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,1]").text = "R123"
Session.findById("wnd[0]/tbar[1]/btn[8]").press()

# Save Download to Folder
Session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
Session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem("&XXL")
Session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\\Users\\" + currentUsername + "\\Desktop\\"
Session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = currentTable + ".xlsx"
Session.findById("wnd[1]/tbar[0]/btn[0]").press()

# Clear instance
Session = None
Connection = None
Application = None
SapGuiAuto = None
