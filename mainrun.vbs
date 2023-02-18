Dim obj
Set obj = WScript.CreateObject("Excel.Application")
Set path = CreateObject("Scripting.FileSystemObject").GetFolder(".")
obj.Visible = true
CreateObject("WScript.Shell").AppActivate obj.Caption
obj.Workbooks.Open "C:\Users\inner.WINDOWS10\Desktop\vbatest.xlsm"
obj.Application.Run "Main"
