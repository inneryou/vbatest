Dim obj
Set obj = WScript.CreateObject("Excel.Application")
Set path = CreateObject("Scripting.FileSystemObject").GetFolder(".")
obj.Visible = true
CreateObject("WScript.Shell").AppActivate obj.Caption
obj.Workbooks.Open "C:\Users\<USER NAME>\Desktop\vbatest.xlsm"
obj.Application.Run "Main"
