Dim xlBook ,objExcel
Set objExcel = CreateObject("Excel.Application")
Set xlBook = objExcel.Workbooks.Open("C:\Users\User\Desktop\AVMDART\Macro\AVM Dart_Automation_Final.xlsm", 0, True) 
objExcel.DisplayAlerts = False
objExcel.Visible = True
objExcel.Run "Main"

objExcel.Quit