Set objExcel = CreateObject("Excel.Application")
objExcel.Application.Run "'D:\Testpress\Macro\combine ssc.xlsm'!Sheet1.CombineWSs"
objExcel.DisplayAlerts = False
objExcel.Application.Quit
Set objExcel = Nothing
