# VB-to-refresh-AP
To refresh the AP report


Sub RefreshAP_report()
Application.Wait (Now + #12:00:10 AM#)

Workbooks.Open Filename:= _
        "G:\Hotline\MIS\Data\Webi Data Dumps\Performance Report by Teams.xlsm"
        
ActiveWorkbook.RefreshAll

Activeworkbook.close True 
End Sub
