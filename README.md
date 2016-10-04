# VB-to-refresh-AP
To refresh the AP report


Sub RefreshAP_report()
Application.Wait (Now + #12:00:10 AM#)

Workbooks.Open Filename:= _
        "G:\Hotline\MIS\Data\Webi Data Dumps\Performance Report by Teams.xlsm"
        
Workbooks(“Performance Report by Teams.xlsm”).RefreshAll

Activeworkbook.close True 
End Sub
