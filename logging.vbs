' Syntax:
'  CSCRIPT datetime.vbs
 
'Returns: Year, Month, Day, Hour, Minute, Seconds, Offset from GMT, Daylight Savings=True/False

strComputer = "."

' Date and time

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")

For Each objItem in colItems
    dtmLocalTime = objItem.LocalDateTime
    dtmHour = Mid(dtmLocalTime, 9, 2)
    dtmMinutes = Mid(dtmLocalTime, 11, 2)
Next

' Daylight savings

Set Win32Computer = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem") 

For Each objItem In Win32Computer
   oGMT = (objItem.CurrentTimeZone / 60) 
   DaySave = objItem.DaylightInEffect 
Next
mode = "AM"
status = "Check In"
if dtmHour > 12 Then
   mode = "PM"
   status = "Check Out"
   dtmHour = dtmHour - 12
end if
quote = status & " " & Right("0" & Cstr(dtmHour),2) & ":" & Right("0" & Cstr(dtmMinutes),2) & " " & mode


Set WshShell = CreateObject("WScript.Shell")
Set oExec = WshShell.Exec("clip")

Set oIn = oExec.stdIn

oIn.WriteLine quote
oIn.Close