''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Name: ZabbixMailFlowCheck
' Version: v1.0
' Created By: David W. Gilmore
' Modified: December 9, 2014
'
' Summary:
' This script runs the imap_checker utility, gets the date of the last message, and returns it to the Zabbix server. Note that
' Epoch time is used because this is how Zabbix stores date/time values
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim servername, hostname, zbxSender, imap_checker, objWshScriptExec, objStdOut, epoch, standardTimeOffset, DSTOffset, arrDST(1)
Dim msgDate, justTheDate

serverName = "1.2.3.4"	'The IP Address of the Zabbix server
hostName = "myEMailHost" 'The value of the Host Name field for the given item
zbxSender = "C:\bin\zabbix\bin\win32\zabbix_sender.exe" 'Path to the local zabbix_sender application
imap_checker = "c:\bin\imap_checker\imap_checker.exe" 'Path to the IMAP Checker utility

'For a list of epoch offset values for various timezones see http://www.epochconverter.com/epoch/timezones.php
standardTimeOffset = 28800 'Epoch offset in seconds (current value is for PST)
DSTOffset = 25200 'Epoch offset in seconds during DST (current value is for PST)
Dim objShell : Set objShell = WScript.CreateObject("WScript.Shell")

'Run the IMAP checker utility and return the date/time it reports
Set objWshScriptExec = objShell.Exec(imap_checker)
Set objStdOut = objWshScriptExec.StdOut

'Convert the date/time to epoch 
msgDate = CDate(objStdOut.ReadLine)
epoch = date2epoch(msgDate)
justTheDate = FormatDateTime(msgDate, 2)

'Determine if we are in DST and apply the desited offset
If IsDST(justTheDate,arrDST) = 1 Then
    epoch = epoch + DSTOffset
Else
    epoch = epoch + standardTimeOffset
    
End If

'Update the item on the zabbix server
objShell.Exec zbxSender & " -z " & serverName & " -s " & hostName & " -k mailflow_check -o " & epoch

'This Function converts a human readable date/time stamp to Unix Epoch time (Zabbix item requires this)
function date2epoch(myDate)
date2epoch = DateDiff("s", "01/01/1970 00:00:00", myDate)
end function

'This function determines if a given date falls into Daylight Savings Time
Function isDST(argDate, argReturn)
 Dim StartDate, EndDate
 
 If (Not IsDate(argDate)) Then
  argReturn(0) = -1
  argReturn(1) = -1
  isDST = -1
  Exit Function
 End If
 
 ' DST start date...
 StartDate = DateSerial(Year(argDate), 3, 1)
 Do While (WeekDay(StartDate) <> vbSunday)
  StartDate = StartDate + 1
 Loop
 StartDate = StartDate + 7
 
 ' DST end date...
 EndDate = DateSerial(Year(argDate), 11, 1)
 Do While (WeekDay(EndDate) <> vbSunday)
  EndDate = EndDate + 1
 Loop

  ' Finish up...
 isDST = 0
 If ((argDate >= StartDate) And (argDate < EndDate)) Then
  argReturn(0) = StartDate
  argReturn(1) = EndDate
  isDST = 1
 End If
End Function