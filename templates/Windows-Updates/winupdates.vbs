'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Name: winupdates.vbs 
' Created By: David W. Gilmore
' Version: v1.0
' Last Updated: 08/12/2013
'
' Summary: This script checks to see if any Windows Updates are waiting to be installed
'			and feeds the value to the zabbix_sender process. It will also report if the 
'			machine needs to be rebooted to complete the install.

serverName = "xxx.xxx.xxx.xxx" 'IP of your Zabbix server
hostName = "myWindowsServer" 'Hostname of Windows machine as it appears in Zabbix
zbxSender = "C:\bin\zabbix\bin\win32\zabbix_sender.exe" 'Path to zabbix_sender process

updatesHigh = 0
updatesOptional = 0

Set objSearcher = CreateObject("Microsoft.Update.Searcher")
Set objSysInfo = CreateObject("Microsoft.Update.SystemInfo")
Set objResults = objSearcher.Search("IsInstalled=0")
Set colUpdates = objResults.Updates
Set WSHShell = CreateObject("WScript.Shell")

For i = 0 to colUpdates.Count - 1

	If (colUpdates.Item(i).IsInstalled = False AND colUpdates.Item(i).AutoSelectOnWebSites = False) Then
		updatesOptional = updatesOptional + 1
	ElseIf (colUpdates.Item(i).IsInstalled = False AND colUpdates.Item(i).AutoSelectOnWebSites = True) Then
		updatesHigh = updatesHigh + 1
	End IF
    
Next

updatesTotal = (updatesHigh + updatesOptional)

WSHShell.Exec zbxSender & " -z " & serverName & " -s " & hostName & " -k win_updates[total] -o " & updatesTotal

WSHShell.Exec zbxSender & " -z " & serverName & " -s " & hostName & " -k win_updates[high] -o " & updatesHigh

WSHShell.Exec zbxSender & " -z " & serverName & " -s " & hostName & " -k win_updates[optional] -o " & updatesOptional

WSHShell.Exec zbxSender & " -z " & serverName & " -s " & hostName & " -k win_updates[reboot] -o " & objSysInfo.RebootRequired


WScript.Quit 0