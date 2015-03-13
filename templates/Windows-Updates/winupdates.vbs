'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Winupdates.vbs
'Version 1.1
'Written By: Dave Gilmore
'Created On: 02/27/2015
'
' 
'Purpose: Given a class of update (high or important) returns a count of the updates waiting
'			to be installed. Will also return reboot status
'
' CHANGELOG
' v1.1
' - Cleaned up the code a bit
' - hostName is now dynamically generated and does not have to be hard coded into the script

Function ReturnUpdateCount(updateType)
	Set updateSession = CreateObject("Microsoft.Update.Session") 
	Set updateSearcher = updateSession.CreateupdateSearcher()
	
	Select Case updateType
	Case "high"
		Set searchResult = updateSearcher.Search("IsAssigned=1 and IsHidden=0 and IsInstalled=0 and Type='Software'") 
	Case "optional"
		Set searchResult = updateSearcher.Search("IsAssigned=0 and IsInstalled=0 and Type='Software'")
	End Select
	ReturnUpdateCount = searchResult.Updates.Count	

End Function

Function RebootCheck
	Set objSysInfo = CreateObject("Microsoft.Update.SystemInfo")
	RebootCheck = objSysinfo.RebootRequired
End Function

Set wshNetwork = WScript.CreateObject( "WScript.Network" )
hostName = wshNetwork.ComputerName
serverName = "1.2.3.4"
zbxSender = "C:\bin\zabbix\bin\win32\zabbix_sender.exe"
Set WSHShell = CreateObject("WScript.Shell")

updatesHigh = ReturnUpdateCount("high")
updatesOptional = ReturnUpdateCount("optional")
strRebootRequired = RebootCheck

WSHShell.Exec zbxSender & " -z " & serverName & " -s " & hostName & " -k win_updates[high] -o " & updatesHigh

WSHShell.Exec zbxSender & " -z " & serverName & " -s " & hostName & " -k win_updates[optional] -o " & updatesOptional

WSHShell.Exec zbxSender & " -z " & serverName & " -s " & hostName & " -k win_updates[reboot] -o " & strRebootRequired

WScript.Quit 0
