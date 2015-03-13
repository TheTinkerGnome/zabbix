'==============================================================
'
'  Script:	check_dhcp_scope.vbs
'
'  Autore: 	Marco Gottardello - marco@gottardello.net
'
'  Data:	26/01/2010
'
'  Descrizione:	Check the health status of scopes on the local DHCP server
'
'  History:     v1.0 : first version
'
'==============================================================


Set objFSO = CreateObject("Scripting.FileSystemObject")
Set tfolder = objfso.GetSpecialFolder(2)

filetoparse= tfolder&"\EXPORT_DHCP_SCOPE.TXT"

bdebug=FALSE

strout=""

Critical_Limit=0
Warning_Limit=4

function parse(strtoparse)
	parse=mid(strtoparse,instr(strtoparse,"=")+2,len(strtoparse)-instr(strtoparse,"=")-2)
end function


'=========main=============

strthecommand="cmd.exe /c netsh dhcp server show all > "&filetoparse&" 2>&1"
CreateObject("WScript.Shell").Run (strthecommand),0,true
WScript.Sleep(5000)  


	ScopeCount=0
	ScopeOk=0
	ScopeDisabled=0
	ScopeWarning=0
	ScopeCritical=0

	
	Const ForReading = 1
	Const ForWriting = 2
	Const ForAppending = 8


	'apro i file sorgente
	Set objFileSorg = objFSO.OpenTextFile(filetoparse, ForReading)

	oldstr2=""
	oldstr3=""	
	oldstr=""	

	Do While objFileSorg.AtEndOfStream <> True  
		strSorg = objFileSorg.ReadLine
		if instr(oldstr2,"Subnet") then
			scopecount=scopecount+1

			SubnetIP=parse(oldstr2)
			FreeIP=parse(strsorg)
			UsedIP=parse(oldstr)
			
			if bdebug then wscript.echo SubnetIP&" "&UsedIP&" "&FreeIP

			if UsedIP=0 then 
				if bdebug then wscript.echo "-"&SubnetIP&" is disabled"
				ScopeDisabled=ScopeDisabled+1
			elseif int(FreeIP)<=Critical_Limit then 
				if bdebug then wscript.echo "-"&SubnetIP&" is Critical ("&FreeIP&" free)"
				StrOut=Strout&SubnetIP&" is Critical ("&FreeIP&" free). "
				ScopeCritical=ScopeCritical+1
			elseif int(FreeIP)<=Warning_Limit then 
				if bdebug then wscript.echo "-"&SubnetIP&" is Warning ("&FreeIP&" free)"
				StrOut=Strout&SubnetIP&" is Warning ("&FreeIP&" free). "
				ScopeWarning=ScopeWarning+1
			end if

		end if
		
		oldstr3=oldstr2
		oldstr2=oldstr
		oldstr=strsorg
	loop
	objFileSorg.Close

	if not strout="" then wscript.echo strout
	
	if ScopeCritical then wscript.quit(2)
	if ScopeWarning then wscript.quit(1)
	wscript.echo FreeIP
	wscript.quit(0)