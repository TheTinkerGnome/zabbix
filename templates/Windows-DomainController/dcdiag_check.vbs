'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'dcdiag_check.vbs
'Version 1.0
'Written By: David W. Gilmore
'Created On: 03/10/2015
'
' 
'Purpose: Run the specified dcdiag.exe test as specified from the command line parameter and returns 
'			pass or fail status. The test returns 1 for pass and 0 for fail
'

'THis is the function that runs the dcdiag test
Function runDiagTest(strTest)
	strResult = "0"
	cmd = "cmd.exe /c dcdiag /test:" & strTest
	Set sh = WScript.CreateObject("WScript.Shell")
	Set exec =  sh.Exec(cmd)
	Do While exec.Status = 0
		do While not Exec.StdOut.AtEndOfStream
			tmpLine = lcase(exec.StdOut.ReadLine())           	
			if instr(tmpLine, lcase(strOK) & " test " & strTest) then
				'we have a strOK String which means we have reached the end of a result output (maybe on newline)
				strResult = "1"
			end if
		Loop
	Loop
	runDiagTest = strResult
End Function

'Lang dependend. Default is english
dim strOK : strOK = "passed"
dim strNotOK : strNotOk = "failed"

'Main execution section

'Set the test name based on the command line parameter
strOption = WScript.Arguments(0)
strTestResult = runDiagTest(strOption)

wscript.echo strTestResult

