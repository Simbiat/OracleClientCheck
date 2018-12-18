Option Explicit
Dim objFSO, WshShell, oproduct, objShell, msgresult

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")
Set objShell = CreateObject("Shell.Application")

oproduct = WshShell.Environment("Process").Item("ORACLE_HOME")
if oproduct = "" then
	msgresult = msgbox ("No Oracle Client detected!" + vbcrlf + "Please, raise 'Request' request to get it." + vbcrlf + "Press 'OK' to open request page.", vbOKCancel+vbCritical+vbSystemModal, "Oracle client not found!")
	if msgresult = vbOk then
		WshShell.Run ("https://request.com")
	end if
	WScript.quit 99
end if

dim orpath, orpaths, pathnum, or64
orpaths = Split(oproduct, ";")
pathnum = 0
or64 = false

for each orpath in orpaths
	if objFSO.FileExists(orpath+ "\bin\oci.dll") AND objFSO.FileExists(orpath+ "\inventory\ContentsXML\comps.xml") Then
		dim oraclexml
		set oraclexml = objFSO.OpenTextFile(orpath+ "\inventory\ContentsXML\comps.xml", 1)
		If InStr(1, oraclexml.ReadAll, "NT_AMD64", vbTextCompare) > 0 Then
			or64 = true
			oraclexml.Close
		else
			oraclexml.Close
			WScript.quit pathnum
		end if
	end if
	pathnum = pathnum + 1
next

if or64 = true then
	msgresult = msgbox ("64-bit version of Oracle Client detected and is not currently supported!" + vbcrlf + "Please, raise 'Request' request to get 32-bit version." + vbcrlf + "Press 'OK' to open request page.", vbOKCancel+vbCritical+vbSystemModal, "Wrong Oracle Client version!")
	if msgresult = vbOk then
		WshShell.Run ("https://request.com")
	end if
	WScript.quit 99
else
	msgresult = msgbox ("No Oracle Client detected!" + vbcrlf + "Please, raise 'Request' request to get it." + vbcrlf + "Press 'OK' to open request page.", vbOKCancel+vbCritical+vbSystemModal, "Oracle client not found!")
	if msgresult = vbOk then
		WshShell.Run ("https://request.com")
	end if
	WScript.quit 99
end if
