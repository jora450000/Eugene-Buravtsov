const net_lan = "192.168.1."  'set your lan subnet with mask /24 here
const strArcFolder = "l:\arc\shares"  'set your folder for backup files here


for i = 2 to 254 

	strComputer = net_lan + CStr(i)
        On Error Resume Next 

	Set objWMIService = GetObject("winmgmts:" _
	    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	if Err.Number = 0  then
		Set colShares = objWMIService.ExecQuery("Select * from Win32_Share")

		For each objShare in colShares
'		    Wscript.Echo "Allow Maximum: " & objShare.AllowMaximum   
'		    Wscript.Echo "Caption: " & objShare.Caption   
'		    Wscript.Echo "Maximum Allowed: " & objShare.MaximumAllowed
'		    Wscript.Echo "Name: " & objShare.Name   
'		    Wscript.Echo "Path: " & objShare.Path   
'		    Wscript.Echo "Type: " & objShare.Type   
		    if InStr(objShare.Name, "$") = 0  and objShare.Type = 0 then  	
			    strShare = "\\" & strComputer & "\" & objShare.Name   
			    strArcFile =  chr(34) & strArcFolder & "\" & strComputer& "_" & objShare.Name  & ".rar" &  chr(34)
			    strCommand= chr (34) & "c:\program files (x86)\winrar\rar.exe" & chr(34) & "  a -r -tn1d  -ms*.docx;*.xlsx  " &  strArcFile & "  " & chr(34) &strShare & "\*.doc*"& chr(34) & " " & chr(34) & strShare & "\*.xls*" & chr(34)
                            Wscript.Echo strCommand
  			    Set WshShell = CreateObject("WScript.Shell")
			    Set WshExec = WshShell.Exec(strCommand)			    	
  
		    end if	
  		Next
	else
		Wscript.Echo "no access or not exist " & strComputer
	end if

next 
