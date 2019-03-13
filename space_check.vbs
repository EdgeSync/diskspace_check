' Script to check Disk Space on Windows Systems - Sends email if disk is < 10%
' Author: Chris Staunton
' Email: redacted
' Github: https://github.com/EdgeSync/
' Date: 2019 March 13th

Function GetCurrentComputerName
 set oWsh = WScript.CreateObject("WScript.Shell")
 set oWshSysEnv = oWsh.Environment("PROCESS")
 GetCurrentComputerName = oWshSysEnv("COMPUTERNAME")
End Function

Function GetCurrentFQDN
 set sysInfo = CreateObject("ADSystemInfo")
 Set wshNetwork = CreateObject("WScript.Network")
 GetCurrentFQDN = sysInfo.DomainDNSName
End Function

Function SendAlertMail
 Set mail=CreateObject("CDO.Message")
 mail.Subject="Alert - Low Disk Space Detected - " & strComputerName & " - " & percleft & "% remaining" 
 mail.From="noreply_" & strComputerName & "@whateverdomain.com"
 mail.To="youremailhere@whateverdomain.com" ' Replace this with the mailbox/pdl you want to send to
 'mail.CC="email1@whateverdomain.com; email2@whateverdomain.com" ' How to specify additional addresses in CC, and multiple addresses on one line
 mail.TextBody="Alert - Low Disk Space Detected." & vbNewLine & vbNewLine & "Server: " & strFQDN & vbNewLine & "Affected Drive: " & drive & ":\" & vbNewLine & "Space Remaining (%): " & percleft & "%" & vbNewLine & vbNewLine & "This requires attention."
 mail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.whateverdomain.com"
 mail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25 
 mail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
 mail.Configuration.Fields.Update
 mail.Send
End Function

set mail=nothing
 
'====================================================================================
' Begin main code below
'====================================================================================

' Constants for drive types, we are only using Fixed, but leaving as reference.
Const Unknown = 0
Const Removable = 1
Const Fixed = 2
Const Remote = 3
Const CDROM = 4
Const RAMDisk = 5

str = ""
set oFs = WScript.CreateObject("Scripting.FileSystemObject")
set oDrives = oFs.Drives

strComputerName = GetCurrentComputerName ' get name only once for performance reasons
strDomain = GetCurrentFQDN
strFQDN = strComputerName & "." & strDomain

for each oDrive in oDrives
 Select case oDrive.DriveType
 Case Fixed 'You can replace Fixed with any of the drive constants from above- or create an array amd loop through the options.
 drive = oDrive.DriveLetter
 total = oDrive.TotalSize
 free = oDrive.FreeSpace
 percleft = free/total * 100
 percleft = Round(percleft)
 
 
 IF percleft < 10 THEN
	'Wscript.Echo "FQDN: " & strFQDN & vbNewLine & "This Drive has " & percleft & "% space. Action Needed." ' Can be used for testing - will pop up a box with variables first.
	SendAlertMail
 End IF
 
 End Select
next
