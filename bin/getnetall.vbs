' getnetall
' Get most important network adapter configuration infos
' author = "help@castellanidavide.it"
' version = "v01.01 2020-05-10"

' Option Explicit

Dim strComputer

strComputer = "."

Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colScheduledJobs = objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration")

For Each objJob in colScheduledJobs
    Wscript.Echo "Caption: " & vbTab & vbTab & vbTab & objJob.Caption
	Wscript.Echo "Description: " & vbTab & vbTab & vbTab & objJob.Description
	Wscript.Echo "SettingID: " & vbTab & vbTab & vbTab & objJob.SettingID
	Wscript.Echo "DatabasePath: " & vbTab & vbTab & vbTab & objJob.DatabasePath
	Wscript.Echo "DNSHostName: " & vbTab & vbTab & vbTab & objJob.DNSHostName
	Wscript.Echo "DomainDNSRegistrationEnabled: " & vbTab & objJob.DomainDNSRegistrationEnabled
	Wscript.Echo "FullDNSRegistrationEnabled: " & vbTab & objJob.FullDNSRegistrationEnabled
	Wscript.Echo "MACAddress: " & vbTab & vbTab & vbTab & objJob.MACAddress
	Wscript.Echo "ServiceName: " & vbTab & vbTab & vbTab & objJob.ServiceName
	Wscript.Echo "TcpWindowSize: " & vbTab & vbTab & vbTab & objJob.TcpWindowSize
	Wscript.Echo "" ' Empty line
Next
