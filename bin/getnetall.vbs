' getnetall
' Get most important network adapter configuration infos
' author = "help@castellanidavide.it"
' version = "v02.01 2020-05-18"

' Option Explicit

' open log file
Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile("../log/log.log", 8, true) ' 8 => appening
objFileToWrite.WriteLine("Start time: " & Now())
objFileToWrite.WriteLine("Running: getnetall.vbs")

' open csv file
Set csv = CreateObject("Scripting.FileSystemObject").OpenTextFile("../flussi/getnetall.csv", 2, true) ' 2 => write
csv.WriteLine("Caption,Description,SettingID,DatabasePath,DNSHostName,DomainDNSRegistrationEnabled,FullDNSRegistrationEnabled,MACAddress,ServiceName,TcpWindowSize")

Dim strComputer

strComputer = "."

objFileToWrite.WriteLine("Testing computer: " & strComputer)

Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Dim istruction
istruction = "Select * from Win32_NetworkAdapterConfiguration"

objFileToWrite.WriteLine(" - Istruction: " & istruction)
Set colScheduledJobs = objWMIService.ExecQuery(istruction)

For Each objJob in colScheduledJobs
	objFileToWrite.WriteLine("   - Testing: " & objJob.Caption)
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
	csv.WriteLine(objJob.Caption & "," & objJob.Description & "," & objJob.SettingID & "," & objJob.DatabasePath & "," & objJob.DNSHostName & "," & objJob.DomainDNSRegistrationEnabled & "," & objJob.FullDNSRegistrationEnabled & "," & objJob.MACAddress & "," & objJob.ServiceName & "," & objJob.TcpWindowSize)
Next

' close log file
objFileToWrite.WriteLine("End time: " & Now())
objFileToWrite.WriteLine()
objFileToWrite.Close
Set objFileToWrite = Nothing

' close csv file
csv.Close
