' getinfo
' Get most important network adapter configuration infos
' author = "help@castellanidavide.it"
' version = "v02.01 2020-05-18"

' Option Explicit

' open log file
Set log_file = CreateObject("Scripting.FileSystemObject").OpenTextFile("../log/log.log", 8, true) ' 8 => appening
log_file.WriteLine("Start time: " & Now())
log_file.WriteLine("Running: getinfo.vbs")

' open csv file (client)
Set csv_file_client = CreateObject("Scripting.FileSystemObject").OpenTextFile("../flussi/getinfo_client.csv", 2, true) ' 2 => write
csv_file_client.WriteLine("Caption,Description,Status,Manufacturer,Name")

' open csv file (server)
Set csv_file_server = CreateObject("Scripting.FileSystemObject").OpenTextFile("../flussi/getinfo_server.csv", 2, true) ' 2 => write
csv_file_server.WriteLine("Caption,Description,GuaranteesDelivery,GuaranteesSequencing,MaximumAddressSize,MaximumMessageSize,Name,SupportsConnectData,SupportsEncryption,SupportsEncryption,SupportsGracefulClosing,SupportsGuaranteedBandwidth,SupportsQualityofService")


Dim strComputer

strComputer = "."

log_file.WriteLine("Testing computer: " & strComputer)

Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Dim istructions
istructions = Array("Select * from Win32_NetworkClient", "Select * from Win32_NetworkProtocol")

For Each istruction in istructions

	log_file.WriteLine(" - Istruction: " & istruction)
	Set colScheduledJobs = objWMIService.ExecQuery(istruction)

	For Each objJob in colScheduledJobs
		log_file.WriteLine("   - Testing: " & objJob.Caption)
		If istruction = istructions(0) Then
			Wscript.Echo "Caption: " & vbTab & vbTab & vbTab & objJob.Caption
			Wscript.Echo "Description: " & vbTab & vbTab & vbTab & objJob.Description
			Wscript.Echo "Status: " & vbTab & vbTab & vbTab & objJob.Status
			Wscript.Echo "Manufacturer: " & vbTab & vbTab & vbTab & objJob.Manufacturer
			Wscript.Echo "Name: " & vbTab & vbTab & vbTab & vbTab & objJob.Name
			Wscript.Echo "" ' Empty line
			csv_file_client.WriteLine(objJob.Caption & "," & objJob.Description & "," & objJob.Status & "," & objJob.Manufacturer & "," & objJob.Name)
		Else
			Wscript.Echo "Caption: " & vbTab & vbTab & vbTab & objJob.Caption
			Wscript.Echo "Description: " & vbTab & vbTab & vbTab & objJob.Description
			Wscript.Echo "GuaranteesDelivery: " & vbTab & vbTab & objJob.GuaranteesDelivery
			Wscript.Echo "GuaranteesSequencing: " & vbTab & vbTab & objJob.GuaranteesSequencing
			Wscript.Echo "MaximumAddressSize: " & vbTab & vbTab & objJob.MaximumAddressSize
			Wscript.Echo "MaximumMessageSize: " & vbTab & vbTab & objJob.MaximumMessageSize
			Wscript.Echo "Name: " & vbTab & vbTab & vbTab & vbTab & objJob.Name
			Wscript.Echo "SupportsConnectData: " & vbTab & vbTab & objJob.SupportsConnectData
			Wscript.Echo "SupportsEncryption: " & vbTab & vbTab & objJob.SupportsEncryption
			Wscript.Echo "SupportsEncryption: " & vbTab & vbTab & objJob.SupportsEncryption
			Wscript.Echo "SupportsGracefulClosing: " & vbTab & objJob.SupportsGracefulClosing
			Wscript.Echo "SupportsGuaranteedBandwidth: " & vbTab & objJob.SupportsGuaranteedBandwidth
			Wscript.Echo "SupportsQualityofService: " & vbTab & objJob.SupportsQualityofService
			Wscript.Echo "" ' Empty line
			csv_file_server.WriteLine(objJob.Caption & "," & objJob.Description & "," & objJob.GuaranteesDelivery & "," & objJob.GuaranteesSequencing & "," & objJob.MaximumAddressSize & "," & objJob.MaximumMessageSize & "," & objJob.Name & "," & objJob.SupportsConnectData & "," & objJob.SupportsEncryption & "," & objJob.SupportsEncryption & "," & objJob.SupportsGracefulClosing & "," & objJob.SupportsGuaranteedBandwidth & "," & objJob.SupportsQualityofService)
		End If
		
	Next
Next

' close log file
log_file.WriteLine("End time: " & Now())
log_file.WriteLine()
log_file.Close
Set log_file = Nothing

' close csv files
csv_file_client.Close
csv_file_server.Close
