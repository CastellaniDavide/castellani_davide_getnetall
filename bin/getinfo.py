"""getinfo
"""
import os, wmi
from datetime import datetime

__author__ = "davidecastellani@castellanidavide.it"
__version__ = "02.01 2020-05-18"

class getinfo:
	def __init__ (self, vs=False):
		"""The core of my project
		"""
		base_dir = "." if vs else ".." # the project "root" in Visual studio it is different

		log = open(os.path.join(base_dir, "log", "log.log"), "a")
		csv_client = open(os.path.join(base_dir, "flussi", "getinfo_client.csv"), "w+")
		csv_server = open(os.path.join(base_dir, "flussi", "getinfo_server.csv"), "w+")
		getinfo.log(log, "Opened all files")
		getinfo.init_csv(csv_client, csv_server, log)
		
		start_time = datetime.now()
		getinfo.log(log, f"Start time: {start_time}")
		getinfo.log(log, "Running: getinfo.py")

		conn = wmi.WMI()
		getinfo.print_client(conn, log, csv_client)
		getinfo.print_protocol(conn, log, csv_server)

		getinfo.log(log, f"End time: {datetime.now()}\nTotal time: {datetime.now() - start_time}")
		getinfo.log(log, "")
		log.close()

	def print_client(conn, log, csv):
		"""Prints the infos by Win32_NetworkClient
		"""
		getinfo.log(log, " - Istruction: Win32_NetworkClient")
		for network_client in conn.Win32_NetworkClient(["Caption", "Description", "Status", "Manufacturer", "Name"]):
			print(network_client)
			getinfo.log(log, f"   - {network_client.Caption}")
			csv.write(str(network_client.Caption) + ',' + str(network_client.Description) + ',' + str(network_client.Status) + ',' + str(network_client.Manufacturer) + ',' + str(network_client.Name) + '\n')

	def print_protocol(conn, log, csv):
		"""Prints the infos by Win32_NetworkProtocol
		"""
		getinfo.log(log, " - Istruction: Win32_NetworkProtocol")
		for network_protocol in conn.Win32_NetworkProtocol(["Caption", "Description", "GuaranteesDelivery", "GuaranteesSequencing", "MaximumAddressSize", "MaximumMessageSize", "Name", "SupportsConnectData", "SupportsEncryption", "SupportsEncryption", "SupportsGracefulClosing", "SupportsGuaranteedBandwidth", "SupportsQualityofService"]):
			print(network_protocol)
			getinfo.log(log, f"   - {network_protocol.Caption}")
			csv.write(str(network_protocol.Caption) + ',' + str(network_protocol.Description) + ',' + str(network_protocol.GuaranteesDelivery) + ',' + str(network_protocol.GuaranteesSequencing) + ',' + str(network_protocol.MaximumAddressSize) + ',' + str(network_protocol.MaximumMessageSize) + ',' + str(network_protocol.Name) + ',' + str(network_protocol.SupportsConnectData) + ',' + str(network_protocol.SupportsEncryption) + ',' + str(network_protocol.SupportsEncryption) + ',' + str(network_protocol.SupportsGracefulClosing) + ',' + str(network_protocol.SupportsGuaranteedBandwidth) + ',' + str(network_protocol.SupportsQualityofService) + '\n')

	def log(file, item):
		"""Writes a line in the log.log file
		"""
		file.write(f"{item}\n")

	def print_and_log(file, item):
		"""Writes on the screen and in the log file
		"""
		print(item)
		getinfo.log(file, item)

	def init_csv(client, server, log):
		"""Init the csv files
		"""
		client.write("Caption,Description,Status,Manufacturer,Name\n")
		server.write("Caption,Description,GuaranteesDelivery,GuaranteesSequencing,MaximumAddressSize,MaximumMessageSize,Name,SupportsConnectData,SupportsEncryption,SupportsEncryption,SupportsGracefulClosing,SupportsGuaranteedBandwidth,SupportsQualityofService\n")
		getinfo.log(log, "csv now initialized")

if __name__ == "__main__":
	# degub flag
	debug = True

	getinfo(debug)