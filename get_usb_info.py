from winreg import *
from Evtx.Evtx import FileHeader
from Evtx.Views import evtx_file_xml_view
from xml.dom import minidom

import sys
import mmap
import contextlib
import csv


def get_usb_info_from_registry():
	global device_name, serial_number, volume_name, drive_name
	varSubkey = "SYSTEM\\ControlSet001\\Enum\\USBSTOR"
	varReg = ConnectRegistry(None, HKEY_LOCAL_MACHINE)
	varKey = OpenKey(varReg, varSubkey)
	device_name = []
	serial_number = []
	volume_name = []
	drive_name = []
	f = open('usb_info.csv', 'w', encoding='utf-8', newline='')
	wr = csv.writer(f)
	for i in range(1024):
		try:
			keyname = EnumKey(varKey, i)
			varSubkey2 = "%s\\%s" % (varSubkey, keyname)
			varKey2 = OpenKey(varReg, varSubkey2)
			for j in range(1024):
				try:
					keyname2 = EnumKey(varKey2, j)
					varSubkey3 = "%s\\%s" % (varSubkey2, keyname2)
					varKey3 = OpenKey(varReg, varSubkey3)
					serial_number.append(keyname2)
					for k in range(1024):
						try:
							n, v, t = EnumValue(varKey3, k)
							if n == "FriendlyName":
								device_name.append(v)
						except:
							errorMsg = "Exception Inner:", sys.exc_info()[0]
				except:
					errorMsg = "Exception Inner:", sys.exc_info()[0]
		except:
			errorMsg = "Exception Outter:", sys.exc_info()[0]
			break
	#print("device_name:", device_name, "\n", "serial_number: ", serial_number) #장치명, 시리얼 넘버
	CloseKey(varKey3)
	CloseKey(varKey2)
	CloseKey(varKey)
	CloseKey(varReg)


	#get volume name
	varSubkey = "SYSTEM\\ControlSet001\\Enum\\WpdBusEnumRoot"
	varReg = ConnectRegistry(None, HKEY_LOCAL_MACHINE)
	varKey = OpenKey(varReg, varSubkey)

	for i in range(1024):
		try:
			keyname = EnumKey(varKey, i)
			varSubkey2 = "%s\\%s" % (varSubkey, keyname)
			varKey2 = OpenKey(varReg, varSubkey2)
			for j in range(1024):
				try:
					keyname2 = EnumKey(varKey2, j)
					varSubkey3 = "%s\\%s" % (varSubkey2, keyname2)
					varKey3 = OpenKey(varReg, varSubkey3)
					for k in range(1024):
						try:
							n, v, t = EnumValue(varKey3, k)
							if n == "FriendlyName":
								volume_name.append(v)
						except:
							errorMsg = "Exception Inner:", sys.exc_info()[0]
				except:
					errorMsg = "Exception Inner:", sys.exc_info()[0]
		except:
			errorMsg = "Exception Outter:", sys.exc_info()[0]
			break
	#print("volume_name: ", volume_name) #볼륨명
	CloseKey(varKey3)
	CloseKey(varKey2)
	CloseKey(varKey)
	CloseKey(varReg)


	#get drive name
	varSubkey = "SOFTWARE\\Microsoft\\Windows Portable Devices\\Devices"
	varReg = ConnectRegistry(None, HKEY_LOCAL_MACHINE)
	varKey = OpenKey(varReg, varSubkey)

	for i in range(1024):
		try:
			keyname = EnumKey(varKey, i)
			varSubkey2 = "%s\\%s" % (varSubkey, keyname)
			varKey2 = OpenKey(varReg, varSubkey2)
			try:
				for k in range(1024):
					if keyname.split('#')[-2] in serial_number:
						n, v, t = EnumValue(varKey2, k)
						if n == "FriendlyName":
							drive_name.append(v)
			except:
				errorMsg = "Exception Inner:", sys.exc_info()[0]
		except:
			errorMsg = "Exception Outter:", sys.exc_info()[0]
			break
	#print("drive_name: ",drive_name)
	i = 0
	try:
		while(True):
			wr.writerow([serial_number[i], device_name[i], volume_name[i], drive_name[i]])
			i += 1	
	except:
		pass
	f.close()

def get_access_time_from_setupapi():
	global get_access_time
	f = open("setupapi.dev.log.txt", "r")
	setupapi_content = f.read()
	get_access_time = []

	#test
	#test_number = ['000FF1103111925014061123&0', '547640570&0', 'KZ5DCQC0153_________&0', 'AA00000000001955&0']
	#for serial in test_number: #test
	for serial in serial_number:
		serial_index = setupapi_content.find(serial)
		start_index = setupapi_content.find("start", serial_index)
		access_time_index = start_index + 6
		get_access_time.append(setupapi_content[access_time_index:access_time_index+23])
	f.close()
	#print("get_access_time: ", get_access_time)


def get_connect_disconnect_pair_from_evtx():
	global connect_time_serial, disconnect_time_serial, life_time
	connect_time = []
	disconnect_time = []
	life_time = []
	outFile = open("test.txt", "a+")
	with open("Microsoft-Windows-DriverFrameworks-UserMode%4Operational.evtx", "r") as f:
		with contextlib.closing(mmap.mmap(f.fileno(), 0, access=mmap.ACCESS_READ)) as buf:
			fh = FileHeader(buf, 0x0)
			for xml, record in evtx_file_xml_view(fh):
				xmldoc = minidom.parseString(xml)
				event_id = xmldoc.getElementsByTagName('EventID')[0].childNodes[0].nodeValue
				event_time = str(xmldoc.getElementsByTagName('TimeCreated')[0].toxml()).split('"')[1]
				userdata = str(xmldoc.getElementsByTagName('UserData')[0].toxml()).replace("&amp;","&")
				if event_id == str("2003"): #Connect
					#for serial in test_field: #test
					for serial in serial_number:
						if serial in userdata:
							serial_start = userdata.find(serial)
							#serial_end = userdata[serial_start:].find('#')
							length = len(serial)
							connect_time_serial = userdata[serial_start:len(serial)]
							connect_time.append(xmldoc.getElementsByTagName('TimeCreated')[0].toxml())
							event_life_time_start = userdata.find("lifetime") + 11
							event_life_time_end = userdata.find('}" xmlns')
							event_life_time = userdata[event_life_time_start:event_life_time_end]
							life_time.append(event_life_time)
				elif event_id == str("2100"): #Disconnect
					#for serial in test_field: #test
					for serial in serial_number:
						if serial in userdata:
							serial_start = userdata.find(serial)
							#serial_end = userdata[serial_start:].find('#')
							length = len(serial)
							disconnect_time_serial = userdata[serial_start:len(serial)]
							disconnect_time.append(xmldoc.getElementsByTagName('TimeCreated')[0].toxml())
			#print("connect_time_serial: ", connect_time_serial)
			#print("disconnect)time_serial: ", disconnect_time_serial)
			#print("connect_time: ", connect_time)
			#print("disconnect_time: ", disconnect_time)
	f = open('evtx_info.csv', 'w', encoding='utf-8', newline='')
	wr = csv.writer(f)
	i = 0
	try:
		while(True):
			wr.writerow([connect_time_serial[i], disconnect_time_serial[i], connect_time[i], disconnect_time[i]])
			i += 1	
	except:
		pass
	f.close()

def get_shellbag_info():
	shellbags  = []
	bagmru_key = shell_key.subkey("BagMRU")
	bags_key   = shell_key.subkey("Bags")
	def shellbag_rec(key, bag_prefix, path_prefix):
		# First, consider the current key, and extract shellbag items
		slot = key.value("NodeSlot").value()

		# Look at ..\Shell, and ..\Desktop, etc.
		for bag in bags_key.subkey( slot ).subkeys():

			# Only consider ITEMPOS keys
			for value in [value for value in bag.values() if \
					"ItemPos" in value.name()]:

				# Call our binary processing code to extract items
				new_items = process_itempos(value)
				for item in new_items:
					shellbags.append(path_prefix + item.path)

		# Next, recurse into each subkey of this BagMRU node (1, 2, 3, ...)
		for value in [value for value in key.values()]:

			# Call our binary processing code to extract item
			new_item = process_bag(value)
			shellbags.append(path_prefix + new_item.path)

			shellbag_rec(key.subkey( value.name() ), 
						bag_prefix + "\\" + value.name(), new_item.path )

	shellbag_rec("HKEY_USERS\\{USERID}\\Software\\Microsoft\\Windows\\ShellNoRoam", 
					"", 
					"")
    #print(shellbags)
	return shellbags


if __name__ == "__main__":
	get_usb_info_from_registry()
	get_access_time_from_setupapi()
	get_connect_disconnect_pair_from_evtx()
	#get_shellbag_info()

