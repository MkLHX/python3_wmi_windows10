"""
Create 2020-02-27
Author: Mickael Lehoux
Email: mickael@lehoux.net
Based on https://blog.ipswitch.com/managing-windows-system-administration-with-wmi-and-python
And https://www.activexperts.com/admin/scripts/wmi/python/
"""
import wmi
import json


def get_all_wmi_classes():
    """
    List all wmi classes you can call
    """
    try:
        print("\n\r### list all classes ###\n\r")
        for wmi_class in wmi.WMI().Win32_Process.methods.keys():
            print(wmi_class)
    except BaseException as e:
        print('An exception occured {}'.format(e))


def get_all_wmi_methods():
    """
    List all wmi classes you can use
    """
    try:
        print("\n\r### list all methods ###\n\r")
        for wmi_method in wmi.WMI().Win32_Process.properties.keys():
            print(wmi_method)
    except BaseException as e:
        print('An exception occured {}'.format(e))


def get_all_processes(cnx, process=None):
    """
    List all processes for connection passed on argument.
    @param cnx => wmi.WMI() instance
    @param process => string filter process name e.g. "chrome.exe" by default None show all processes
    """
    try:
        print("\n\r### list all processes ###\n\r")
        processes = {}
        i = 0
        # name=process if None is not process else ''
        for process in cnx.Win32_Process():
            # print(process)
            print("ID: {0}\nHandleCount: {1}\nProcessName: {2}\n".format(
                process.ProcessId, process.HandleCount, process.Name))
            processes[i] = {'ID': process.ProcessId,
                            'HandleCount': process.HandleCount,
                            'ProcessName': process.Name}
            i += 1
        return json.dumps(processes)
    except BaseException as e:
        print('An exception occured {}'.format(e))


def get_all_disks_space(cnx, disk=None):
    """
    List all disks free space
    @param cnx => wmi.WMI() instance
    @param disk => string filter disk by root caption e.g. "c:" by default None show all disks space
    """
    try:
        print("\n\r### list all disks space ###\n\r")
        disks = {}
        i = 0
        # name=disk if None is not disk else ''
        for disk in cnx.Win32_LogicalDisk():
            # print(disk)
            if None is not disk.Size:
                print("%s(%s) is %.2f%% free, %.2f / %.2f" % (disk.VolumeName, disk.Caption, (100*float(
                    disk.FreeSpace)/float(disk.Size)), float(disk.FreeSpace), float(disk.Size)))
                disks[i] = {'disk_volume_name': disk.VolumeName,
                            'disk_letter': disk.Caption,
                            'disk_free_space': 100*float(disk.FreeSpace)/float(disk.Size),
                            'free_space': disk.FreeSpace,
                            'disk_size': disk.Size,
                            }
            i += 1
        return json.dumps(disks)
    except BaseException as e:
        print('An exception occured {}'.format(e))


def get_all_processors_data(cnx):
    """
    List processors data
    @param cnx => wmi.WMI() instance
    """
    try:
        print("\n\r### list processors data ###\n\r")
        data_proc = {}
        i = 0
        for proc in cnx.Win32_Processor():
            # print(proc)
            print(proc.Caption, proc.Name, proc.DeviceID,
                  proc.LoadPercentage)
            data_proc[i] = {'proc_caption': proc.Caption,
                            'proc_name': proc.Name,
                            'proc_device_id': proc.DeviceID,
                            'proc_load_percentage': proc.LoadPercentage}
            i += 1
        return json.dumps(data_proc)
    except BaseException as e:
        print('An exception occured {}'.format(e))


def get_all_computer_data(cnx):
    """
    List computer data
    @param cnx => wmi.WMI() instance
    """
    try:
        print("\n\r### list computer data ###\n\r")
        data_computer = {}
        i = 0
        for data in cnx.Win32_ComputerSystem():
            # print(data)
            print("Manufacturer %s, Model %s, SKU Number %s, System type %s" % (
                data.Manufacturer, data.Model, data.SystemSKUNumber, data.SystemType))
            data_computer[i] = {'manufacturer': data.Manufacturer,
                                'model': data.Model,
                                'sku_number': data.SystemSKUNumber,
                                'system_type': data.SystemType}
            i += 1
        return json.dumps(data_computer)
    except BaseException as e:
        print('An exception occured {}'.format(e))


# def get_all_temperature(cnx):
#     """
#     List all temperatures
#     @param cnx => wmi.WMI() instance
#     """
#     try:
#         print("\n\r### list all temperature ###\n\r")
#         temps = {}
#         i = 0
#         for temp in cnx.Win32_TemperatureProbe():
#             # print(temp)
#             temps[i] = {}
#         i += 1
#     except BaseException as e:
#         print('An exception occured {}'.format(e))


# connect to the machine
local_cnx = wmi.WMI()
# remote_cnx = wmi.WMI("13.76.128.231", user=r"prateek", password="P@ssw0rd@123")
# get_all_wmi_classes()
# get_all_wmi_methods()
# get_all_processes(local_cnx)
# get_all_disks_space(local_cnx)
# get_all_processors_data(local_cnx)
# get_all_computer_data(local_cnx)
# get_all_temperature(local_cnx)

# w = wmi.WMI(namespace="root\wmi")
# temperature_info = w.MSAcpi_ThermalZoneTemperature()[0]
# print(temperature_info)


# https://openhardwaremonitor.org/documentation/
# w = wmi.WMI(namespace="root\OpenHardwareMonitor")

# print(w.Sensor())
# temperature_info = w.Sensor()
# for sensor in temperature_info:
#     if sensor.SensorType==u'Temperature':
#         print(sensor.Name)
#         print(sensor.Value)
