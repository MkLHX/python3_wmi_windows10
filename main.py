"""
Create 2020-02-27
Author: Mickael Lehoux
Email: mickael@lehoux.net
Based on https://blog.ipswitch.com/managing-windows-system-administration-with-wmi-and-python
And https://www.activexperts.com/admin/scripts/wmi/python/
And https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/cimwin32-wmi-providers
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
    @return json file
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
    @return json file
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
    @return json file
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
    @return json file
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


"""
Getting temperatures depend of the computer
In certain cases you can use:
    - Win32_TemperatureProbe class or
    - Win32_CurrentTemp class or
    - Or call OpenHardwareMonitor https://openhardwaremonitor.org/documentation/
"""


def get_all_temperature(cnx):
    """
    List all temperatures
    @param cnx => wmi.WMI() instance
    @return json file
    """
    try:
        print("\n\r### list all temperature ###\n\r")
        temps = {}
        i = 0
        for temp in cnx.Win32_TemperatureProbe():
            # print(temp)
            temps[i] = {'temperature': temp}
        i += 1
        return json.dumps(temps)
    except BaseException as e:
        print('An exception occured {}'.format(e))


def get_all_temperature_OHM():
    w = wmi.WMI(namespace="root\OpenHardwareMonitor")
    temperature_infos = w.Sensor()
    temps = {}
    i = 0
    for sensor in temperature_infos:
        if sensor.SensorType == u'Temperature':
            print(sensor.Name, sensor.Value)
            temps[i] = {sensor.Name: sensor.Value}
        i += 1
    return json.dumps(temps)


if __name__ == "__main__":
    """
    List WMI informations
    """
    # get_all_wmi_classes()
    # get_all_wmi_methods()

    """
    Uncomment line to get data you need
    """
    # connect to the machine
    """
    Connect to the local machine
    """
    local_cnx = wmi.WMI()
    """
    Connect to a remote machine on the same network
    NOTA: instead of pass local_cnx on each method, use remote_cnx
    """
    # remote_cnx = wmi.WMI("<machine_ip>", user=r"<user_account>", password="<user_password>")

    # get_all_processes(local_cnx)
    # get_all_disks_space(local_cnx)
    # get_all_processors_data(local_cnx)
    get_all_computer_data(local_cnx)
    # get_all_temperature(local_cnx)
    get_all_temperature_OHM()
    pass
