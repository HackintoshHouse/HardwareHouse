import platform
import wmi
import psutil
import cpuinfo
import json
import csv
import threading
import time
import customtkinter as ctk
from datetime import datetime

# Initialize WMI interface for Windows hardware queries
c = wmi.WMI()

# Set up customtkinter appearance and theme
ctk.set_appearance_mode("System")  # "Dark", "Light", or "System"
ctk.set_default_color_theme("blue")  # Themes: "blue", "dark-blue", "green"

# --------------------------
# Hardware Info Collectors
# --------------------------

def get_system_info():
    """Return basic OS and system info"""
    uname = platform.uname()
    return {
        "System": uname.system,
        "Node Name": uname.node,
        "Release": uname.release,
        "Version": uname.version,
        "Machine": uname.machine,
        "Processor": uname.processor,
        "Python Version": platform.python_version(),
    }

def get_cpu_info():
    """Return detailed CPU info"""
    cpu_data = {}
    try:
        for cpu in c.Win32_Processor():
            cpu_data = {
                "Name": cpu.Name.strip(),
                "Manufacturer": cpu.Manufacturer,
                "Cores (Physical)": cpu.NumberOfCores,
                "Threads (Logical)": cpu.NumberOfLogicalProcessors,
                "Max Clock Speed (MHz)": cpu.MaxClockSpeed,
                "Current Clock Speed (MHz)": cpu.CurrentClockSpeed,
                "Architecture": cpu.Architecture,
                "Processor ID": cpu.ProcessorId,
                "L2 Cache Size (KB)": cpu.L2CacheSize,
                "L3 Cache Size (KB)": cpu.L3CacheSize,
            }
            break
    except Exception as e:
        cpu_data["Error"] = f"Failed to get CPU info: {str(e)}"
    return cpu_data

def get_gpu_info():
    """Return list of GPU info dicts"""
    gpus = []
    try:
        for gpu in c.Win32_VideoController():
            gpus.append({
                "Name": gpu.Name,
                "Driver Version": gpu.DriverVersion,
                "Video Processor": gpu.VideoProcessor,
                "RAM (MB)": int(gpu.AdapterRAM) // (1024*1024) if gpu.AdapterRAM else "Unknown",
                "Video Mode": gpu.VideoModeDescription,
                "Status": gpu.Status
            })
    except Exception as e:
        gpus.append({"Error": f"Failed to get GPU info: {str(e)}"})
    return gpus

def get_ram_info():
    """Return RAM info using psutil"""
    try:
        vm = psutil.virtual_memory()
        swap = psutil.swap_memory()
        return {
            "Total RAM (GB)": round(vm.total / (1024 ** 3), 2),
            "Available RAM (GB)": round(vm.available / (1024 ** 3), 2),
            "Used RAM (GB)": round(vm.used / (1024 ** 3), 2),
            "RAM Usage (%)": vm.percent,
            "Total Swap (GB)": round(swap.total / (1024 ** 3), 2),
            "Used Swap (GB)": round(swap.used / (1024 ** 3), 2),
            "Swap Usage (%)": swap.percent,
        }
    except Exception as e:
        return {"Error": f"Failed to get RAM info: {str(e)}"}

def get_disk_info():
    """Return disk partitions and usage"""
    disks = []
    try:
        partitions = psutil.disk_partitions()
        for p in partitions:
            usage = psutil.disk_usage(p.mountpoint)
            disks.append({
                "Device": p.device,
                "Mountpoint": p.mountpoint,
                "File System": p.fstype,
                "Total Size (GB)": round(usage.total / (1024 ** 3), 2),
                "Used (GB)": round(usage.used / (1024 ** 3), 2),
                "Free (GB)": round(usage.free / (1024 ** 3), 2),
                "Usage (%)": usage.percent,
            })
    except Exception as e:
        disks.append({"Error": f"Failed to get disk info: {str(e)}"})
    return disks

def get_network_info():
    """Return active network adapters info"""
    adapters = []
    try:
        addrs = psutil.net_if_addrs()
        stats = psutil.net_if_stats()
        for iface_name, iface_addrs in addrs.items():
            is_up = stats[iface_name].isup if iface_name in stats else False
            ips = [addr.address for addr in iface_addrs if addr.family == psutil.AF_INET]
            macs = [addr.address for addr in iface_addrs if addr.family == psutil.AF_LINK]
            adapters.append({
                "Interface": iface_name,
                "Status": "Up" if is_up else "Down",
                "IPv4 Addresses": ips if ips else ["None"],
                "MAC Addresses": macs if macs else ["None"],
            })
    except Exception as e:
        adapters.append({"Error": f"Failed to get network info: {str(e)}"})
    return adapters

def get_bios_info():
    """Return BIOS info from WMI"""
    bios_data = {}
    try:
        for bios in c.Win32_BIOS():
            bios_data = {
                "Manufacturer": bios.Manufacturer,
                "Version": bios.SMBIOSBIOSVersion,
                "Release Date": bios.ReleaseDate,
                "Serial Number": bios.SerialNumber,
                "BIOS Language": bios.BIOSLanguage if hasattr(bios, 'BIOSLanguage') else "N/A"
            }
            break
    except Exception as e:
        bios_data["Error"] = f"Failed to get BIOS info: {str(e)}"
    return bios_data

def get_motherboard_info():
    """Return motherboard info"""
    board_data = {}
    try:
        for board in c.Win32_BaseBoard():
            board_data = {
                "Manufacturer": board.Manufacturer,
                "Product": board.Product,
                "Serial Number": board.SerialNumber,
                "Version": board.Version,
                "Model": board.Model if hasattr(board, 'Model') else "N/A",
            }
            break
    except Exception as e:
        board_data["Error"] = f"Failed to get Motherboard info: {str(e)}"
    return board_data

def get_sound_devices():
    """Return list of sound devices info"""
    devices = []
    try:
        for sound in c.Win32_SoundDevice():
            devices.append({
                "Name": sound.Name,
                "Status": sound.Status,
                "Manufacturer": sound.Manufacturer
            })
    except Exception as e:
        devices.append({"Error": f"Failed to get Sound Devices info: {str(e)}"})
    return devices

def get_battery_info():
    """Return battery status info"""
    try:
        battery = psutil.sensors_battery()
        if battery:
            return {
                "Percent": f"{battery.percent}%",
                "Plugged In": battery.power_plugged,
                "Secs Left": battery.secsleft,
            }
        else:
            return {"Info": "No battery detected"}
    except Exception as e:
        return {"Error": f"Failed to get battery info: {str(e)}"}

def get_process_count():
    """Return number of running processes"""
    try:
        return {"Running Processes": len(psutil.pids())}
    except Exception as e:
        return {"Error": f"Failed to get process count: {str(e)}"}

def get_boot_time():
    """Return system boot time as human-readable string"""
    try:
        boot_timestamp = psutil.boot_time()
        boot_time = datetime.fromtimestamp(boot_timestamp).strftime("%Y-%m-%d %H:%M:%S")
        return {"Boot Time": boot_time}
    except Exception as e:
        return {"Error": f"Failed to get boot time: {str(e)}"}

# --------------------------
# Benchmark Functions
# --------------------------

def cpu_benchmark(duration=3):
    """Simple CPU benchmark by busy loop"""
    end = time.time() + duration
    count = 0
    while time.time() < end:
        count += 1  # Simulate work
        _ = count ** 0.5
    return {"CPU Benchmark (loops in {}s)".format(duration): count}

def memory_benchmark():
    """Simple memory benchmark by reading a large bytearray"""
    data = bytearray(100_000_000)
    start = time.time()
    s = sum(data)
    elapsed = time.time() - start
    return {"Memory Benchmark (sum bytes)": f"{elapsed:.4f} seconds"}

def disk_benchmark():
    """Simple disk write/read benchmark using temp file"""
    import tempfile
    import os
    try:
        tmp_file = tempfile.NamedTemporaryFile(delete=False)
        filename = tmp_file.name
        tmp_file.close()
        data = b"x" * (10**7)  # 10 MB
        start = time.time()
        with open(filename, "wb") as f:
            f.write(data)
        with open(filename, "rb") as f:
            content = f.read()
        elapsed = time.time() - start
        os.remove(filename)
        return {"Disk Benchmark (write/read 10MB)": f"{elapsed:.4f} seconds"}
    except Exception as e:
        return {"Disk Benchmark Error": str(e)}

# --------------------------
# Export Functions
# --------------------------

def export_json(data, filename="hardware_report.json"):
    """Export data dictionary to JSON file"""
    try:
        with open(filename, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4)
        return True
    except Exception as e:
        return str(e)

def export_csv(data, filename="hardware_report.csv"):
    """Export data dictionary to CSV file. Handles nested lists as multiple rows."""
    try:
        with open(filename, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)

            def write_dict(d, prefix=""):
                for k, v in d.items():
                    if isinstance(v, list):
                        writer.writerow([prefix + k])
                        for i, item in enumerate(v, 1):
                            if isinstance(item, dict):
                                writer.writerow([f"  Item {i}"])
                                write_dict(item, prefix + "    ")
                            else:
                                writer.writerow([f"  Item {i}: {item}"])
                    else:
                        writer.writerow([prefix + k, v])

            write_dict(data)
        return True
    except Exception as e:
        return str(e)

# --------------------------
# GUI Application Class
# --------------------------

class HardwareHouseApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("HardwareHouse - Made by HackintoshHouse")
        self.geometry("900x700")
        self.minsize(900, 700)

        # Create dropdown for categories
        self.category_options = [
            "System Info",
            "CPU Info",
            "GPU Info",
            "RAM Info",
            "Disk Info",
            "Network Info",
            "BIOS Info",
            "Motherboard Info",
            "Sound Devices",
            "Battery Info",
            "Process Count",
            "Boot Time",
            "Benchmarks",
        ]
        self.combo = ctk.CTkOptionMenu(self, values=self.category_options)
        self.combo.set("System Info")
        self.combo.pack(pady=10)

        # Buttons Frame
        btn_frame = ctk.CTkFrame(self)
        btn_frame.pack(pady=5)

        # Show Button
        self.show_btn = ctk.CTkButton(btn_frame, text="Show Info", command=self.show_info)
        self.show_btn.grid(row=0, column=0, padx=10)

        # Export JSON Button
        self.export_json_btn = ctk.CTkButton(btn_frame, text="Export JSON", command=self.export_json)
        self.export_json_btn.grid(row=0, column=1, padx=10)

        # Export CSV Button
        self.export_csv_btn = ctk.CTkButton(btn_frame, text="Export CSV", command=self.export_csv)
        self.export_csv_btn.grid(row=0, column=2, padx=10)

        # Textbox to display info
        self.textbox = ctk.CTkTextbox(self, width=860, height=530, font=("Segoe UI", 14))
        self.textbox.pack(pady=10)

        # Footer Label
        self.footer = ctk.CTkLabel(self, text="Made by HackintoshHouse", font=("Segoe UI", 10))
        self.footer.pack(side="bottom", pady=5)

        # Storage for current info
        self.current_info = {}

    def show_info(self):
        """Retrieve and show selected category info"""
        cat = self.combo.get()
        self.textbox.delete("0.0", "end")
        info = {}

        if cat == "System Info":
            info = get_system_info()
        elif cat == "CPU Info":
            info = get_cpu_info()
        elif cat == "GPU Info":
            info = {"GPUs": get_gpu_info()}
        elif cat == "RAM Info":
            info = get_ram_info()
        elif cat == "Disk Info":
            info = {"Disks": get_disk_info()}
        elif cat == "Network Info":
            info = {"Network Adapters": get_network_info()}
        elif cat == "BIOS Info":
            info = get_bios_info()
        elif cat == "Motherboard Info":
            info = get_motherboard_info()
        elif cat == "Sound Devices":
            info = {"Sound Devices": get_sound_devices()}
        elif cat == "Battery Info":
            info = get_battery_info()
        elif cat == "Process Count":
            info = get_process_count()
        elif cat == "Boot Time":
            info = get_boot_time()
        elif cat == "Benchmarks":
            self.textbox.insert("end", "Running CPU benchmark...\n")
            info = {}
            # Run benchmarks in thread to keep UI responsive
            def run_benchmarks():
                results = {}
                results.update(cpu_benchmark())
                results.update(memory_benchmark())
                results.update(disk_benchmark())
                self.current_info = results
                # Update UI in main thread
                self.textbox.after(0, lambda: self.display_info(results))

            threading.Thread(target=run_benchmarks).start()
            return
        else:
            info = {"Error": "Unknown category"}

        self.current_info = info
        self.display_info(info)

    def display_info(self, info):
        """Display info dict in the textbox"""
        self.textbox.delete("0.0", "end")

        def recursive_display(data, indent=0):
            indent_str = " " * (indent * 4)
            if isinstance(data, dict):
                for k, v in data.items():
                    if isinstance(v, (dict, list)):
                        self.textbox.insert("end", f"{indent_str}{k}:\n")
                        recursive_display(v, indent + 1)
                    else:
                        self.textbox.insert("end", f"{indent_str}{k}: {v}\n")
            elif isinstance(data, list):
                for i, item in enumerate(data, 1):
                    self.textbox.insert("end", f"{indent_str}- Item {i}:\n")
                    recursive_display(item, indent + 1)
            else:
                self.textbox.insert("end", f"{indent_str}{data}\n")

        recursive_display(info)

    def export_json(self):
        """Export current info to JSON file"""
        if self.current_info:
            res = export_json(self.current_info)
            if res is True:
                self.textbox.insert("end", "\nExported data to hardware_report.json\n")
            else:
                self.textbox.insert("end", f"\nError exporting JSON: {res}\n")
        else:
            self.textbox.insert("end", "\nNo data to export!\n")

    def export_csv(self):
        """Export current info to CSV file"""
        if self.current_info:
            res = export_csv(self.current_info)
            if res is True:
                self.textbox.insert("end", "\nExported data to hardware_report.csv\n")
            else:
                self.textbox.insert("end", f"\nError exporting CSV: {res}\n")
        else:
            self.textbox.insert("end", "\nNo data to export!\n")
# --------------------------
# Extended Hardware Info
# --------------------------

def get_usb_devices():
    """Return list of connected USB devices"""
    devices = []
    try:
        for usb in c.Win32_USBControllerDevice():
            dependent = usb.Dependent
            if dependent:
                dev = dependent.split('=')[1].strip('"')
                devices.append(dev)
    except Exception as e:
        devices.append({"Error": f"Failed to get USB devices: {str(e)}"})
    return devices

def get_display_monitors():
    """Return display monitor info"""
    monitors = []
    try:
        for monitor in c.Win32_DesktopMonitor():
            monitors.append({
                "Name": monitor.Name,
                "Screen Height": monitor.ScreenHeight,
                "Screen Width": monitor.ScreenWidth,
                "Status": monitor.Status
            })
    except Exception as e:
        monitors.append({"Error": f"Failed to get monitor info: {str(e)}"})
    return monitors

def get_printers():
    """Return installed printers info"""
    printers = []
    try:
        for printer in c.Win32_Printer():
            printers.append({
                "Name": printer.Name,
                "Status": printer.Status,
                "Default": printer.Default,
                "Network": printer.Network,
                "Shared": printer.Shared
            })
    except Exception as e:
        printers.append({"Error": f"Failed to get printers info: {str(e)}"})
    return printers

def get_installed_software():
    """Return list of installed software"""
    software_list = []
    try:
        for software in c.Win32_Product():
            software_list.append({
                "Name": software.Name,
                "Version": software.Version,
                "Vendor": software.Vendor,
                "Install Date": software.InstallDate
            })
    except Exception as e:
        software_list.append({"Error": f"Failed to get software list: {str(e)}"})
    return software_list

def get_power_plan():
    """Return current power plan GUID and name"""
    import subprocess
    try:
        output = subprocess.check_output("powercfg /getactivescheme", shell=True, text=True)
        # Output example: "Power Scheme GUID: xxx-xxx-xxx (Balanced)"
        line = output.strip()
        return {"Power Plan": line}
    except Exception as e:
        return {"Error": f"Failed to get power plan: {str(e)}"}

def get_system_locale():
    """Return system locale and timezone info"""
    import locale, time
    try:
        loc = locale.getdefaultlocale()
        tz = time.tzname
        return {
            "Locale": loc,
            "Timezone": tz
        }
    except Exception as e:
        return {"Error": f"Failed to get locale info: {str(e)}"}

def get_system_uptime():
    """Return system uptime in human readable format"""
    try:
        uptime_seconds = time.time() - psutil.boot_time()
        days = int(uptime_seconds // 86400)
        hours = int((uptime_seconds % 86400) // 3600)
        minutes = int((uptime_seconds % 3600) // 60)
        seconds = int(uptime_seconds % 60)
        return {"Uptime": f"{days}d {hours}h {minutes}m {seconds}s"}
    except Exception as e:
        return {"Error": f"Failed to get uptime: {str(e)}"}

# --------------------------
# Extended Benchmarks
# --------------------------

def extended_cpu_benchmark(duration=5):
    """Extended CPU benchmark: perform multiple math ops"""
    import math
    end = time.time() + duration
    count = 0
    while time.time() < end:
        for i in range(1, 100):
            _ = math.sqrt(i * i + 1) * math.sin(i) + math.log(i + 1)
        count += 1
    return {"Extended CPU Benchmark loops": count}

def extended_memory_benchmark():
    """Allocate large arrays and perform operations"""
    import numpy as np
    try:
        arr = np.random.rand(10**7)  # 10 million floats
        start = time.time()
        s = np.sum(arr)
        elapsed = time.time() - start
        return {"Extended Memory Benchmark (sum 10M floats)": f"{elapsed:.4f} seconds"}
    except Exception as e:
        return {"Error": f"Failed extended memory benchmark: {str(e)}"}

def extended_disk_benchmark():
    """Perform multiple write/read cycles on disk"""
    import tempfile
    import os
    try:
        filename = tempfile.mktemp()
        data = b"x" * (5 * 10**7)  # 50 MB
        start = time.time()
        for _ in range(3):
            with open(filename, "wb") as f:
                f.write(data)
            with open(filename, "rb") as f:
                _ = f.read()
        elapsed = time.time() - start
        os.remove(filename)
        return {"Extended Disk Benchmark (3x 50MB write/read)": f"{elapsed:.4f} seconds"}
    except Exception as e:
        return {"Error": f"Failed extended disk benchmark: {str(e)}"}

# --------------------------
# GUI Extended Features
# --------------------------

class ExtendedHardwareHouseApp(HardwareHouseApp):
    def __init__(self):
        super().__init__()
        # Add more categories
        self.category_options.extend([
            "USB Devices",
            "Display Monitors",
            "Printers",
            "Installed Software",
            "Power Plan",
            "Locale & Timezone",
            "System Uptime",
            "Extended Benchmarks",
        ])
        self.combo.configure(values=self.category_options)

    def show_info(self):
        cat = self.combo.get()
        self.textbox.delete("0.0", "end")
        info = {}

        # Check new categories first
        if cat == "USB Devices":
            info = {"USB Devices": get_usb_devices()}
        elif cat == "Display Monitors":
            info = {"Display Monitors": get_display_monitors()}
        elif cat == "Printers":
            info = {"Printers": get_printers()}
        elif cat == "Installed Software":
            info = {"Installed Software": get_installed_software()}
        elif cat == "Power Plan":
            info = get_power_plan()
        elif cat == "Locale & Timezone":
            info = get_system_locale()
        elif cat == "System Uptime":
            info = get_system_uptime()
        elif cat == "Extended Benchmarks":
            self.textbox.insert("end", "Running extended benchmarks...\n")
            def run_extended_benchmarks():
                results = {}
                results.update(extended_cpu_benchmark())
                results.update(extended_memory_benchmark())
                results.update(extended_disk_benchmark())
                self.current_info = results
                self.textbox.after(0, lambda: self.display_info(results))

            threading.Thread(target=run_extended_benchmarks).start()
            return
        else:
            # Fallback to parent categories
            super().show_info()
            return

        self.current_info = info
        self.display_info(info)

# --------------------------
# Main Entry Point
# --------------------------

if __name__ == "__main__":
    app = ExtendedHardwareHouseApp()
    app.mainloop()
