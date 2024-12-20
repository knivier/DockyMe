import os
import wmi
import tkinter as tk
import threading
import logging
from logging.handlers import RotatingFileHandler
from datetime import datetime
from queue import Queue
from typing import Dict, Any
import pythoncom

# Global variables
logger = None
root = None
text_box = None
event_queue = Queue()
running = True

class USBMonitor:
    def __init__(self):
        self.setup_logging()
        self.setup_gui()

    def setup_logging(self) -> None:
        """Initialize rotating log file handler"""
        global logger
        logger = logging.getLogger('USB_Monitor')
        log_file = 'usb_device_log.txt'
        handler = RotatingFileHandler(
            log_file,
            maxBytes=1024 * 1024,
            backupCount=5
        )
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        handler.setFormatter(formatter)
        logger.addHandler(handler)
        logger.setLevel(logging.INFO)

    def get_device_details(self, device) -> Dict[str, str]:
        """Extract device details from WMI device object"""
        device_details = {
            'Name': 'Unknown',
            'Vendor ID': 'Unknown',
            'Status': 'Unknown'
        }

        try:
            if hasattr(device, 'Caption'):
                device_details['Name'] = device.Caption
            elif hasattr(device, 'Description'):
                device_details['Name'] = device.Description

            if hasattr(device, 'DeviceID'):
                device_id = device.DeviceID
                if "VID_" in device_id:
                    device_details['Vendor ID'] = device_id.split("VID_")[1].split("&")[0]

            if hasattr(device, 'Status'):
                device_details['Status'] = device.Status

        except Exception as e:
            logger.error(f"Error getting device details: {str(e)}")

        return device_details

    def setup_gui(self) -> None:
        """Initialize and setup the GUI"""
        global root, text_box
        root = tk.Tk()
        root.title("USB Device Monitor")
        root.geometry("800x600")

        # Create main frame
        main_frame = tk.Frame(root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Add text widget with scrollbar
        text_box = tk.Text(main_frame, height=20, width=100)
        scrollbar = tk.Scrollbar(main_frame, orient=tk.VERTICAL, command=text_box.yview)
        text_box.configure(yscrollcommand=scrollbar.set)

        text_box.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Add buttons
        button_frame = tk.Frame(root)
        button_frame.pack(fill=tk.X, padx=10, pady=5)

        tk.Button(
            button_frame,
            text="Clear Log",
            command=lambda: text_box.delete(1.0, tk.END)
        ).pack(side=tk.LEFT, padx=5)

        tk.Button(
            button_frame,
            text="Exit",
            command=self.cleanup
        ).pack(side=tk.RIGHT, padx=5)

        root.protocol("WM_DELETE_WINDOW", self.cleanup)
        self.update_gui_from_queue()

    def update_gui_from_queue(self) -> None:
        """Process events from queue and update GUI"""
        try:
            while not event_queue.empty():
                message = event_queue.get_nowait()
                text_box.insert(tk.END, f"{message}\n")
                text_box.see(tk.END)
        except Exception as e:
            logger.error(f"Error updating GUI: {str(e)}")
        finally:
            if root and running:
                root.after(100, self.update_gui_from_queue)

    def monitor_devices(self):
        """Monitor USB devices in a separate thread"""
        # Initialize COM for this thread
        pythoncom.CoInitialize()
        try:
            # Create WMI connection in this thread
            c = wmi.WMI()
            
            # Monitor for device creation
            device_creation = c.Win32_USBHub.watch_for("creation")
            device_deletion = c.Win32_USBHub.watch_for("deletion")
            
            # First, get current devices
            current_devices = c.Win32_USBHub()
            event_queue.put("=== Current USB Devices ===")
            for device in current_devices:
                details = self.get_device_details(device)
                message = (
                    f"Device Found:\n"
                    f"Name: {details['Name']}\n"
                    f"Vendor ID: {details['Vendor ID']}\n"
                    f"Status: {details['Status']}\n"
                    f"{'-'*50}"
                )
                event_queue.put(message)

            # Then monitor for changes
            while running:
                try:
                    # Check for new devices
                    new_device = device_creation(timeout_ms=100)
                    if new_device:
                        details = self.get_device_details(new_device)
                        message = (
                            f"New Device Connected:\n"
                            f"Name: {details['Name']}\n"
                            f"Vendor ID: {details['Vendor ID']}\n"
                            f"Status: {details['Status']}\n"
                            f"{'-'*50}"
                        )
                        event_queue.put(message)
                        logger.info(f"New device connected: {details['Name']}")

                    # Check for removed devices
                    removed_device = device_deletion(timeout_ms=100)
                    if removed_device:
                        details = self.get_device_details(removed_device)
                        message = (
                            f"Device Disconnected:\n"
                            f"Name: {details['Name']}\n"
                            f"Vendor ID: {details['Vendor ID']}\n"
                            f"{'-'*50}"
                        )
                        event_queue.put(message)
                        logger.info(f"Device disconnected: {details['Name']}")

                except wmi.x_wmi_timed_out:
                    continue
                except Exception as e:
                    logger.error(f"Error in device monitoring: {str(e)}")
                    event_queue.put(f"Error monitoring devices: {str(e)}")

        except Exception as e:
            logger.error(f"Error in monitor_devices: {str(e)}")
            event_queue.put(f"Error initializing device monitoring: {str(e)}")
        finally:
            pythoncom.CoUninitialize()

    def cleanup(self) -> None:
        """Cleanup resources before exit"""
        global running
        running = False
        logger.info("Shutting down USB Monitor")
        root.destroy()

    def run(self) -> None:
        """Start the application"""
        # Start device monitoring in a separate thread
        monitor_thread = threading.Thread(target=self.monitor_devices, daemon=True)
        monitor_thread.start()

        # Start GUI main loop
        root.mainloop()

if __name__ == "__main__":
    try:
        monitor = USBMonitor()
        monitor.run()
    except Exception as e:
        if logger:
            logger.critical(f"Critical error in main: {str(e)}")
        raise