#!/usr/bin/env python3

import os
import getpass
import concurrent.futures
from exchangelib import Credentials, Account, Configuration, DELEGATE, EWSDateTime, EWSTimeZone
from exchangelib.errors import ErrorItemNotFound, ErrorServerBusy, ErrorMailboxStoreUnavailable
# Import timezone conversion dictionaries
from exchangelib.winzone import MS_TIMEZONE_TO_IANA_MAP, CLDR_TO_MS_TIMEZONE_MAP
import time
import sys
import random
import tkinter as tk
from tkinter import ttk
from datetime import datetime
import threading
import statistics
from colorama import init, Fore, Back, Style
import multiprocessing
import queue
import subprocess
import shutil
import select
import platform
import functools
import inspect

# Add a custom entry for the "Customized Time Zone"
# Use Europe/Paris as a reasonable value (adjust according to your needs)
MS_TIMEZONE_TO_IANA_MAP['Customized Time Zone'] = "Europe/Paris"
print(f"{Fore.CYAN}Custom timezone added: 'Customized Time Zone' -> 'Europe/Paris'{Style.RESET_ALL}")

# Platform detection
IS_WINDOWS = False  # Force Linux mode

# Install rich if necessary
try:
    from rich.console import Console
    from rich.table import Table
    from rich.panel import Panel
    from rich.layout import Layout
    from rich.live import Live
    RICH_AVAILABLE = True
except ImportError:
    RICH_AVAILABLE = False
    print("The 'rich' library is not installed. Attempting installation...")
    try:
        import subprocess
        subprocess.check_call([sys.executable, "-m", "pip", "install", "rich"])
        print("Installation successful! Importing rich...")
        from rich.console import Console
        from rich.table import Table
        from rich.panel import Panel
        from rich.layout import Layout
        from rich.live import Live
        RICH_AVAILABLE = True
    except:
        print("Failed to install 'rich'. Statistics will be displayed in simple mode.")

# Initialize colorama
init()

# Logo ASCII art - original version with Unicode characters
ews_logo = f"""{Fore.CYAN}╔══════════════════════════════════════════════════════════════════════════╗
{Fore.CYAN}║ {Fore.GREEN}███████╗██╗    ██╗███████╗     ██████╗██╗     ███████╗ █████╗ ███╗   ██╗{Fore.CYAN} ║
{Fore.CYAN}║ {Fore.GREEN}██╔════╝██║    ██║██╔════╝    ██╔════╝██║     ██╔════╝██╔══██╗████╗  ██║{Fore.CYAN} ║
{Fore.CYAN}║ {Fore.GREEN}█████╗  ██║ █╗ ██║███████╗    ██║     ██║     █████╗  ███████║██╔██╗ ██║{Fore.CYAN} ║
{Fore.CYAN}║ {Fore.GREEN}██╔══╝  ██║███╗██║╚════██║    ██║     ██║     ██╔══╝  ██╔══██║██║╚██╗██║{Fore.CYAN} ║
{Fore.CYAN}║ {Fore.GREEN}███████╗╚███╔███╔╝███████║    ╚██████╗███████╗███████╗██║  ██║██║ ╚████║{Fore.CYAN} ║
{Fore.CYAN}║ {Fore.GREEN}╚══════╝ ╚══╝╚══╝ ╚══════╝     ╚═════╝╚══════╝╚══════╝╚═╝  ╚═╝╚═╝  ╚═══╝{Fore.CYAN} ║
{Fore.CYAN}║                                                                              ║
{Fore.CYAN}║                      {Fore.YELLOW}[ Exchange Web Services Cleaner ]{Fore.CYAN}                     ║
{Fore.CYAN}╚══════════════════════════════════════════════════════════════════════════╝{Style.RESET_ALL}"""

# Function to print the logo
def print_logo():
    """Display the program logo"""
    print(ews_logo)

# Function to get credentials
def get_credentials():
    """Get username and password interactively"""
    username = input(f"{Fore.GREEN}Enter your email address: {Style.RESET_ALL}")
    password = getpass.getpass(f"{Fore.GREEN}Enter your password: {Style.RESET_ALL}")
    return username, password

# Class to track EWS call times
class EWSStats:
    def __init__(self):
        self.call_times = []
        self.call_types = {}  # Dictionary to store times by call type
        self.lock = threading.Lock()
        self.last_call_time = 0
        self.max_calls_to_keep = 1000  # Limit the number of entries to avoid excessive memory usage
        self.last_commands = {}  # Store the last commands by type
        self.active_calls = 0  # Count of active calls
    
    def add_call_time(self, ms, call_type="generic", command_details=""):
        with self.lock:
            self.last_call_time = ms
            self.call_times.append(ms)
            
            # Limit the size of lists
            if len(self.call_times) > self.max_calls_to_keep:
                self.call_times = self.call_times[-self.max_calls_to_keep:]
            
            # Add by type
            if call_type not in self.call_types:
                self.call_types[call_type] = []
            
            self.call_types[call_type].append(ms)
            
            # Store command details
            self.last_commands[call_type] = command_details
            
            # Limit the size of lists by type
            if len(self.call_types[call_type]) > self.max_calls_to_keep:
                self.call_types[call_type] = self.call_types[call_type][-self.max_calls_to_keep:]
            
            # Add a log for this call
            log_level = "INFO"
            
            # If it's an error, always log it as an error
            if call_type.startswith("error_"):
                log_level = "ERROR"
                # Force log addition to the unified interface for errors
                if ews_unified_interface and ews_unified_interface.running:
                    ews_unified_interface.add_log(f"{command_details} - {ms:.2f}ms", log_level)
            
            # Add a special log for slow calls (more than 1000ms)
            if ms > 1000:
                # Add an emoji to make it more visible
                slow_log_message = f"🕒 SLOW EWS CALL: {call_type} - {ms:.2f}ms - {command_details}"
                ews_logger.add_log(slow_log_message, "WARN")
                # Also add to the unified interface with priority
                if ews_unified_interface and ews_unified_interface.running:
                    ews_unified_interface.add_log(slow_log_message, "WARN")
            
            # Standard log for all calls
            log_message = f"EWS call: {command_details} - {call_type} - {ms:.2f}ms"
            ews_logger.add_log(log_message, log_level)
    
    def start_call(self):
        with self.lock:
            self.active_calls += 1
    
    def end_call(self):
        with self.lock:
            if self.active_calls > 0:
                self.active_calls -= 1
    
    def get_active_calls(self):
        with self.lock:
            return self.active_calls
    
    def get_last_command(self, call_type):
        with self.lock:
            return self.last_commands.get(call_type, "")
    
    def get_stats(self):
        with self.lock:
            if not self.call_times:
                return {"last": 0, "min": 0, "max": 0, "avg": 0, "median": 0, "count": 0, "active": self.active_calls}
            
            return {
                "last": self.last_call_time,
                "min": min(self.call_times),
                "max": max(self.call_times),
                "avg": sum(self.call_times) / len(self.call_times),
                "median": statistics.median(self.call_times) if len(self.call_times) > 0 else 0,
                "count": len(self.call_times),
                "active": self.active_calls
            }
    
    def get_type_stats(self, call_type):
        with self.lock:
            if call_type not in self.call_types or not self.call_types[call_type]:
                return {"min": 0, "max": 0, "avg": 0, "count": 0}
            
            call_list = self.call_types[call_type]
            return {
                "min": min(call_list),
                "max": max(call_list),
                "avg": sum(call_list) / len(call_list),
                "count": len(call_list),
                "last_command": self.last_commands.get(call_type, "")
            }

# Create a global instance for statistics
ews_stats = EWSStats()

# Class for EWS logs
class EWSLogger:
    def __init__(self, log_file=None):
        self.log_entries = []
        self.log_queue = queue.Queue()
        self.log_thread = None
        self.running = False
        self.log_to_console = False
        
        # Set the log file path
        if log_file:
            self.log_file = log_file
        else:
            # Default log file in /tmp on Linux or temp on Windows
            if platform.system() == "Windows":
                self.log_file = os.path.join(os.environ.get("TEMP", "C:\\Temp"), "ews_logs.txt")
            else:
                self.log_file = "/tmp/ews_logs.txt"
        
        # Ensure the log file is accessible
        self.setup_log_file()
    
    def setup_log_file(self):
        """Ensure the log file is created and accessible"""
        try:
            # Get the directory of the log file
            log_dir = os.path.dirname(self.log_file)
            
            # Check if the directory exists and create it if needed
            if not os.path.exists(log_dir) and log_dir:
                try:
                    os.makedirs(log_dir, exist_ok=True)
                    print(f"{Fore.GREEN}Created log directory: {log_dir}{Style.RESET_ALL}")
                except Exception as e:
                    print(f"{Fore.RED}Failed to create log directory: {log_dir} - {e}{Style.RESET_ALL}")
                    # Fall back to using home directory
                    home_dir = os.path.expanduser("~")
                    self.log_file = os.path.join(home_dir, "ews_logs.txt")
                    print(f"{Fore.YELLOW}Using alternate log location: {self.log_file}{Style.RESET_ALL}")
            
            # Create the log file if it doesn't exist
            with open(self.log_file, 'a') as f:
                if os.path.getsize(self.log_file) == 0:
                    # File was just created, write header
                    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    f.write(f"=== EWS CLEANER LOG STARTED AT {timestamp} ===\n")
            
            # Verify the file exists and is writeable
            if not os.path.exists(self.log_file):
                raise IOError(f"Log file could not be created: {self.log_file}")
            
            print(f"{Fore.GREEN}Log file initialized: {self.log_file}{Style.RESET_ALL}")
            print(f"{Fore.CYAN}To view logs in real-time, use: tail -f {self.log_file}{Style.RESET_ALL}")
        except Exception as e:
            print(f"{Fore.RED}Failed to initialize log file: {e}{Style.RESET_ALL}")
            # Continue without logging to file
            self.log_file = None
    
    def start_logging(self):
        """Start the logging thread"""
        pass
    
    def add_log(self, message, level="INFO"):
        """Add a message to the log"""
        timestamp = datetime.now().strftime("%H:%M:%S.%f")[:-3]
        log_entry = {
            "timestamp": timestamp,
            "message": message,
            "level": level
        }
        # Log to console if enabled
        if self.log_to_console:
            color = Fore.GREEN
            if level == "ERROR":
                color = Fore.RED
            elif level == "WARN":
                color = Fore.YELLOW
            print(f"{color}[{timestamp}] {message}{Style.RESET_ALL}")
        
        # Store log in memory
        self.log_entries.append(log_entry)
        if len(self.log_entries) > 1000:  # limit to 1000 entries
            self.log_entries = self.log_entries[-1000:]
            
        # Log to file
        if self.log_file:
            try:
                with open(self.log_file, 'a') as f:
                    f.write(f"[{timestamp}] [{level}] {message}\n")
            except Exception as e:
                print(f"{Fore.RED}Failed to write to log file: {e}{Style.RESET_ALL}")
    
    def show_log_window(self):
        """Display log window (dummy implementation for Linux)"""
        print(f"{Fore.GREEN}Log window: using file {self.log_file}{Style.RESET_ALL}")
    
    def stop(self):
        """Stop logging"""
        self.running = False

# Class to display EWS statistics in a separate window
class EWSStatsWindow:
    def __init__(self):
        self.running = False
        self.user_exit_requested = False
    
    def show_stats_window(self):
        """Display statistics window"""
        self.running = True
        print(f"{Fore.GREEN}Statistics window opened{Style.RESET_ALL}")
    
    def exit_stats_window(self):
        """Close statistics window"""
        self.running = False
        self.user_exit_requested = True
        print(f"{Fore.YELLOW}Statistics window closed{Style.RESET_ALL}")
    
    def stop(self):
        """Stop statistics window"""
        self.running = False
        print(f"{Fore.YELLOW}Statistics window stopped{Style.RESET_ALL}")

# Modify the EWSUnifiedInterface class to control when monitoring starts
class EWSUnifiedInterface:
    def __init__(self):
        self.log_queue = multiprocessing.Queue()
        self.command_queue = multiprocessing.Queue()
        self.data_queue = multiprocessing.Queue()
        self.interface_process = None
        self.running = False
        self.update_thread = None
        self.user_exit_requested = False
        self.monitoring_active = False  # New flag to control monitoring activation
        # Processing progress data
        self.progress_data = {
            "folder_name": "",
            "processed": 0,
            "remaining": 0,
            "speed": 0,
            "est_time": 0,
            "active": False
        }
    
    def add_log(self, message, level="INFO"):
        """Add a message to the log"""
        timestamp = datetime.now().strftime("%H:%M:%S.%f")[:-3]
        log_entry = {
            "timestamp": timestamp,
            "message": message,
            "level": level
        }
        if self.running:
            try:
                self.log_queue.put(log_entry)
            except:
                pass
    
    def add_folders(self, folders_data):
        """Add available folders"""
        if self.running:
            try:
                self.data_queue.put({"type": "folders", "data": folders_data})
            except:
                pass
    
    def get_command(self, timeout=0.1):
        """Get a command from the user interface"""
        try:
            return self.command_queue.get(timeout=timeout)
        except queue.Empty:
            return None
    
    def start(self, start_monitoring=False):
        """Start the unified interface"""
        self.running = True
        self.monitoring_active = start_monitoring
        print(f"{Fore.GREEN}Unified interface started (monitoring: {start_monitoring}){Style.RESET_ALL}")
    
    def start_monitoring(self):
        """Activate EWS call monitoring"""
        if not self.monitoring_active and self.running:
            self.monitoring_active = True
            stop_event = multiprocessing.Event()
            
            self.interface_process = multiprocessing.Process(
                target=interface_process,
                args=(self.command_queue, self.data_queue, self.log_queue, stop_event)
            )
            self.interface_process.daemon = True
            self.interface_process.start()
            
            # Start the statistics update thread
            self.update_thread = threading.Thread(target=self.update_stats)
            self.update_thread.daemon = True
            self.update_thread.start()
            
            print(f"{Fore.GREEN}EWS monitoring activated{Style.RESET_ALL}")
            return True
        return False
    
    def update_stats(self):
        """Periodically update statistics"""
        last_stats = {}
        
        while self.running and self.monitoring_active:
            try:
                # Get current statistics
                stats = ews_stats.get_stats()
                
                # Only update if stats have changed
                if stats != last_stats:
                    last_stats = stats.copy()
                    
                    # Prepare call type data
                    call_types = {}
                    for call_type in ews_stats.call_types.keys():
                        call_types[call_type] = ews_stats.get_type_stats(call_type)
                    
                    # Send stats to the interface process
                    stats_data = {
                        "type": "stats",
                        "data": {
                            "stats": stats,
                            "call_types": call_types,
                            "timestamp": datetime.now().strftime("%H:%M:%S")
                        }
                    }
                    
                    if self.running and self.monitoring_active:
                        self.data_queue.put(stats_data)
            
            except Exception as e:
                print(f"Error updating stats: {e}")
            
            # Wait before the next update
            time.sleep(0.5)
    
    def stop(self):
        """Properly stop the interface"""
        self.running = False
        self.monitoring_active = False
        
        if self.update_thread and self.update_thread.is_alive():
            self.update_thread.join(timeout=2.0)
        
        if self.interface_process and self.interface_process.is_alive():
            try:
                self.interface_process.terminate()
            except:
                pass
        
        print(f"{Fore.YELLOW}Unified interface stopped{Style.RESET_ALL}")

    def update_progress(self, folder_name, processed, remaining, speed, est_time):
        """Update progress data"""
        self.progress_data = {
            "folder_name": folder_name,
            "processed": processed,
            "remaining": remaining,
            "speed": speed,
            "est_time": est_time,
            "active": True
        }
        
        # Send progress data to the interface process
        if self.running and self.monitoring_active:
            try:
                self.data_queue.put({
                    "type": "progress",
                    "data": self.progress_data
                })
            except:
                pass
    
    def reset_progress(self):
        """Reset progress data"""
        self.progress_data = {
            "folder_name": "",
            "processed": 0,
            "remaining": 0,
            "speed": 0,
            "est_time": 0,
            "active": False
        }
        
        # Send reset progress data
        if self.running and self.monitoring_active:
            try:
                self.data_queue.put({
                    "type": "progress",
                    "data": self.progress_data
                })
            except:
                pass

# Function to intercept and time EWS calls
def intercept_ews_calls():
    """Intercepts EWS calls to measure their execution time"""
    from exchangelib.protocol import Protocol
    
    # Try to intercept the post method (which is typically the one that makes HTTP requests)
    if hasattr(Protocol, 'post'):
        # Save the original method
        original_post = Protocol.post
        
        @functools.wraps(original_post)
        def wrapped_post(self, *args, **kwargs):
            # Measure execution time
            ews_stats.start_call()
            start_time = time.time()
            
            # Try to determine the call type
            call_type = "unknown"
            call_info = "EWS request"
            
            try:
                if len(args) > 0 and hasattr(args[0], 'tag'):
                    call_type = args[0].tag.localname
                    call_info = f"Operation: {call_type}"
                elif 'data' in kwargs and hasattr(kwargs['data'], 'tag'):
                    call_type = kwargs['data'].tag.localname
                    call_info = f"Operation: {call_type}"
            except:
                pass
            
            try:
                response = original_post(self, *args, **kwargs)
                elapsed_ms = (time.time() - start_time) * 1000
                ews_stats.add_call_time(elapsed_ms, call_type, call_info)
                ews_stats.end_call()
                return response
            except Exception as e:
                elapsed_ms = (time.time() - start_time) * 1000
                error_message = f"EWS Error ({call_type}): {str(e)}"
                
                # For slow but successful calls
                if elapsed_ms > 1000:
                    pause_duration = 3  # shorter pause for successful calls
                    current_time = datetime.now().strftime("%H:%M:%S")
                    
                    # Pause message with timestamp
                    pause_message = f"⚠️ BEGINNING PAUSE OF {pause_duration}s ({current_time}): Slow but successful operation - {elapsed_ms:.2f}ms"
                    print(f"{Fore.YELLOW}{pause_message}{Style.RESET_ALL}")
                    ews_logger.add_log(pause_message, "WARN")
                    if ews_unified_interface and ews_unified_interface.running:
                        ews_unified_interface.add_log(pause_message, "WARN")
                    
                    # Execute the pause with verification
                    pause_start = time.time()
                    time.sleep(pause_duration)
                    actual_pause = time.time() - pause_start
                    
                    # Message after the pause
                    current_time = datetime.now().strftime("%H:%M:%S")
                    end_pause_message = f"✅ END PAUSE ({current_time}): Actual duration = {actual_pause:.1f}s"
                    print(f"{Fore.GREEN}{end_pause_message}{Style.RESET_ALL}")
                    ews_logger.add_log(end_pause_message, "INFO")
                    if ews_unified_interface and ews_unified_interface.running:
                        ews_unified_interface.add_log(end_pause_message, "INFO")
                
                # For database errors
                if "mailbox database is temporarily unavailable" in str(e).lower():
                    pause_duration = 30  # 30 second pause
                    current_time = datetime.now().strftime("%H:%M:%S")
                    
                    # Message BEFORE the pause, with timestamp
                    pause_message = f"⚠️⚠️ BEGINNING PAUSE OF {pause_duration}s ({current_time}): Database unavailable - {elapsed_ms:.2f}ms"
                    print(f"{Fore.RED}{pause_message}{Style.RESET_ALL}")
                    # Add the pause message to logs and interface with high priority
                    ews_logger.add_log(pause_message, "ERROR")
                    ews_unified_interface.add_log(pause_message, "ERROR")
                    
                    # Execute the pause with verification
                    pause_start = time.time()
                    time.sleep(pause_duration)
                    actual_pause = time.time() - pause_start
                    
                    # Message AFTER the pause, with actual duration
                    current_time = datetime.now().strftime("%H:%M:%S")
                    end_pause_message = f"✅ END PAUSE ({current_time}): Actual duration = {actual_pause:.1f}s"
                    print(f"{Fore.GREEN}{end_pause_message}{Style.RESET_ALL}")
                    ews_logger.add_log(end_pause_message, "INFO")
                    ews_unified_interface.add_log(end_pause_message, "INFO")
                
                ews_stats.add_call_time(elapsed_ms, f"error_{call_type}", f"Error in {call_info}: {str(e)}")
                ews_logger.add_log(error_message, "ERROR")
                ews_stats.end_call()
                raise
        
        # Replace the original method with our wrapper
        Protocol.post = wrapped_post
        print(f"{Fore.GREEN}EWS monitoring enabled: Protocol.post method intercepted{Style.RESET_ALL}")
        return True
    
    # If post doesn't exist, try with send which is also commonly used
    elif hasattr(Protocol, 'send'):
        # Save the original method
        original_send = Protocol.send
        
        @functools.wraps(original_send)
        def wrapped_send(self, *args, **kwargs):
            # Measure execution time
            ews_stats.start_call()
            start_time = time.time()
            
            # Try to determine the call type
            call_type = "unknown"
            call_info = "EWS request"
            
            try:
                if len(args) > 0 and hasattr(args[0], 'tag'):
                    call_type = args[0].tag.localname
                    call_info = f"Operation: {call_type}"
            except:
                pass
            
            try:
                response = original_send(self, *args, **kwargs)
                elapsed_ms = (time.time() - start_time) * 1000
                ews_stats.add_call_time(elapsed_ms, call_type, call_info)
                ews_stats.end_call()
                return response
            except Exception as e:
                elapsed_ms = (time.time() - start_time) * 1000
                error_message = f"EWS Error ({call_type}): {str(e)}"
                ews_stats.add_call_time(elapsed_ms, f"error_{call_type}", f"Error in {call_info}: {str(e)}")
                ews_logger.add_log(error_message, "ERROR")
                ews_stats.end_call()
                raise
        
        # Replace the original method with our wrapper
        Protocol.send = wrapped_send
        print(f"{Fore.GREEN}EWS monitoring enabled: Protocol.send method intercepted{Style.RESET_ALL}")
        return True
    
    # If no standard method is found, try another approach - intercept calls at the Account level
    from exchangelib.items import Item
    
    if hasattr(Item, 'delete'):
        original_delete = Item.delete
        
        @functools.wraps(original_delete)
        def wrapped_delete(self, *args, **kwargs):
            # Measure execution time
            ews_stats.start_call()
            start_time = time.time()
            call_type = "delete"
            call_info = f"Delete item: {self.__class__.__name__}"
            
            try:
                response = original_delete(self, *args, **kwargs)
                elapsed_ms = (time.time() - start_time) * 1000
                ews_stats.add_call_time(elapsed_ms, call_type, call_info)
                ews_stats.end_call()
                
                # Add a pause for successful but slow calls
                if elapsed_ms > 1000:
                    pause_duration = 3  # shorter pause for successful calls
                    time_before = datetime.now().strftime("%H:%M:%S.%f")[:-3]
                    pause_message = f"⏱️ BEGINNING PAUSE {time_before} - Slow but successful deletion ({elapsed_ms:.0f}ms) - waiting {pause_duration}s"
                    print(f"{Fore.YELLOW}{pause_message}{Style.RESET_ALL}")
                    ews_logger.add_log(pause_message, "WARN")
                    ews_unified_interface.add_log(pause_message, "WARN")
                    
                    time.sleep(pause_duration)
                    
                    time_after = datetime.now().strftime("%H:%M:%S.%f")[:-3]
                    after_message = f"✅ END PAUSE {time_after} - Resuming after {pause_duration}s wait"
                    print(f"{Fore.GREEN}{after_message}{Style.RESET_ALL}")
                    ews_logger.add_log(after_message, "INFO")
                    ews_unified_interface.add_log(after_message, "INFO")
                
                return response
            except Exception as e:
                elapsed_ms = (time.time() - start_time) * 1000
                error_message = f"Error while deleting an item: {str(e)}"
                error_type = "error_delete"
                
                # Enhanced log for high-latency errors
                if elapsed_ms > 1000:
                    error_type = "error_delete_slow"
                    ews_logger.add_log(f"SLOW DELETE ERROR: {error_message} - {elapsed_ms:.2f}ms", "ERROR")
                    ews_unified_interface.add_log(f"Slow deletion error detected: {elapsed_ms:.2f}ms", "ERROR")
                    
                    # Forced pause for slow deletions
                    if "mailbox database is temporarily unavailable" in str(e).lower():
                        pause_duration = 30  # 30 second pause
                        time_before = datetime.now().strftime("%H:%M:%S.%f")[:-3]
                        pause_message = f"⛔ BEGINNING PAUSE {time_before} - Database unavailable ({elapsed_ms:.0f}ms) - waiting {pause_duration}s"
                        print(f"{Fore.RED}{pause_message}{Style.RESET_ALL}")
                        ews_logger.add_log(pause_message, "ERROR")
                        ews_unified_interface.add_log(pause_message, "ERROR")
                        
                        time.sleep(pause_duration)
                        
                        time_after = datetime.now().strftime("%H:%M:%S.%f")[:-3]
                        after_message = f"✅ END PAUSE {time_after} - Resuming after {pause_duration}s wait"
                        print(f"{Fore.GREEN}{after_message}{Style.RESET_ALL}")
                        ews_logger.add_log(after_message, "INFO")
                        ews_unified_interface.add_log(after_message, "INFO")
                    else:
                        # Other slow errors
                        pause_duration = 5  # 5 second pause
                        pause_message = f"⚠️ PAUSE OF {pause_duration}s: Other slow deletion error - {elapsed_ms:.2f}ms"
                        print(f"{Fore.YELLOW}{pause_message}{Style.RESET_ALL}")
                        ews_logger.add_log(pause_message, "WARN")
                        ews_unified_interface.add_log(pause_message, "WARN")
                        time.sleep(pause_duration)
                
                ews_stats.add_call_time(elapsed_ms, error_type, f"Error in {call_info}: {str(e)}")
                ews_logger.add_log(error_message, "ERROR")
                ews_stats.end_call()
                raise
        
        # Replace the original method with our wrapper
        Item.delete = wrapped_delete
        print(f"{Fore.GREEN}EWS monitoring enabled: Item.delete method intercepted{Style.RESET_ALL}")
        return True
    
    # If no method could be intercepted
    print(f"{Fore.RED}Unable to enable EWS monitoring: no known method was found in exchangelib{Style.RESET_ALL}")
    return False

# Function to list folders
def list_folders(account):
    """List all folders in the account"""
    print(f"\n{Fore.CYAN}Listing folders...{Style.RESET_ALL}")
    folders = []
    
    try:
        # Traverse all folders recursively
        for i, folder in enumerate(account.root.walk(), 1):
            try:
                # Add information about the folder
                print(f"{Fore.WHITE}{i}. {Fore.GREEN}{folder.name} {Fore.CYAN}({folder.total_count} items){Style.RESET_ALL}")
                folders.append(folder)
                
                # Record in logs
                ews_logger.add_log(f"Found folder: {folder.name} ({folder.total_count} items)")
                
                # We don't need to manually add time measurements here
                # because the interceptor already measures them
            except Exception as e:
                print(f"{Fore.RED}Error getting folder info: {e}{Style.RESET_ALL}")
                ews_logger.add_log(f"Error getting folder info: {e}", "ERROR")
    except Exception as e:
        print(f"{Fore.RED}Error listing folders: {e}{Style.RESET_ALL}")
        ews_logger.add_log(f"Error listing folders: {e}", "ERROR")
    
    return folders

# Function to empty a folder
def empty_folder(folder, batch_size=100):
    """Empty a folder by deleting all items in batches"""
    print(f"\n{Fore.YELLOW}Processing folder: {Fore.GREEN}{folder.name}{Style.RESET_ALL}")
    print(f"{Fore.CYAN}Total items: {folder.total_count}{Style.RESET_ALL}")
    
    if folder.total_count == 0:
        print(f"{Fore.GREEN}Folder is already empty.{Style.RESET_ALL}")
        return
    
    # Start measuring time
    start_time = time.time()
    processed = 0
    
    try:
        # Get the folder ID to be able to find it again later
        folder_id = folder.id
        folder_name = folder.name
        folder_path = folder.absolute
        account = folder.account
        
        # Process in batches
        while True:
            # Get updated folder from the server
            try:
                # ... existing code to refresh the folder ...
                if hasattr(folder, 'refresh'):
                    folder.refresh()
                    actual_remaining = folder.total_count
                else:
                    # ... existing code ...
                    if hasattr(account.root, 'get_folder_by_path'):
                        try:
                            folder = account.root.get_folder_by_path(folder_path)
                            actual_remaining = folder.total_count
                        except:
                            actual_remaining = folder.total_count
                    else:
                        actual_remaining = folder.total_count
                
                if actual_remaining == 0:
                    print(f"{Fore.GREEN}Folder is now empty.{Style.RESET_ALL}")
                    # Reset progress data
                    ews_unified_interface.reset_progress()
                    break
            except Exception as e:
                print(f"{Fore.RED}Error refreshing folder info: {e}{Style.RESET_ALL}")
                ews_logger.add_log(f"Error refreshing folder info: {e}", "ERROR")
                actual_remaining = folder.total_count
                if actual_remaining == 0:
                    ews_unified_interface.reset_progress()
                    break
            
            # Get a batch of items
            # We don't need to measure manually because the interceptor does it
            try:
                items = list(folder.all().order_by('datetime_received')[:batch_size])
            except Exception as e:
                error_message = str(e)
                print(f"{Fore.RED}Error getting items: {error_message}{Style.RESET_ALL}")
                
                # Check if it's the specific error for database temporarily unavailable
                if "mailbox database is temporarily unavailable" in error_message.lower():
                    wait_time = 30  # 30 second wait for this specific error
                    error_log = f"Mailbox database temporarily unavailable. Waiting {wait_time} seconds."
                    print(f"{Fore.YELLOW}{error_log}{Style.RESET_ALL}")
                    ews_logger.add_log(error_log, "WARN")
                    time.sleep(wait_time)
                else:
                    # For other errors, wait only 2 seconds
                    ews_logger.add_log(f"Error getting items: {error_message}", "ERROR")
                    time.sleep(2)  # Pause in case of error
                
                continue
            
            if not items:
                print(f"{Fore.YELLOW}No more items found in folder.{Style.RESET_ALL}")
                ews_unified_interface.reset_progress()
                break
            
            # Delete the batch
            print(f"{Fore.CYAN}Deleting batch of {len(items)} items...{Style.RESET_ALL}")
            ews_logger.add_log(f"Deleting batch of {len(items)} items from {folder.name}")
            
            # Add tracking of consecutive failures
            consecutive_failures = 0
            
            try:
                # Delete items one by one - the interceptor will measure each call
                for item in items:
                    try:
                        # Add a delay between each call to reduce load
                        time.sleep(0.05)  # 50ms between each call
                        
                        item.delete()  # The interceptor automatically measures this call
                        consecutive_failures = 0  # Reset counter on success
                    except Exception as delete_error:
                        error_message = str(delete_error)
                        
                        # Check if it's the specific error for database temporarily unavailable
                        # but only if it hasn't already been handled by wrapped_delete
                        if "mailbox database is temporarily unavailable" in error_message.lower() and "⚠️ PAUSE" not in error_message:
                            consecutive_failures += 1
                            
                            # If we have multiple consecutive failures, take a longer pause
                            if consecutive_failures >= 3:
                                pause_duration = 60  # 1 minute pause
                                pause_message = f"⚠️ PAUSE OF {pause_duration}s: Database temporarily unavailable (multiple failures)"
                                print(f"{Fore.RED}{pause_message}{Style.RESET_ALL}")
                                ews_logger.add_log(pause_message, "ERROR")
                                
                                # Add a log in the unified interface
                                ews_unified_interface.add_log(pause_message, "ERROR")
                                
                                # Pause
                                time.sleep(pause_duration)
                                consecutive_failures = 0  # Reset after the pause
                            else:
                                # Shorter pause for first errors
                                short_pause = 5  # 5 seconds
                                print(f"{Fore.YELLOW}Database temporarily unavailable. Pause of {short_pause} seconds.{Style.RESET_ALL}")
                                ews_logger.add_log(f"Database temporarily unavailable. Pause of {short_pause} seconds.", "WARN")
                                time.sleep(short_pause)
                        else:
                            error_message = f"EWS deletion error: {str(delete_error)}"
                            print(f"{Fore.RED}{error_message}{Style.RESET_ALL}")
                            ews_logger.add_log(error_message, "ERROR")
                            # Continue with other items
                
                processed += len(items)
                
                # Update the counter without trying to refresh the folder
                try:
                    actual_remaining = folder.total_count 
                except Exception as e:
                    print(f"{Fore.YELLOW}Could not get folder count: {e}{Style.RESET_ALL}")
                    actual_remaining = max(0, folder.total_count - len(items))
                
                # Display progress
                elapsed = time.time() - start_time
                items_per_sec = processed / elapsed if elapsed > 0 else 0
                est_time = actual_remaining / items_per_sec if items_per_sec > 0 else 0
                
                # Update progress data in the interface
                ews_unified_interface.update_progress(
                    folder_name=folder.name,
                    processed=processed,
                    remaining=actual_remaining,
                    speed=items_per_sec,
                    est_time=est_time
                )
                
                formatted_time = format_time_remaining(est_time)
                print(f"{Fore.GREEN}Processed: {processed} items, {Fore.YELLOW}Remaining: {actual_remaining}, {Fore.CYAN}Speed: {items_per_sec:.2f} items/sec, {Fore.MAGENTA}Est. time left: {formatted_time}{Style.RESET_ALL}")
                
                # Add a delay to avoid too much load on the server
                time.sleep(0.1)
            
            except ErrorServerBusy as e:
                back_off = e.back_off
                print(f"{Fore.YELLOW}Server busy. Waiting for {back_off} seconds...{Style.RESET_ALL}")
                ews_logger.add_log(f"Server busy. Waiting for {back_off} seconds", "WARN")
                time.sleep(back_off)
            
            except ErrorMailboxStoreUnavailable:
                print(f"{Fore.YELLOW}Mailbox store unavailable. Waiting...{Style.RESET_ALL}")
                ews_logger.add_log("Mailbox store unavailable. Waiting...", "WARN")
                time.sleep(5)
            
            except Exception as e:
                print(f"{Fore.RED}Error deleting items: {e}{Style.RESET_ALL}")
                ews_logger.add_log(f"Error deleting items: {e}", "ERROR")
                time.sleep(2)  # Pause in case of error
        
        # Display summary
        elapsed = time.time() - start_time
        print(f"\n{Fore.GREEN}Folder processing complete!{Style.RESET_ALL}")
        print(f"{Fore.CYAN}Processed {processed} items in {elapsed:.2f} seconds ({processed/elapsed:.2f} items/sec){Style.RESET_ALL}")
        ews_logger.add_log(f"Completed processing folder {folder.name}. Deleted {processed} items in {elapsed:.2f} seconds.")
        
        # Reset progress data after completion
        ews_unified_interface.reset_progress()
    
    except Exception as e:
        print(f"{Fore.RED}Error processing folder: {e}{Style.RESET_ALL}")
        ews_logger.add_log(f"Error processing folder: {e}", "ERROR")
        # Reset progress data in case of error
        ews_unified_interface.reset_progress()

# Create global instances for logs and statistics
ews_logger = EWSLogger()
ews_stats_window = EWSStatsWindow()

# Create a global instance of the unified interface
ews_unified_interface = EWSUnifiedInterface()

# Function to format time in days, hours, minutes, seconds
def format_time_remaining(seconds):
    """Converts time in seconds to format days, hours, minutes, seconds"""
    if seconds <= 0:
        return "0s"
    
    days, remainder = divmod(int(seconds), 86400)
    hours, remainder = divmod(remainder, 3600)
    minutes, seconds = divmod(remainder, 60)
    
    result = ""
    if days > 0:
        result += f"{days}d "
    if hours > 0 or days > 0:
        result += f"{hours}h "
    if minutes > 0 or hours > 0 or days > 0:
        result += f"{minutes}m "
    result += f"{seconds}s"
    
    return result

# Function to handle the interface in a separate process
def interface_process(command_queue, data_queue, log_queue, stop_event):
    """Separate process that handles the user interface with a simple console display"""
    try:
        # Interface state
        logs = []
        stats_data = None
        ews_calls = []  # To store recent EWS calls
        max_calls_to_display = 20  # Maximum number of calls to display
        progress_data = None  # To store current progress information
        progress_history = []  # To store progress history
        last_progress_time = 0  # To limit frequency of adding to history
        
        print("\033[92m=== EWS Cleaner - Linux Edition ===\033[0m")
        print("\033[36mMonitoring started. Press Ctrl+C to stop.\033[0m")
        print("-" * 80)
        
        # Main loop
        while not stop_event.is_set():
            try:
                # Check data
                if not data_queue.empty():
                    data = data_queue.get_nowait()
                    
                    if data["type"] == "stats":
                        stats_data = data
                        
                        # Extract EWS call information
                        if "call_types" in data["data"]:
                            for call_type, stats in data["data"]["call_types"].items():
                                if "last_command" in stats and stats["last_command"]:
                                    timestamp = datetime.now().strftime("%H:%M:%S")
                                    ews_calls.append({
                                        "timestamp": timestamp,
                                        "type": call_type,
                                        "time": stats["avg"],
                                        "details": stats["last_command"]
                                    })
                    elif data["type"] == "progress":
                        # Update progress data
                        progress_data = data["data"]
                        
                        # Add to progress history if it's active progress
                        # and at least 5 seconds have passed since the last entry
                        current_time = time.time()
                        if progress_data and progress_data["active"] and (current_time - last_progress_time >= 5):
                            progress_entry = progress_data.copy()
                            progress_entry["timestamp"] = datetime.now().strftime("%H:%M:%S")
                            progress_history.append(progress_entry)
                            last_progress_time = current_time
                            
                            # Limit history size
                            if len(progress_history) > 20:  # Keep the last 20 entries
                                progress_history = progress_history[-20:]
                
                # Check logs
                if not log_queue.empty():
                    log = log_queue.get_nowait()
                    logs.append(log)
                    
                    # If it's an EWS call log, add it to the calls
                    if "EWS call" in log["message"]:
                        try:
                            parts = log["message"].split(" - ")
                            if len(parts) >= 3:
                                call_type = parts[1].strip()
                                time_str = parts[2].strip().replace("ms", "").strip()
                                time_ms = float(time_str)
                                details = parts[0] if len(parts) == 3 else parts[3]
                                
                                ews_calls.append({
                                    "timestamp": log["timestamp"],
                                    "type": call_type,
                                    "time": time_ms,
                                    "details": details
                                })
                        except:
                            pass
                
                # Limit collections
                if len(logs) > 100:
                    logs = logs[-100:]
                if len(ews_calls) > max_calls_to_display * 2:
                    ews_calls = ews_calls[-max_calls_to_display:]
                
                # Display statistics and latest calls every 2 seconds
                if stats_data and (time.time() % 2) < 0.1:
                    # Clear screen (Linux compatible)
                    print("\033c", end="")
                    
                    # Display logo
                    print_logo()
                    
                    # Current date and time
                    current_time = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                    print(f"\033[97m{current_time:^80}\033[0m")
                    print("-" * 80)
                    
                    # Display progress information if available (current progress)
                    if progress_data and progress_data["active"]:
                        print("\033[93m=== CURRENT PROCESSING PROGRESS ===\033[0m")
                        print(f"Folder: \033[96m{progress_data['folder_name']}\033[0m")
                        formatted_time = format_time_remaining(progress_data['est_time'])
                        print(f"Processed: \033[92m{progress_data['processed']}\033[0m items | " +
                              f"Remaining: \033[93m{progress_data['remaining']}\033[0m | " +
                              f"Speed: \033[96m{progress_data['speed']:.2f}\033[0m items/sec | " +
                              f"Time remaining: \033[95m{formatted_time}\033[0m")
                        print("-" * 80)
                    
                    # Display progress history
                    if progress_history:
                        print("\033[93m=== PROGRESS HISTORY ===\033[0m")
                        print(f"{'Time':<10} {'Folder':<20} {'Processed':<10} {'Remaining':<10} {'Speed/s':<10} {'Time rem.':<10}")
                        print("-" * 80)
                        
                        # Display the last 5 history entries in reverse order (most recent at the top)
                        for entry in reversed(progress_history[-5:]):
                            formatted_time = format_time_remaining(entry['est_time'])
                            print(f"{entry['timestamp']:<10} " +
                                 f"\033[96m{entry['folder_name'][:20]:<20}\033[0m " +
                                 f"\033[92m{entry['processed']:<10}\033[0m " +
                                 f"\033[93m{entry['remaining']:<10}\033[0m " +
                                 f"\033[96m{entry['speed']:.2f}\033[0m".ljust(11) +
                                 f"\033[95m{formatted_time}\033[0m")
                        print("-" * 80)
                    
                    # Display statistics
                    stats = stats_data["data"]["stats"]
                    print("\033[96m--- Statistics ---\033[0m")
                    print(f"Active calls: \033[92m{stats['active']}\033[0m | Total: \033[92m{stats['count']}\033[0m")
                    
                    if stats['count'] > 0:
                        min_time = stats['min']
                        avg_time = stats['avg']
                        max_time = stats['max']
                        
                        # Colorize according to times
                        min_color = "\033[92m"  # Green
                        avg_color = "\033[92m"  # Green by default
                        max_color = "\033[92m"  # Green by default
                        
                        if avg_time > 1000:
                            avg_color = "\033[91m"  # Red
                        elif avg_time > 500:
                            avg_color = "\033[93m"  # Yellow
                            
                        if max_time > 1000:
                            max_color = "\033[91m"  # Red
                        elif max_time > 500:
                            max_color = "\033[93m"  # Yellow
                        
                        print(f"Min: {min_color}{min_time:.2f}ms\033[0m | " +
                              f"Avg: {avg_color}{avg_time:.2f}ms\033[0m | " +
                              f"Max: {max_color}{max_time:.2f}ms\033[0m")
                    
                    # Display recent logs
                    print("\n\033[96m--- Recent Logs ---\033[0m")
                    for log in logs[-5:]:  # Limit to last 5 logs
                        color = "\033[92m"  # Green by default
                        if log["level"] == "ERROR":
                            color = "\033[91m"  # Red
                        elif log["level"] == "WARN":
                            color = "\033[93m"  # Yellow
                        
                        print(f"[{log['timestamp']}] {color}{log['message']}\033[0m")
                    
                    # Display recent EWS calls
                    print("\n\033[96m--- EWS Request Monitoring ---\033[0m")
                    print(f"{'Time':<10} {'Type':<15} {'Time (ms)':<12} {'Details':<200}")
                    print("-" * 240)
                    
                    # Limit the number of errors of each type to avoid them remaining displayed permanently
                    error_counts = {}  # To count errors of each type
                    filtered_calls = []
                    
                    # Go through recent calls and limit errors
                    for call in ews_calls[-30:]:  # Consider the last 30 calls
                        call_type = call["type"]
                        
                        # If it's an error, keep only the most recent of each type
                        if call_type.startswith("error_"):
                            if call_type not in error_counts:
                                error_counts[call_type] = 0
                            
                            # Keep only one error of each type
                            if error_counts[call_type] < 1:
                                filtered_calls.append(call)
                                error_counts[call_type] += 1
                        else:
                            # For normal calls, add them all
                            filtered_calls.append(call)
                    
                    # Keep only the 10 most recent calls after filtering
                    recent_calls = sorted(filtered_calls[-10:], key=lambda x: x.get('timestamp', ''))
                    
                    for call in reversed(recent_calls):  # Display most recent first
                        # Colorize according to time
                        time_color = "\033[92m"  # Green
                        if call["time"] > 1000:
                            time_color = "\033[91m"  # Red
                        elif call["time"] > 500:
                            time_color = "\033[93m"  # Yellow
                        
                        # Truncate details if too long
                        details = call["details"]
                        if len(details) > 200:  # Increased to 200 characters
                            details = details[:197] + "..."
                        
                        print(f"{call['timestamp']:<10} " +
                              f"\033[96m{call['type'][:15]:<15}\033[0m " +
                              f"{time_color}{call['time']:.2f}ms\033[0m ".ljust(12) +
                              f"{details:<200}")  # Increased to 200 characters
                    
                    print("\n\033[36mPress Ctrl+C to stop monitoring.\033[0m")
                
                # Small pause to avoid consuming too much CPU
                time.sleep(0.1)
                
            except queue.Empty:
                pass
            except KeyboardInterrupt:
                break
            except Exception as e:
                print(f"\033[91mError in interface: {e}\033[0m")
        
        print("\033[92mMonitoring finished.\033[0m")
    except Exception as e:
        print(f"\033[91mError in interface process: {e}\033[0m")
        import traceback
        traceback.print_exc()

# Modify the main function to set up the new flow
def main():
    # Ensure that multiprocessing is properly initialized
    multiprocessing.freeze_support()
    
    # Display the logo
    print_logo()
    
    # Enhanced command line options
    if len(sys.argv) > 1:
        # Check if --help is requested
        if sys.argv[1] in ['-h', '--help']:
            print(f"{Fore.GREEN}Usage: {sys.argv[0]} [OPTIONS] [username] [password] [impersonated_user]{Style.RESET_ALL}")
            print(f"{Fore.CYAN}Options:{Style.RESET_ALL}")
            print(f"  -h, --help          Display this help")
            print(f"  --no-log-window     Disable opening a separate log window")
            print(f"  --console-log       Enable log display in the main console")
            print(f"  --no-stats-window   Disable opening a separate statistics window")
            print(f"  --server SERVER     Specify the Exchange server")
            print(f"  --classic-ui        Use the classic interface instead of the rich unified interface")
            print(f"  --auto-monitor      Automatically start EWS monitoring (otherwise: manual activation)")
            sys.exit(0)
    
    # Process options
    server = ''  # Server must be specified
    show_log_window = True
    show_stats_window = True
    console_log = False
    use_unified_interface = True
    auto_monitor = False  # By default, don't start monitoring automatically
    
    # Filter arguments to extract options
    filtered_args = []
    i = 1
    while i < len(sys.argv):
        if sys.argv[i] == '--no-log-window':
            show_log_window = False
        elif sys.argv[i] == '--no-stats-window':
            show_stats_window = False
        elif sys.argv[i] == '--console-log':
            console_log = True
        elif sys.argv[i] == '--classic-ui':
            use_unified_interface = False
        elif sys.argv[i] == '--auto-monitor':
            auto_monitor = True
        elif sys.argv[i] == '--server' and i+1 < len(sys.argv):
            server = sys.argv[i+1]
            i += 1
        else:
            filtered_args.append(sys.argv[i])
        i += 1
    
    # Create and configure the EWS logger
    ews_logger.log_to_console = console_log
    
    # Display the EWS logs window if requested
    if show_log_window:
        ews_logger.show_log_window()
    
    # Display the statistics window if requested
    if show_stats_window:
        ews_stats_window.show_stats_window()
    
    # Check connection arguments
    if len(filtered_args) >= 2:
        username = filtered_args[0]
        password = filtered_args[1]
        impersonated_user = filtered_args[2] if len(filtered_args) > 2 else None
    else:
        # Get credentials interactively if not provided via command line
        username, password = get_credentials()

        # Get impersonated user if needed
        impersonate = input(f"{Fore.GREEN}Do you want to impersonate another user? (y/n): {Style.RESET_ALL}").lower() == 'y'
        impersonated_user = None
        if impersonate:
            impersonated_user = input(f"{Fore.GREEN}Enter email address to impersonate: {Style.RESET_ALL}")

    try:
        # Before connecting to the server, install the EWS call interceptor
        intercept_ews_calls()
        print(f"{Fore.CYAN}Individual monitoring of EWS calls enabled{Style.RESET_ALL}")
        
        # Connect to account
        print(f"\n{Fore.CYAN}Connecting to Exchange server {server}...{Style.RESET_ALL}")
        
        # Update the connect_to_account function to use the specified server
        def connect_with_server(username, password, impersonated_user=None):
            credentials = Credentials(username, password)
            config = Configuration(server=server, credentials=credentials)
            
            if impersonated_user:
                return Account(
                    primary_smtp_address=impersonated_user,
                    config=config,
                    access_type=DELEGATE,
                    autodiscover=False
                )
            else:
                return Account(
                    primary_smtp_address=username,
                    config=config,
                    access_type=DELEGATE,
                    autodiscover=False
                )
        
        account = connect_with_server(username, password, impersonated_user)
        print(f"{Fore.GREEN}Connected successfully to {Fore.YELLOW}{account.primary_smtp_address}{Style.RESET_ALL}")
        
        if use_unified_interface:
            # Use the unified interface but WITHOUT starting monitoring automatically
            ews_unified_interface.start(start_monitoring=auto_monitor)
            
            # Get the list of folders
            print(f"\n{Fore.CYAN}Listing folders...{Style.RESET_ALL}")
            folders = []
            for i, folder in enumerate(account.root.walk(), 1):
                folder_info = {
                    "index": i,
                    "name": folder.name,
                    "path": folder.absolute,
                    "total_count": folder.total_count,
                    "object": folder
                }
                folders.append(folder_info)
                ews_unified_interface.add_log(f"Found folder: {folder.name} ({folder.total_count} items)")
            
            # Send folders to the interface
            ews_unified_interface.add_folders(folders)
            
            # Variable to track interface state
            current_view = "main"  # "main" or "stats"
            
            # Display instructions
            print(f"\n{Fore.YELLOW}======= AVAILABLE COMMANDS ======={Style.RESET_ALL}")
            print(f"{Fore.CYAN}Enter a command:{Style.RESET_ALL}")
            print(f"{Fore.GREEN}  1-9{Style.RESET_ALL} : Select a folder by its number")
            print(f"{Fore.GREEN}  m{Style.RESET_ALL}   : Enable/disable EWS monitoring")
            print(f"{Fore.GREEN}  s{Style.RESET_ALL}   : Display EWS statistics")
            print(f"{Fore.GREEN}  q{Style.RESET_ALL}   : Quit the program")
            
            # Display folders
            print(f"\n{Fore.YELLOW}======= AVAILABLE FOLDERS ======={Style.RESET_ALL}")
            for i, folder in enumerate(folders, 1):
                print(f"{Fore.WHITE}{i}. {Fore.GREEN}{folder['name']} {Fore.CYAN}({folder['total_count']} items){Style.RESET_ALL}")
            
            # Use a simple input mode for Linux which is more reliable
            print(f"\n{Fore.YELLOW}Standard input mode for Linux activated.{Style.RESET_ALL}")
            
            # Main loop to process interface commands
            try:
                while True:
                    # Display prompt and wait for input
                    user_input = input(f"{Fore.GREEN}Command > {Style.RESET_ALL}").strip().lower()
                    
                    # No input, continue
                    if not user_input:
                        continue
                    
                    # Take the first character as the command
                    key = user_input[0]
                    
                    # Process the command
                    if key == 'q':
                        # Quit the application
                        if current_view == "stats":
                            # If in stats mode, return to the menu
                            ews_stats_window.exit_stats_window()
                            current_view = "main"
                            print(f"\n{Fore.YELLOW}======= BACK TO MAIN MENU ======={Style.RESET_ALL}")
                            print(f"{Fore.CYAN}Use 's' to see statistics or 'q' to quit{Style.RESET_ALL}")
                        else:
                            # Otherwise quit the application
                            print(f"\n{Fore.YELLOW}Exiting program...{Style.RESET_ALL}")
                            break
                    elif key == 'm':
                        # Enable/disable EWS monitoring
                        if ews_unified_interface.monitoring_active:
                            # Disable monitoring
                            ews_unified_interface.stop()
                            ews_unified_interface.start(start_monitoring=False)
                            print(f"{Fore.YELLOW}EWS monitoring disabled{Style.RESET_ALL}")
                        else:
                            # Enable monitoring
                            if ews_unified_interface.start_monitoring():
                                print(f"{Fore.GREEN}EWS monitoring enabled. EWS requests will be displayed in a separate window.{Style.RESET_ALL}")
                            else:
                                print(f"{Fore.RED}Unable to enable EWS monitoring.{Style.RESET_ALL}")
                    elif key == 's' and current_view == "main":
                        # Display statistics
                        current_view = "stats"
                        print(f"\n{Fore.YELLOW}======= DISPLAYING STATISTICS ======={Style.RESET_ALL}")
                        print(f"{Fore.CYAN}Use 'q' to return to the main menu{Style.RESET_ALL}")
                        ews_stats_window.show_stats_window()
                    elif key.isdigit() and current_view == "main":
                        # Try to convert the complete input in case of multi-digit number
                        try:
                            folder_num = int(user_input)
                            if 1 <= folder_num <= len(folders):
                                selected_folder = folders[folder_num-1]
                                print(f"\n{Fore.YELLOW}Processing folder: {Fore.GREEN}{selected_folder['name']}{Style.RESET_ALL}")
                                
                                # Ask for confirmation
                                print(f"{Fore.RED}Do you really want to empty this folder containing {selected_folder['total_count']} items? (y/n){Style.RESET_ALL}")
                                
                                # Read confirmation
                                confirm = input(f"{Fore.GREEN}Confirmation > {Style.RESET_ALL}").strip().lower()
                                
                                if confirm == 'y':
                                    # Enable monitoring if not already done
                                    if not ews_unified_interface.monitoring_active:
                                        print(f"{Fore.YELLOW}Enabling EWS monitoring to track operations...{Style.RESET_ALL}")
                                        ews_unified_interface.start_monitoring()
                                        time.sleep(1)  # Allow time for monitoring to start
                                    
                                    # Empty the folder
                                    try:
                                        empty_folder(selected_folder["object"])
                                        print(f"\n{Fore.GREEN}Operation completed.{Style.RESET_ALL}")
                                    except Exception as e:
                                        print(f"\n{Fore.RED}Error: {e}{Style.RESET_ALL}")
                                else:
                                    print(f"\n{Fore.YELLOW}Operation cancelled.{Style.RESET_ALL}")
                            else:
                                print(f"{Fore.RED}Invalid folder number. Please enter a number between 1 and {len(folders)}.{Style.RESET_ALL}")
                        except ValueError:
                            print(f"{Fore.RED}Invalid input. Please enter a number to select a folder.{Style.RESET_ALL}")
                    elif key == 'h' or key == '?':
                        # Display help
                        print(f"\n{Fore.YELLOW}======= HELP ======={Style.RESET_ALL}")
                        print(f"{Fore.CYAN}Available commands:{Style.RESET_ALL}")
                        print(f"{Fore.GREEN}  1-9{Style.RESET_ALL} : Select a folder by its number")
                        print(f"{Fore.GREEN}  m{Style.RESET_ALL}   : Enable/disable EWS monitoring")
                        print(f"{Fore.GREEN}  s{Style.RESET_ALL}   : Display EWS statistics")
                        print(f"{Fore.GREEN}  q{Style.RESET_ALL}   : Quit the program")
                        print(f"{Fore.GREEN}  h, ?{Style.RESET_ALL} : Display this help")
                    else:
                        print(f"{Fore.RED}Command not recognized. Type 'h' to display help.{Style.RESET_ALL}")
                    
                    # Check if the user requested to quit statistics
                    if ews_stats_window.user_exit_requested:
                        ews_stats_window.user_exit_requested = False
                        current_view = "main"
                        print(f"\n{Fore.YELLOW}======= BACK TO MAIN MENU ======={Style.RESET_ALL}")
                        print(f"{Fore.CYAN}Use 's' to see statistics or 'q' to quit{Style.RESET_ALL}")
                    
                    # Process unified interface commands
                    command = ews_unified_interface.get_command(timeout=0.1)
                    if command:
                        if command["command"] == "quit":
                            break
                        elif command["command"] == "stats":
                            # Switch to statistics view
                            if current_view != "stats":
                                ews_stats_window.show_stats_window()
                                current_view = "stats"
                                print(f"{Fore.YELLOW}Displaying EWS statistics. Use 'q' to return to main menu.{Style.RESET_ALL}")
                        elif command["command"] == "main":
                            # Return to main view
                            if current_view != "main":
                                ews_stats_window.exit_stats_window()
                                current_view = "main"
                        elif command["command"] == "process_folder":
                            folder_index = command["folder_index"]
                            if 0 <= folder_index < len(folders):
                                selected_folder = folders[folder_index]

                                # Inform the interface that processing is starting
                                ews_unified_interface.data_queue.put({
                                    "type": "processing_start",
                                    "data": {"folder": selected_folder}
                                })

                                ews_unified_interface.add_log(f"Starting processing of folder: {selected_folder['name']}", "INFO")

                                # Start processing
                                try:
                                    # Empty the folder
                                    empty_folder(selected_folder["object"])

                                    ews_unified_interface.add_log(f"Finished processing folder: {selected_folder['name']}", "INFO")
                                except Exception as e:
                                    ews_unified_interface.add_log(f"Error processing folder: {e}", "ERROR")

                                # Inform the interface that processing is complete
                                ews_unified_interface.data_queue.put({
                                    "type": "processing_end",
                                    "data": {"success": True}
                                })
                        elif command["command"] == "cancel_processing":
                            ews_unified_interface.add_log("Processing cancelled by user", "WARN")

                    # Small pause to avoid consuming too much CPU
                    time.sleep(0.1)
            finally:
                pass  # No restoration needed in standard input mode
        else:
            # Classic interface (existing code)
            # List folders and let user choose
            folders = list_folders(account)
            while True:
                choice = input(f"\n{Fore.GREEN}Enter folder number to empty (or 'q' to quit): {Style.RESET_ALL}")
                if choice.lower() == 'q':
                    break

                try:
                    folder_index = int(choice) - 1
                    if 0 <= folder_index < len(folders):
                        selected_folder = folders[folder_index]

                        # Confirm deletion
                        confirm = input(f"{Fore.RED}Are you sure you want to PERMANENTLY delete all {Fore.YELLOW}{selected_folder.total_count}{Fore.RED} items from '{Fore.YELLOW}{selected_folder.name}{Fore.RED}'? (yes/no): {Style.RESET_ALL}")
                        if confirm.lower() == 'yes':
                            # Empty the folder
                            empty_folder(selected_folder)
                        else:
                            print(f"{Fore.YELLOW}Operation cancelled.{Style.RESET_ALL}")
                    else:
                        print(f"{Fore.RED}Invalid folder number.{Style.RESET_ALL}")
                except ValueError:
                    print(f"{Fore.RED}Please enter a valid number.{Style.RESET_ALL}")
                except Exception as e:
                    print(f"{Fore.RED}Error processing folder: {e}{Style.RESET_ALL}")

    except Exception as e:
        print(f"{Fore.RED}Error: {e}{Style.RESET_ALL}")
    finally:
        # Stop all processes
        if use_unified_interface:
            ews_unified_interface.stop()
        else:
            ews_logger.stop()
            ews_stats_window.stop()

        # Wait a moment to allow seeing the statistics
        print(f"\n{Fore.GREEN}Program completed.{Style.RESET_ALL}")


if __name__ == "__main__":
    main() 