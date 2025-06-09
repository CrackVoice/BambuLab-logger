#!/usr/bin/env python3
"""
Bambu Lab Print Logger
Automatically logs print data to Excel using local REST API calls.
LAN-only version for direct printer communication.
"""

import json
import time
import sys
import argparse
import requests
import threading
from datetime import datetime, timedelta
from dataclasses import dataclass
from typing import Optional, Dict, Any
import pandas as pd
import os
import urllib3

# Disable SSL warnings for local printer connections
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


@dataclass
class PrintLog:
    """Represents a single print log entry"""
    start_time: str
    end_time: str
    print_duration: str
    duration_minutes: int
    gcode_file: str
    filament_type: str
    filament_used_grams: float
    notes: str = ""
    bed_temp: float = 0.0
    nozzle_temp: float = 0.0


class BambuLocalAPILogger:
    """Bambu Lab printer logger using local REST API"""
    
    def __init__(self, bambu_ip: str, access_code: str, excel_file: str = "print_log.xlsx"):
        self.bambu_ip = bambu_ip
        self.access_code = access_code
        self.excel_file = excel_file
        
        # API configuration
        self.base_url = f"http://{bambu_ip}"  # Try HTTP first
        self.https_base_url = f"https://{bambu_ip}"  # Fallback to HTTPS
        self.use_https = False
        
        # Session for connection reuse
        self.session = requests.Session()
        self.session.verify = False  # For local HTTPS connections
        
        # Current print tracking
        self.print_start_time: Optional[datetime] = None
        self.current_gcode_file: str = ""
        self.is_printing = False
        self.last_progress = 0
        self.print_start_gcode_time: int = 0
        
        # Print data collection
        self.bed_temp = 0.0
        self.nozzle_temp = 0.0
        self.current_filament_type = ""
        
        # Polling control
        self.polling = False
        self.poll_interval = 3  # seconds
        self.message_count = 0
        
        # Initialize Excel file
        self.init_excel_file()

    def test_connection(self) -> bool:
        """Test connection to local printer API"""
        print(f" Testing connection to {self.bambu_ip}...")
        
        # Test both HTTP and HTTPS
        for use_https in [False, True]:
            base_url = self.https_base_url if use_https else self.base_url
            protocol = "HTTPS" if use_https else "HTTP"
            
            print(f" Trying {protocol} connection...")
            
            # Common API endpoints to try
            endpoints = [
                "/v1/status",
                "/api/v1/status", 
                "/api/status",
                "/status"
            ]
            
            for endpoint in endpoints:
                try:
                    url = f"{base_url}{endpoint}"
                    headers = self.get_headers()
                    
                    response = self.session.get(url, headers=headers, timeout=5)
                    
                    if response.status_code == 200:
                        print(f" {protocol} connection successful on {endpoint}")
                        self.base_url = base_url
                        self.use_https = use_https
                        
                        # Verify we can get printer data
                        data = response.json()
                        if self.validate_printer_data(data):
                            print(f" Printer data accessible")
                            return True
                        else:
                            print(f" Connected but printer data format unexpected")
                            
                    elif response.status_code == 401:
                        print(f" Authentication failed on {endpoint} - check access code")
                        
                    elif response.status_code == 404:
                        continue  # Try next endpoint
                        
                except requests.exceptions.ConnectionError:
                    continue  # Try next endpoint or protocol
                except requests.exceptions.Timeout:
                    continue
                except Exception as e:
                    continue
        
        print(f"Could not establish connection to {self.bambu_ip}")
        print(f"Possible issues:")
        print(f"    - Printer is not powered on")
        print(f"    - Wrong IP address")
        print(f"    - Wrong access code")
        print(f"    - Printer doesn't support API access")
        return False

    def validate_printer_data(self, data: Dict[str, Any]) -> bool:
        """Validate that we received expected printer data"""
        # Look for common printer data fields
        expected_fields = ['print', 'status', 'state', 'progress', 'temperature']
        return any(field in data for field in expected_fields)

    def get_headers(self) -> Dict[str, str]:
        """Get headers for API requests"""
        headers = {
            'Content-Type': 'application/json',
            'User-Agent': 'BambuLocalLogger/1.0'
        }
        
        if self.access_code:
            # Try different authentication methods
            headers['Authorization'] = f'Bearer {self.access_code}'
            headers['X-Access-Code'] = self.access_code
        
        return headers

    def get_printer_status(self) -> Optional[Dict[str, Any]]:
        """Get current printer status via local API"""
        try:
            # Try the working endpoint we found during connection test
            endpoints = ["/v1/status", "/api/v1/status", "/api/status", "/status"]
            
            for endpoint in endpoints:
                url = f"{self.base_url}{endpoint}"
                headers = self.get_headers()
                
                response = self.session.get(url, headers=headers, timeout=3)
                
                if response.status_code == 200:
                    return response.json()
                elif response.status_code == 404:
                    continue
                else:
                    # Log error but continue trying
                    if self.message_count <= 3:
                        print(f"  API returned HTTP {response.status_code} for {endpoint}")
            
            return None
            
        except requests.exceptions.RequestException as e:
            if self.message_count <= 3:
                print(f" API request error: {e}")
            return None
        except Exception as e:
            if self.message_count <= 3:
                print(f" Unexpected error: {e}")
            return None

    def extract_print_data(self, status_data: Dict[str, Any]) -> Dict[str, Any]:
        """Extract print information from status data"""
        if not status_data:
            return {}
        
        # Handle different possible API response formats
        print_data = status_data.get('print', status_data)
        
        # Extract basic print information
        extracted = {
            'progress': self.safe_get_numeric(print_data, ['mc_percent', 'progress', 'percent'], 0),
            'state': self.safe_get_string(print_data, ['gcode_state', 'state', 'status'], ''),
            'gcode_file': self.safe_get_string(print_data, ['gcode_file', 'filename', 'file'], ''),
            'bed_temp': self.safe_get_numeric(print_data, ['bed_temper', 'bed_temp', 'bed_temperature'], 0.0),
            'nozzle_temp': self.safe_get_numeric(print_data, ['nozzle_temper', 'nozzle_temp', 'nozzle_temperature'], 0.0),
            'remaining_time': self.safe_get_numeric(print_data, ['mc_remaining_time', 'remaining_time', 'time_remaining'], 0),
            'start_time': self.safe_get_numeric(print_data, ['gcode_start_time', 'start_time', 'print_start_time'], 0)
        }
        
        # Extract filament information
        ams_data = print_data.get('ams', {})
        extracted['filament_type'] = self.extract_filament_info(ams_data, print_data)
        
        return extracted

    def safe_get_numeric(self, data: Dict[str, Any], keys: list, default: float = 0.0) -> float:
        """Get numeric value from data using multiple possible keys"""
        for key in keys:
            if key in data:
                try:
                    return float(data[key])
                except (ValueError, TypeError):
                    continue
        return default

    def safe_get_string(self, data: Dict[str, Any], keys: list, default: str = '') -> str:
        """Get string value from data using multiple possible keys"""
        for key in keys:
            if key in data and data[key]:
                return str(data[key])
        return default

    def extract_filament_info(self, ams_data: Dict[str, Any], print_data: Dict[str, Any]) -> str:
        """Extract filament type from AMS data or print data"""
        try:
            # Try to get from direct filament field first
            filament_direct = self.safe_get_string(print_data, ['filament_type', 'material', 'filament'], '')
            if filament_direct:
                return filament_direct
            
            # Try to extract from AMS data
            if ams_data and "ams" in ams_data:
                current_tray = ams_data.get("tray_now", "0")
                
                try:
                    tray_num = int(current_tray)
                    ams_index = tray_num // 4
                    tray_index = tray_num % 4
                except:
                    return "Unknown"
                
                ams_list = ams_data.get("ams", [])
                if ams_index < len(ams_list):
                    trays = ams_list[ams_index].get("tray", [])
                    if tray_index < len(trays):
                        return trays[tray_index].get("tray_type", "Unknown")
            
            return "Unknown"
        except Exception:
            return "Unknown"

    def init_excel_file(self):
        """Initialize Excel file with headers if it doesn't exist"""
        if not os.path.exists(self.excel_file):
            df = pd.DataFrame(columns=[
                'Start Time', 'End Time', 'Print Duration', 'Duration (min)',
                'G-code File', 'Filament Type', 'Filament Used (g)', 
                'Bed Temp (°C)', 'Nozzle Temp (°C)', 'Notes'
            ])
            df.to_excel(self.excel_file, index=False)
            print(f" Created new Excel file: {self.excel_file}")
        else:
            print(f" Using existing Excel file: {self.excel_file}")

    def start_print_tracking(self, gcode_file: str, gcode_start_time: int, filament_type: str):
        """Start tracking a new print"""
        self.is_printing = True
        self.print_start_time = datetime.now()
        self.current_gcode_file = gcode_file
        self.current_filament_type = filament_type
        self.print_start_gcode_time = gcode_start_time
        
        print(f"\n" + "="*60)
        print(f" PRINT STARTED")
        print(f" Start Time: {self.print_start_time.strftime('%Y-%m-%d %H:%M:%S')}")
        print(f" File: {os.path.basename(gcode_file) if gcode_file else 'Unknown'}")
        print(f" Filament: {filament_type}")
        print(f"  Bed: {self.bed_temp}°C | Nozzle: {self.nozzle_temp}°C")
        print("="*60)

    def update_progress(self, progress: int, remaining_time: int):
        """Update progress display"""
        remaining_str = f"{remaining_time}min" if remaining_time > 0 else "Unknown"
        current_time = datetime.now().strftime('%H:%M:%S')
        filament_str = self.current_filament_type if self.current_filament_type else "Unknown"
        print(f"\r [{current_time}] Progress: {progress:3d}% | Remaining: {remaining_str:>8} | {filament_str}", end="", flush=True)

    def end_print_tracking(self, failed: bool = False):
        """End print tracking and log to Excel"""
        if not self.is_printing or not self.print_start_time:
            return
        
        self.is_printing = False
        end_time = datetime.now()
        duration = end_time - self.print_start_time
        duration_minutes = int(duration.total_seconds() / 60)
        
        status_emoji = "" if failed else ""
        status_text = "FAILED" if failed else "COMPLETED"
        
        print(f"\n" + "="*60)
        print(f"{status_emoji} PRINT {status_text}")
        print(f"End Time: {end_time.strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"Duration: {self.format_duration(duration_minutes)}")
        
        # Estimate filament usage (rough calculation based on print time)
        estimated_filament = self.estimate_filament_usage(duration_minutes)
        
        # Create log entry
        log_entry = PrintLog(
            start_time=self.print_start_time.strftime('%Y-%m-%d %H:%M:%S'),
            end_time=end_time.strftime('%Y-%m-%d %H:%M:%S'),
            print_duration=self.format_duration(duration_minutes),
            duration_minutes=duration_minutes,
            gcode_file=os.path.basename(self.current_gcode_file) if self.current_gcode_file else "Unknown",
            filament_type=self.current_filament_type if self.current_filament_type else "Unknown",
            filament_used_grams=estimated_filament,
            bed_temp=self.bed_temp,
            nozzle_temp=self.nozzle_temp,
            notes="FAILED PRINT" if failed else ""
        )
        
        # Save to Excel
        self.save_to_excel(log_entry)
        
        print(f"Logged to Excel: {self.excel_file}")
        print(f"Manual updates recommended:")
        print(f"    - Verify filament used (estimated: {estimated_filament:.1f}g)")
        print(f"    - Add notes about print quality/issues")
        print("="*60)
        print("Waiting for next print...")

    def estimate_filament_usage(self, duration_minutes: int) -> float:
        """Rough estimate of filament usage based on print time"""
        # Very rough estimate: 10g per hour average
        return (duration_minutes / 60.0) * 10.0

    def format_duration(self, minutes: int) -> str:
        """Format duration in a readable way"""
        hours = minutes // 60
        mins = minutes % 60
        if hours > 0:
            return f"{hours}h {mins}m"
        else:
            return f"{mins}m"

    def save_to_excel(self, log_entry: PrintLog):
        """Save print log to Excel file"""
        try:
            # Read existing data
            try:
                df = pd.read_excel(self.excel_file)
            except:
                df = pd.DataFrame()
            
            # Create new row
            new_row = {
                'Start Time': log_entry.start_time,
                'End Time': log_entry.end_time,
                'Print Duration': log_entry.print_duration,
                'Duration (min)': log_entry.duration_minutes,
                'G-code File': log_entry.gcode_file,
                'Filament Type': log_entry.filament_type,
                'Filament Used (g)': log_entry.filament_used_grams,
                'Bed Temp (°C)': log_entry.bed_temp,
                'Nozzle Temp (°C)': log_entry.nozzle_temp,
                'Notes': log_entry.notes
            }
            
            # Add new row
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
            
            # Save to Excel
            df.to_excel(self.excel_file, index=False)
            
        except Exception as e:
            print(f" Error saving to Excel: {e}")

    def monitor_prints(self):
        """Main monitoring loop"""
        print(f"Starting print monitoring...")
        print(f"Polling every {self.poll_interval} seconds")
        print("Waiting for prints to start...\n")
        
        consecutive_errors = 0
        max_errors = 10
        
        while self.polling:
            try:
                # Get printer status
                status = self.get_printer_status()
                
                if status:
                    consecutive_errors = 0  # Reset error counter
                    self.process_status_update(status)
                else:
                    consecutive_errors += 1
                    if consecutive_errors >= max_errors:
                        print(f"\nToo many consecutive API errors ({max_errors})")
                        print(f"Check printer connection and restart logger")
                        break
                
                self.message_count += 1
                
                # Sleep between polls
                time.sleep(self.poll_interval)
                
            except KeyboardInterrupt:
                break
            except Exception as e:
                print(f"Error in monitoring loop: {e}")
                consecutive_errors += 1
                if consecutive_errors >= max_errors:
                    break
                time.sleep(self.poll_interval)

    def process_status_update(self, status_data: Dict[str, Any]):
        """Process a status update from the printer"""
        print_data = self.extract_print_data(status_data)
        
        progress = int(print_data.get('progress', 0))
        state = print_data.get('state', '').upper()
        gcode_file = print_data.get('gcode_file', '')
        bed_temp = print_data.get('bed_temp', 0.0)
        nozzle_temp = print_data.get('nozzle_temp', 0.0)
        remaining_time = int(print_data.get('remaining_time', 0))
        filament_type = print_data.get('filament_type', 'Unknown')
        start_time = int(print_data.get('start_time', 0))
        
        # Update current temperatures
        self.bed_temp = bed_temp
        self.nozzle_temp = nozzle_temp
        
        # Show confirmation of data reception (only first few times)
        if self.message_count <= 3:
            print(f"Update #{self.message_count}: Progress: {progress}%, State: {state}")
            if self.message_count == 3:
                print(f"API polling working - switching to print monitoring mode")
        
        # Check if print is starting
        if not self.is_printing and state in ['RUNNING', 'PRINTING'] and progress > 0:
            self.start_print_tracking(gcode_file, start_time, filament_type)
        
        # Check if print is completed or failed
        elif self.is_printing and (progress >= 100 or state in ['FINISH', 'FINISHED', 'FAILED', 'PAUSED', 'STOPPED']):
            failed = state in ['FAILED', 'STOPPED']
            self.end_print_tracking(failed)
        
        # Update progress for current print
        if self.is_printing and progress != self.last_progress:
            self.update_progress(progress, remaining_time)
            self.last_progress = progress
        
        # Update filament type if we got better info
        if filament_type != "Unknown":
            self.current_filament_type = filament_type
        
        # Show periodic status updates when not printing
        if not self.is_printing and self.message_count % 20 == 0:  # Every ~1 minute
            current_time = datetime.now().strftime('%H:%M:%S')
            print(f"💤 [{current_time}] Idle - Bed: {bed_temp}°C | Nozzle: {nozzle_temp}°C | State: {state}")

    def run(self):
        """Start the logging process"""
        try:
            print("     BAMBU LAB LOCAL API PRINT LOGGER")
            print("="*50)
            
            # Test connection first
            if not self.test_connection():
                print("\n Connection test failed. Please check:")
                print("     1. Printer IP address is correct")
                print("     2. Printer is powered on and connected to network") 
                print("     3. Access code is correct")
                print("     4. Printer firmware supports API access")
                return
            
            print(f"\nStarting print logger...")
            print(f"Monitoring: {self.bambu_ip}")
            print(f"Excel file: {self.excel_file}")
            
            # Start monitoring
            self.polling = True
            self.monitor_prints()
            
        except KeyboardInterrupt:
            print("\n\nStopping logger...")
            self.display_summary()
        except Exception as e:
            print(f"Fatal error: {e}")
        finally:
            self.polling = False

    def display_summary(self):
        """Display current session summary"""
        try:
            df = pd.read_excel(self.excel_file)
            if len(df) > 0:
                total_prints = len(df)
                total_minutes = df['Duration (min)'].sum()
                total_filament = df['Filament Used (g)'].sum()
                
                print(f"\nSESSION SUMMARY:")
                print(f"    Total prints logged: {total_prints}")
                print(f"    Total print time: {self.format_duration(int(total_minutes))}")
                print(f"    Total filament used: {total_filament:.1f}g")
                
                if total_prints > 0:
                    print(f"\nRecent prints:")
                    for _, row in df.tail(3).iterrows():
                        print(f"   • {row['G-code File']} - {row['Print Duration']} ({row['Filament Type']})")
        except:
            print("No previous prints found")


def get_printer_info():
    """Interactive function to get printer information"""
    print("     BAMBU LAB LOCAL API LOGGER SETUP")
    print("="*50)
    
    while True:
        ip = input("Enter printer IP address (e.g., 192.168.1.100): ").strip()
        if not ip:
            print("IP address cannot be empty")
            continue
        
        # Basic IP validation
        parts = ip.split('.')
        if len(parts) != 4:
            print("Invalid IP format. Use format: 192.168.1.100")
            continue
        
        try:
            for part in parts:
                if not (0 <= int(part) <= 255):
                    raise ValueError
            break
        except ValueError:
            print("Invalid IP address. Each number must be 0-255")
            continue
    
    while True:
        access_code = input("Enter printer access code: ").strip()
        if not access_code:
            print("Access code cannot be empty")
            continue
        break
    
    excel_file = input("Excel filename (press Enter for 'print_log.xlsx'): ").strip()
    if not excel_file:
        excel_file = "print_log.xlsx"
    elif not excel_file.endswith('.xlsx'):
        excel_file += '.xlsx'
    
    return ip, access_code, excel_file


def main():
    parser = argparse.ArgumentParser(description="Log Bambu Lab prints to Excel using local API")
    parser.add_argument("--ip", help="IP address of the Bambu Lab printer")
    parser.add_argument("--code", help="Printer access code")
    parser.add_argument("--excel", "-e", default="print_log.xlsx", 
                       help="Excel file name (default: print_log.xlsx)")
    
    args = parser.parse_args()
    
    # Interactive mode if no arguments provided
    if not args.ip or not args.code:
        ip, access_code, excel_file = get_printer_info()
    else:
        ip, access_code, excel_file = args.ip, args.code, args.excel
    
    logger = BambuLocalAPILogger(ip, access_code, excel_file)
    logger.run()


if __name__ == "__main__":
    main()