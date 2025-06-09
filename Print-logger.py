#!/usr/bin/env python3
"""
Bambu Lab Personal Print Logger with Connection Testing
Automatically logs print data to Excel with proper connection validation.
"""

import json
import time
import sys
import argparse
import socket
import threading
from datetime import datetime, timedelta
from dataclasses import dataclass, asdict
from typing import Optional, Dict, List
import paho.mqtt.client as mqtt
import pandas as pd
import os
from pathlib import Path


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


class BambuPrintLogger:
    def __init__(self, bambu_ip: str, bambu_id: str, excel_file: str = "print_log.xlsx"):
        self.bambu_ip = bambu_ip
        self.bambu_id = bambu_id
        self.excel_file = excel_file
        self.mqtt_client = mqtt.Client()
        
        # Connection status tracking
        self.connection_confirmed = False
        self.data_received = False
        self.connection_timeout = 30  # seconds
        
        # Current print tracking
        self.print_start_time: Optional[datetime] = None
        self.print_start_gcode_time: int = 0
        self.current_gcode_file: str = ""
        self.is_printing = False
        self.last_mc_percent = 0
        
        # Print data collection
        self.bed_temp = 0.0
        self.nozzle_temp = 0.0
        self.current_filament_type = ""
        self.filament_used = 0.0
        
        # Message counter for confirmation
        self.message_count = 0
        
        # MQTT setup
        self.mqtt_client.on_connect = self.on_connect
        self.mqtt_client.on_message = self.on_message
        self.mqtt_client.on_disconnect = self.on_disconnect
        
        # Initialize Excel file
        self.init_excel_file()

    def test_network_connectivity(self) -> bool:
        """Test basic network connectivity to the printer IP"""
        print(f"🔍 Testing network connectivity to {self.bambu_ip}...")
        
        # Test if host is reachable on common ports
        ports_to_test = [1883, 80, 443]  # MQTT, HTTP, HTTPS
        
        for port in ports_to_test:
            try:
                sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                sock.settimeout(5)
                result = sock.connect_ex((self.bambu_ip, port))
                sock.close()
                
                if result == 0:
                    print(f"✅ Network connectivity OK - Port {port} is open")
                    return True
            except Exception as e:
                continue
        
        print(f"❌ Network connectivity FAILED - No open ports found")
        print(f"   Make sure {self.bambu_ip} is the correct IP address")
        print(f"   Check if printer is powered on and connected to network")
        return False

    def test_mqtt_connection(self) -> bool:
        """Test MQTT connection and wait for data"""
        print(f"🔗 Testing MQTT connection to {self.bambu_ip}:1883...")
        
        connection_event = threading.Event()
        data_event = threading.Event()
        
        def on_connect_test(client, userdata, flags, rc):
            if rc == 0:
                print(f"✅ MQTT connection successful")
                topic = f"device/{self.bambu_id}/report"
                client.subscribe(topic)
                print(f"📡 Subscribed to topic: {topic}")
                connection_event.set()
            else:
                print(f"❌ MQTT connection failed (code: {rc})")
                connection_event.set()
        
        def on_message_test(client, userdata, msg):
            try:
                data = json.loads(msg.payload.decode())
                if "print" in data:
                    print(f"✅ Receiving printer data successfully!")
                    print(f"📊 Sample data: Print stage, temperatures, progress, etc.")
                    self.data_received = True
                    data_event.set()
            except:
                pass
        
        # Set up test client
        test_client = mqtt.Client()
        test_client.on_connect = on_connect_test
        test_client.on_message = on_message_test
        
        try:
            test_client.connect(self.bambu_ip, 1883, 60)
            test_client.loop_start()
            
            # Wait for connection
            if not connection_event.wait(timeout=10):
                print(f"❌ MQTT connection timeout")
                test_client.loop_stop()
                test_client.disconnect()
                return False
            
            # Wait for data
            print(f"⏳ Waiting for printer data (up to 15 seconds)...")
            if data_event.wait(timeout=15):
                print(f"🎉 Connection test PASSED - Ready to log prints!")
                test_client.loop_stop()
                test_client.disconnect()
                return True
            else:
                print(f"⚠️  MQTT connected but no data received")
                print(f"   This might be normal if printer is idle")
                print(f"   The logger will still work when printing starts")
                test_client.loop_stop()
                test_client.disconnect()
                return True  # Consider this OK
                
        except Exception as e:
            print(f"❌ MQTT connection error: {e}")
            return False

    def validate_printer_id(self) -> bool:
        """Validate that the printer ID format looks correct"""
        if len(self.bambu_id) < 10:
            print(f"⚠️  Warning: Printer ID '{self.bambu_id}' seems short")
            print(f"   Typical format: 00M00A261900054 (15 characters)")
            response = input("Continue anyway? (y/N): ").lower()
            return response == 'y'
        return True

    def init_excel_file(self):
        """Initialize Excel file with headers if it doesn't exist"""
        if not os.path.exists(self.excel_file):
            df = pd.DataFrame(columns=[
                'Start Time', 'End Time', 'Print Duration', 'Duration (min)',
                'G-code File', 'Filament Type', 'Filament Used (g)', 
                'Bed Temp (°C)', 'Nozzle Temp (°C)', 'Notes'
            ])
            df.to_excel(self.excel_file, index=False)
            print(f"✅ Created new Excel file: {self.excel_file}")
        else:
            print(f"✅ Using existing Excel file: {self.excel_file}")

    def on_connect(self, client, userdata, flags, rc):
        if rc == 0:
            print(f"🔗 Connected to MQTT broker")
            topic = f"device/{self.bambu_id}/report"
            client.subscribe(topic)
            print(f"📡 Listening for printer data...")
            self.connection_confirmed = True
        else:
            print(f"❌ MQTT connection failed (code: {rc})")

    def on_disconnect(self, client, userdata, rc):
        print(f"🔌 Disconnected from MQTT broker")
        if rc != 0:
            print(f"⚠️  Unexpected disconnection (code: {rc})")

    def extract_filament_info(self, ams_data):
        """Extract filament type from AMS data"""
        try:
            if not ams_data or "ams" not in ams_data:
                return "Unknown"
            
            # Get current tray info
            current_tray = ams_data.get("tray_now", "0")
            
            # Parse tray number (format might be like "6" meaning AMS 1, Tray 2)
            try:
                tray_num = int(current_tray)
                ams_index = tray_num // 4  # Each AMS has 4 trays
                tray_index = tray_num % 4
            except:
                return "Unknown"
            
            # Get filament type from the appropriate AMS and tray
            ams_list = ams_data.get("ams", [])
            if ams_index < len(ams_list):
                trays = ams_list[ams_index].get("tray", [])
                if tray_index < len(trays):
                    return trays[tray_index].get("tray_type", "Unknown")
            
            return "Unknown"
        except Exception as e:
            return "Unknown"

    def on_message(self, client, userdata, msg):
        try:
            # Parse JSON message
            data = json.loads(msg.payload.decode())
            print_data = data.get("print", {})
            
            # Increment message counter for debugging
            self.message_count += 1
            
            # Show confirmation of data reception (only first few times)
            if self.message_count <= 3:
                print(f"📨 Message #{self.message_count}: Receiving data from printer")
                if self.message_count == 3:
                    print(f"✅ Data reception confirmed - switching to print monitoring mode")
            
            # Extract key fields
            mc_percent = print_data.get("mc_percent", 0)
            gcode_state = print_data.get("gcode_state", "")
            gcode_file = print_data.get("gcode_file", "")
            gcode_start_time = print_data.get("gcode_start_time", "0")
            bed_temp = print_data.get("bed_temper", 0.0)
            nozzle_temp = print_data.get("nozzle_temper", 0.0)
            mc_remaining_time = print_data.get("mc_remaining_time", 0)
            
            # Extract AMS/filament info
            ams_data = print_data.get("ams", {})
            filament_type = self.extract_filament_info(ams_data)
            
            # Update current temperatures
            self.bed_temp = bed_temp
            self.nozzle_temp = nozzle_temp
            
            # Show current status (less frequent updates)
            if self.message_count % 10 == 0:  # Every 10th message
                print(f"🔄 Status: {gcode_state} | Progress: {mc_percent}% | Bed: {bed_temp}°C | Nozzle: {nozzle_temp}°C")
            
            # Check if print is starting
            if not self.is_printing and gcode_state == "RUNNING" and mc_percent > 0:
                self.start_print_tracking(gcode_file, gcode_start_time, filament_type)
            
            # Check if print is completed
            elif self.is_printing and (mc_percent >= 100 or gcode_state in ["FINISH", "FAILED"]):
                self.end_print_tracking(gcode_state == "FAILED")
            
            # Update progress for current print
            if self.is_printing:
                self.update_progress(mc_percent, mc_remaining_time)
            
            # Update filament type if changed
            if filament_type != "Unknown":
                self.current_filament_type = filament_type
            
            self.last_mc_percent = mc_percent
            self.data_received = True
            
        except json.JSONDecodeError:
            print("❌ Failed to parse JSON message")
        except Exception as e:
            print(f"❌ Error processing message: {e}")

    def start_print_tracking(self, gcode_file: str, gcode_start_time: str, filament_type: str):
        """Start tracking a new print"""
        self.is_printing = True
        self.print_start_time = datetime.now()
        self.current_gcode_file = gcode_file
        self.current_filament_type = filament_type
        
        try:
            self.print_start_gcode_time = int(gcode_start_time)
        except:
            self.print_start_gcode_time = 0
        
        print(f"\n" + "="*60)
        print(f"🚀 PRINT STARTED")
        print(f"⏰ Start Time: {self.print_start_time.strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"📁 File: {gcode_file}")
        print(f"🧵 Filament: {filament_type}")
        print(f"🌡️  Bed: {self.bed_temp}°C | Nozzle: {self.nozzle_temp}°C")
        print("="*60)

    def update_progress(self, mc_percent: int, mc_remaining_time: int):
        """Update progress display"""
        remaining_str = f"{mc_remaining_time}min" if mc_remaining_time > 0 else "Unknown"
        current_time = datetime.now().strftime('%H:%M:%S')
        print(f"\r🖨️  [{current_time}] Progress: {mc_percent:3d}% | Remaining: {remaining_str:>8} | {self.current_filament_type}", end="", flush=True)

    def end_print_tracking(self, failed: bool = False):
        """End print tracking and log to Excel"""
        if not self.is_printing or not self.print_start_time:
            return
        
        self.is_printing = False
        end_time = datetime.now()
        duration = end_time - self.print_start_time
        duration_minutes = int(duration.total_seconds() / 60)
        
        status_emoji = "❌" if failed else "✅"
        status_text = "FAILED" if failed else "COMPLETED"
        
        print(f"\n" + "="*60)
        print(f"{status_emoji} PRINT {status_text}")
        print(f"⏰ End Time: {end_time.strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"⏱️  Duration: {self.format_duration(duration_minutes)}")
        
        # Estimate filament usage (rough calculation)
        estimated_filament = self.estimate_filament_usage(duration_minutes)
        
        # Create log entry
        log_entry = PrintLog(
            start_time=self.print_start_time.strftime('%Y-%m-%d %H:%M:%S'),
            end_time=end_time.strftime('%Y-%m-%d %H:%M:%S'),
            print_duration=self.format_duration(duration_minutes),
            duration_minutes=duration_minutes,
            gcode_file=os.path.basename(self.current_gcode_file),
            filament_type=self.current_filament_type,
            filament_used_grams=estimated_filament,
            bed_temp=self.bed_temp,
            nozzle_temp=self.nozzle_temp,
            notes="FAILED PRINT" if failed else ""
        )
        
        # Save to Excel
        self.save_to_excel(log_entry)
        
        print(f"📊 Logged to Excel: {self.excel_file}")
        print(f"💡 Manual updates needed:")
        print(f"   - Actual filament used (estimated: {estimated_filament:.1f}g)")
        print(f"   - G-code filename if needed")
        print(f"   - Add notes about print quality")
        print("="*60)
        print("⏳ Waiting for next print...")

    def estimate_filament_usage(self, duration_minutes: int) -> float:
        """Rough estimate of filament usage based on print time"""
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
            print(f"❌ Error saving to Excel: {e}")

    def run_connection_tests(self) -> bool:
        """Run all connection tests"""
        print("🔍 RUNNING CONNECTION TESTS")
        print("="*60)
        
        # Test 1: Network connectivity
        if not self.test_network_connectivity():
            return False
        
        # Test 2: Printer ID validation
        if not self.validate_printer_id():
            return False
        
        # Test 3: MQTT connection and data reception
        if not self.test_mqtt_connection():
            return False
        
        print("="*60)
        print("🎉 ALL TESTS PASSED - PRINTER IS ACCESSIBLE")
        print("="*60)
        return True

    def run(self):
        """Start the logging loop"""
        try:
            # Run connection tests first
            if not self.run_connection_tests():
                print("\n❌ Connection tests failed. Please check:")
                print("   1. Printer IP address is correct")
                print("   2. Printer is powered on and connected to network")
                print("   3. Printer ID is correct")
                print("   4. No firewall blocking MQTT port 1883")
                return
            
            print(f"\n🚀 Starting print logger...")
            print(f"📍 Monitoring: {self.bambu_ip} (ID: {self.bambu_id})")
            print(f"📊 Excel file: {self.excel_file}")
            print("⏳ Waiting for prints to start...\n")
            
            self.mqtt_client.connect(self.bambu_ip, 1883, 60)
            self.mqtt_client.loop_start()
            
            # Keep running until interrupted
            while True:
                time.sleep(1)
                
        except KeyboardInterrupt:
            print("\n\n🛑 Stopping logger...")
            self.display_summary()
        except Exception as e:
            print(f"❌ Error: {e}")
        finally:
            self.mqtt_client.loop_stop()
            self.mqtt_client.disconnect()

    def display_summary(self):
        """Display current session summary"""
        try:
            df = pd.read_excel(self.excel_file)
            if len(df) > 0:
                total_prints = len(df)
                total_minutes = df['Duration (min)'].sum()
                total_filament = df['Filament Used (g)'].sum()
                
                print(f"\n📈 SESSION SUMMARY:")
                print(f"   Total prints logged: {total_prints}")
                print(f"   Total print time: {self.format_duration(int(total_minutes))}")
                print(f"   Total filament used: {total_filament:.1f}g")
                
                if total_prints > 0:
                    print(f"\n📋 Recent prints:")
                    for _, row in df.tail(3).iterrows():
                        print(f"   • {row['G-code File']} - {row['Print Duration']} ({row['Filament Type']})")
        except:
            print("📊 No previous prints found")


def get_printer_info():
    """Interactive function to get printer information"""
    print("🖨️  BAMBU LAB PRINT LOGGER SETUP")
    print("="*50)
    
    while True:
        ip = input("Enter printer IP address (e.g., 192.168.1.100): ").strip()
        if not ip:
            print("❌ IP address cannot be empty")
            continue
        
        # Basic IP validation
        parts = ip.split('.')
        if len(parts) != 4:
            print("❌ Invalid IP format. Use format: 192.168.1.100")
            continue
        
        try:
            for part in parts:
                if not (0 <= int(part) <= 255):
                    raise ValueError
            break
        except ValueError:
            print("❌ Invalid IP address. Each number must be 0-255")
            continue
    
    while True:
        printer_id = input("Enter printer ID (e.g., 00M00A261900054): ").strip()
        if not printer_id:
            print("❌ Printer ID cannot be empty")
            continue
        break
    
    excel_file = input("Excel filename (press Enter for 'print_log.xlsx'): ").strip()
    if not excel_file:
        excel_file = "print_log.xlsx"
    elif not excel_file.endswith('.xlsx'):
        excel_file += '.xlsx'
    
    return ip, printer_id, excel_file


def main():
    parser = argparse.ArgumentParser(description="Log Bambu Lab prints to Excel with connection testing")
    parser.add_argument("--ip", help="IP address of the Bambu Lab printer")
    parser.add_argument("--id", help="Bambu Lab printer ID")
    parser.add_argument("--excel", "-e", default="print_log.xlsx", 
                       help="Excel file name (default: print_log.xlsx)")
    
    args = parser.parse_args()
    
    # Interactive mode if no arguments provided
    if not args.ip or not args.id:
        ip, printer_id, excel_file = get_printer_info()
    else:
        ip, printer_id, excel_file = args.ip, args.id, args.excel
    
    logger = BambuPrintLogger(ip, printer_id, excel_file)
    logger.run()


if __name__ == "__main__":
    main()
