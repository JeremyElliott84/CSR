#!/usr/bin/env python3
"""
Meraki Network Refresh Complete Tool
Complete network refresh: keep switches in place, clear old assignments, remove old devices, add new devices, update addresses
"""

import meraki
import json
import time
import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os
import sys
from typing import List, Dict, Optional
from dotenv import load_dotenv

# =============================================================================
# CONFIGURATION
# =============================================================================
# Get the directory where the script/executable is located
if getattr(sys, 'frozen', False):
    # Running as executable
    script_dir = os.path.dirname(sys.executable)
else:
    # Running as script
    script_dir = os.path.dirname(os.path.abspath(__file__))

# Load environment variables from .env file in the same directory as the executable
env_path = os.path.join(script_dir, '.env')
load_dotenv(env_path)

API_KEY = os.getenv('API_KEY')
ORG_ID = os.getenv('ORG_ID')
DATA_FILE_PATH = os.getenv('DATA_FILE_PATH')

SPREADSHEET_COLUMNS = {
    'store_number': 'storeNumber', 'network_id': 'NetworkID', 'template_id': 'New Template ID',
    'mx_a_sn': 'MX-A SN', 'mx_a_name': 'MX-A Name', 'mx_b_sn': 'MX-B SN', 'mx_b_name': 'MX-B Name',
    'cw_a_sn': 'CW-A SN', 'cw_a_name': 'CW-A Name', 'cw_a_ip': 'CW-A IP',
    'cw_b_sn': 'CW-B SN', 'cw_b_name': 'CW-B Name', 'cw_b_ip': 'CW-B IP',
    'mt40_sn': 'MT40 SN', 'mt40_name': 'MT40 Name',
    # Switch columns
    'sw_a_name': 'SW-A Name', 'sw_b_name': 'SW-B Name',
    # Address columns
    'address': 'Address', 'city': 'City', 'state': 'State'
}

PRESERVE_DHCP_CLIENTS = ['MS120-A', 'MS120-B', 'MS130-A', 'MS130-B']  # Switch names will be preserved dynamically
DEVICE_TYPES = {'mx_a': 'MX67 Primary', 'mx_b': 'MX67 Secondary', 'cw_a': 'AP Primary', 'cw_b': 'AP Secondary', 'mt40': 'MT40 Sensor'}

class NetworkRefreshManager:
    def __init__(self, api_key: str, org_id: str = None):
        self.dashboard = meraki.DashboardAPI(api_key, suppress_logging=True)
        self.org_id = org_id if org_id else self.dashboard.organizations.getOrganizations()[0]['id']

    def clear_non_switch_assignments(self, network_id: str):
        """Step 1: Clear all non-switch assignments (keep switches in place)"""
        print("Step 1: Clearing non-switch assignments (preserving switches)...")
        try:
            # Get current network devices to identify switch MAC addresses
            devices = self.dashboard.networks.getNetworkDevices(network_id)
            switch_macs = set()
            for device in devices:
                if device.get('model', '').upper().startswith(('MS120', 'MS130')):
                    mac = device.get('mac', '').lower().replace(':', '')
                    if mac:
                        switch_macs.add(mac)
            
            # Only work with VLAN 1
            dhcp_settings = self.dashboard.appliance.getNetworkApplianceVlan(network_id, "1")
            assignments = dhcp_settings.get('fixedIpAssignments', {})
            new_assignments = {}
            
            # Keep switch assignments (by MAC) and legacy named assignments
            for mac, assignment in assignments.items():
                client_name = assignment.get('name', '').strip()
                mac_clean = mac.lower().replace(':', '')
                
                # Preserve if it's a known switch MAC or legacy switch name
                if mac_clean in switch_macs or client_name in PRESERVE_DHCP_CLIENTS:
                    new_assignments[mac] = assignment
                    print(f"  Preserved: {client_name} -> {assignment.get('ip')} (MAC: {mac})")
            
            self.dashboard.appliance.updateNetworkApplianceVlan(network_id, "1", fixedIpAssignments=new_assignments)
            time.sleep(1)
            print(f"  Cleared {len(assignments) - len(new_assignments)} non-switch assignments")
            return len(assignments) - len(new_assignments)
        except Exception as e:
            print(f"  Error processing VLAN 1: {e}")
            return 0

    def remove_iboot_ranges(self, network_id: str):
        """Step 2: Remove iBoot reserved ranges"""
        print("Step 2: Removing iBoot ranges...")
        try:
            # Only work with VLAN 1
            dhcp_settings = self.dashboard.appliance.getNetworkApplianceVlan(network_id, "1")
            ranges = dhcp_settings.get('reservedIpRanges', [])
            new_ranges = [r for r in ranges if r.get('comment', '').lower() != 'iboot']
            
            if len(new_ranges) != len(ranges):
                self.dashboard.appliance.updateNetworkApplianceVlan(network_id, "1", reservedIpRanges=new_ranges)
                time.sleep(1)
                removed_count = len(ranges) - len(new_ranges)
                print(f"  Removed {removed_count} iBoot ranges")
                return removed_count
            else:
                print("  No iBoot ranges found to remove")
                return 0
        except Exception as e:
            print(f"  Error processing VLAN 1: {e}")
            return 0

    def capture_mx64_static_ip_settings(self, network_id: str):
        """Step 2.5: Capture static IP settings from MX64 devices before removal"""
        print("Step 2.5: Capturing static IP settings from MX64 devices...")
        static_ip_settings = None
        
        try:
            devices = self.dashboard.networks.getNetworkDevices(network_id)
            mx64_devices = [d for d in devices if d.get('model', '').upper().startswith('MX64')]
            
            if not mx64_devices:
                print("  No MX64 devices found")
                return static_ip_settings
            
            for device in mx64_devices:
                serial = device['serial']
                device_name = device.get('name', serial)
                
                try:
                    # Get management interface settings
                    mgmt_settings = self.dashboard.devices.getDeviceManagementInterface(serial)
                    wan1_settings = mgmt_settings.get('wan1', {})
                    
                    # Check if using static IP
                    using_static = wan1_settings.get('usingStaticIp', False)
                    
                    if using_static:
                        # Extract static IP configuration
                        static_config = {
                            'usingStaticIp': True,
                            'staticIp': wan1_settings.get('staticIp', ''),
                            'staticSubnetMask': wan1_settings.get('staticSubnetMask', ''),
                            'staticGatewayIp': wan1_settings.get('staticGatewayIp', ''),
                            'staticDns': wan1_settings.get('staticDns', []),
                            'vlan': wan1_settings.get('vlan')
                        }
                        
                        # Remove None values
                        static_config = {k: v for k, v in static_config.items() if v is not None}
                        
                        static_ip_settings = static_config
                        print(f"  Captured static IP settings from {device_name} ({serial}):")
                        print(f"    IP: {static_config.get('staticIp', 'N/A')}")
                        print(f"    Subnet: {static_config.get('staticSubnetMask', 'N/A')}")
                        print(f"    Gateway: {static_config.get('staticGatewayIp', 'N/A')}")
                        print(f"    DNS: {', '.join(static_config.get('staticDns', []))}")
                        if static_config.get('vlan'):
                            print(f"    VLAN: {static_config.get('vlan')}")
                        
                        # Use settings from first MX64 with static IP
                        break
                    else:
                        print(f"  {device_name} ({serial}): Using DHCP/Auto configuration")
                
                except Exception as e:
                    error_msg = f"Failed to capture settings from {device_name}: {e}"
                    print(f"  ERROR: {error_msg}")
                    if hasattr(self, 'results'):
                        self.results['errors'].append(error_msg)
        
        except Exception as e:
            error_msg = f"Failed to capture MX64 static IP settings: {e}"
            print(f"  ERROR: {error_msg}")
            if hasattr(self, 'results'):
                self.results['errors'].append(error_msg)
        
        if not static_ip_settings:
            print("  No static IP configurations found on MX64 devices")
        
        return static_ip_settings

    def remove_old_devices(self, network_id: str):
        """Step 3: Remove old devices"""
        print("Step 3: Removing old devices...")
        devices = self.dashboard.networks.getNetworkDevices(network_id)
        removed_devices = []
        failed_removals = []
        
        for device in devices:
            model = device.get('model', '').upper()
            device_name = device.get('name', device.get('serial', 'Unknown'))
            
            # Only remove specific old device models (MX64, not MX67/MX68)
            if any(old_model in model for old_model in ['MX64', 'MR33', 'MR36', 'CW9162']):
                # Additional check: skip if this is a new device we just added
                if device.get('serial') in getattr(self, 'new_device_serials', set()):
                    print(f"  Skipping removal of newly added device: {device_name} ({device['serial']})")
                    continue
                
                try:
                    self.dashboard.networks.removeNetworkDevices(network_id, serial=device['serial'])
                    removed_devices.append(f"{model} - {device_name} ({device['serial']})")
                    print(f"  Removed: {model} - {device_name} ({device['serial']})")
                    time.sleep(0.5)
                except Exception as e:
                    error_str = str(e)
                    
                    # Special handling for firmware upgrade errors
                    if "firmware upgrade" in error_str.lower():
                        error_msg = f"Failed to remove {device_name} ({device['serial']}): Device is currently undergoing firmware upgrade. Please wait for upgrade to complete and try again."
                        print(f"  FIRMWARE ERROR: {device_name} ({device['serial']}) - Device is undergoing firmware upgrade")
                        print(f"    You may need to manually remove this device after the firmware upgrade completes")
                    else:
                        error_msg = f"Failed to remove {device_name} ({device['serial']}): {e}"
                        print(f"  ERROR: {error_msg}")
                    
                    failed_removals.append(error_msg)
                    if hasattr(self, 'results'):
                        self.results['errors'].append(error_msg)
        
        if failed_removals:
            print(f"  WARNING: {len(failed_removals)} devices failed to be removed")
            print(f"  Check the errors section in the summary for details")
        
        return removed_devices

    def add_new_devices(self, network_id: str, devices: List[Dict]):
        """Step 4: Add and name new devices"""
        print("Step 4: Adding new devices...")
        current_devices = self.dashboard.networks.getNetworkDevices(network_id)
        current_serials = {d['serial'] for d in current_devices}
        added_devices = []
        
        # Track new device serials for removal logic
        self.new_device_serials = set()
        
        # Only process devices that have actual serial numbers (not existing device updates)
        new_devices = [d for d in devices if d.get('serial') and not d.get('update_existing')]
        
        for device in new_devices:
            serial = device['serial']
            self.new_device_serials.add(serial)
            
            if serial not in current_serials:
                try:
                    self.dashboard.networks.claimNetworkDevices(network_id, serials=[serial])
                    added_devices.append(f"{device['device_type']} ({serial})")
                    print(f"  Added: {serial}")
                    time.sleep(0.5)
                except Exception as e:
                    error_msg = f"Failed to add {serial}: {e}"
                    print(f"  ERROR: {error_msg}")
                    if hasattr(self, 'results'):
                        self.results['errors'].append(error_msg)
        
        # Update names for all new devices
        for device in new_devices:
            try:
                self.dashboard.devices.updateDevice(serial=device['serial'], name=device['name'])
                time.sleep(0.5)
            except Exception as e:
                error_msg = f"Failed to name {device['serial']}: {e}"
                print(f"  ERROR: {error_msg}")
                if hasattr(self, 'results'):
                    self.results['errors'].append(error_msg)
        
        return added_devices

    def convert_mx67_port2_to_wan(self, network_id: str):
        """Step 4.6: Convert port 2 to WAN on newly added MX67 devices"""
        print("Step 4.6: Converting port 2 to WAN on newly added MX67 devices...")
        converted_devices = []
        
        try:
            # Get all devices in the network
            devices = self.dashboard.networks.getNetworkDevices(network_id)
            
            # Find MX67 devices that were just added
            mx67_devices = []
            for device in devices:
                model = device.get('model', '').upper()
                serial = device.get('serial', '')
                
                # Check if it's an MX67 and was newly added
                if 'MX67' in model and serial in getattr(self, 'new_device_serials', set()):
                    mx67_devices.append(device)
            
            if not mx67_devices:
                print("  No newly added MX67 devices found")
                return converted_devices
            
            print(f"  Found {len(mx67_devices)} newly added MX67 device(s)")
            
            for device in mx67_devices:
                serial = device['serial']
                device_name = device.get('name', serial)
                
                try:
                    # Get current management interface settings to check if WAN2 is already enabled
                    current_settings = self.dashboard.devices.getDeviceManagementInterface(serial)
                    
                    # Check if WAN2 is already enabled
                    wan2_settings = current_settings.get('wan2', {})
                    wan2_enabled = wan2_settings.get('wanEnabled', 'not configured')
                    
                    if wan2_enabled == 'enabled':
                        print(f"  {device_name} ({serial}): Port 2 already configured as WAN2, skipping")
                        continue
                    
                    # Convert port 2 to WAN by enabling WAN2
                    update_payload = {
                        'wan2': {
                            'wanEnabled': 'enabled'
                        }
                    }
                    
                    self.dashboard.devices.updateDeviceManagementInterface(serial, **update_payload)
                    converted_devices.append(f"{device_name} ({serial})")
                    print(f"  Converted port 2 to WAN: {device_name} ({serial})")
                    time.sleep(1)  # Rate limiting and allow device to process change
                    
                except Exception as e:
                    error_msg = f"Port 2 WAN conversion failed for {device_name} ({serial}): {e}"
                    print(f"  ERROR: {error_msg}")
                    if hasattr(self, 'results'):
                        self.results['errors'].append(error_msg)
        
        except Exception as e:
            error_msg = f"Failed to convert MX67 port 2 to WAN: {e}"
            print(f"  ERROR: {error_msg}")
            if hasattr(self, 'results'):
                self.results['errors'].append(error_msg)
        
        return converted_devices

    def apply_static_ip_to_mx67(self, network_id: str, static_ip_settings: Dict):
        """Step 4.7: Apply static IP settings to newly added MX67 devices"""
        if not static_ip_settings:
            return []
        
        print("Step 4.7: Applying static IP settings to newly added MX67 devices...")
        configured_devices = []
        
        try:
            # Get all devices in the network
            devices = self.dashboard.networks.getNetworkDevices(network_id)
            
            # Find MX67 devices that were just added
            mx67_devices = []
            for device in devices:
                model = device.get('model', '').upper()
                serial = device.get('serial', '')
                
                # Check if it's an MX67 and was newly added
                if 'MX67' in model and serial in getattr(self, 'new_device_serials', set()):
                    mx67_devices.append(device)
            
            if not mx67_devices:
                print("  No newly added MX67 devices found")
                return configured_devices
            
            print(f"  Found {len(mx67_devices)} newly added MX67 device(s)")
            print(f"  Applying static IP configuration:")
            print(f"    IP: {static_ip_settings.get('staticIp')}")
            print(f"    Subnet: {static_ip_settings.get('staticSubnetMask')}")
            print(f"    Gateway: {static_ip_settings.get('staticGatewayIp')}")
            print(f"    DNS: {', '.join(static_ip_settings.get('staticDns', []))}")
            if static_ip_settings.get('vlan'):
                print(f"    VLAN: {static_ip_settings.get('vlan')}")
            
            for device in mx67_devices:
                serial = device['serial']
                device_name = device.get('name', serial)
                
                try:
                    # Apply static IP settings to WAN1
                    update_payload = {
                        'wan1': static_ip_settings
                    }
                    
                    self.dashboard.devices.updateDeviceManagementInterface(serial, **update_payload)
                    configured_devices.append(f"{device_name} ({serial})")
                    print(f"  Applied static IP settings to: {device_name} ({serial})")
                    time.sleep(1)  # Rate limiting
                    
                except Exception as e:
                    error_msg = f"Static IP configuration failed for {device_name} ({serial}): {e}"
                    print(f"  ERROR: {error_msg}")
                    if hasattr(self, 'results'):
                        self.results['errors'].append(error_msg)
        
        except Exception as e:
            error_msg = f"Failed to apply static IP settings to MX67 devices: {e}"
            print(f"  ERROR: {error_msg}")
            if hasattr(self, 'results'):
                self.results['errors'].append(error_msg)
        
        return configured_devices

    def update_existing_mt40s(self, network_id: str, mt40_name: str):
        """Step 4.8: Update names of existing MT40 devices"""
        print("Step 4.8: Updating existing MT40 devices...")
        updated_devices = []
        
        try:
            devices = self.dashboard.networks.getNetworkDevices(network_id)
            mt40_devices = [d for d in devices if d.get('model', '').upper().startswith('MT')]
            
            if not mt40_devices:
                print("  No existing MT40 devices found")
                return updated_devices
            
            for i, device in enumerate(mt40_devices):
                try:
                    # Use provided name or generate one if multiple MT40s exist
                    new_name = mt40_name if len(mt40_devices) == 1 else f"{mt40_name}-{i+1}"
                    self.dashboard.devices.updateDevice(serial=device['serial'], name=new_name)
                    updated_devices.append(f"{new_name} ({device['serial']})")
                    print(f"  Updated MT40 name: {new_name} ({device['serial']})")
                    time.sleep(0.5)
                except Exception as e:
                    error_msg = f"MT40 name update failed for {device['serial']}: {e}"
                    print(f"  ERROR: {error_msg}")
                    if hasattr(self, 'results'):
                        self.results['errors'].append(error_msg)
        
        except Exception as e:
            error_msg = f"Failed to find MT40 devices: {e}"
            print(f"  ERROR: {error_msg}")
            if hasattr(self, 'results'):
                self.results['errors'].append(error_msg)
        
        return updated_devices

    def update_device_addresses(self, network_id: str, address_info: Dict):
        """Step 5: Update device addresses for all devices in the network"""
        print("Step 5: Updating device addresses...")
        
        if not address_info or not all(address_info.get(key) for key in ['address', 'city', 'state']):
            print("  No valid address information provided, skipping address updates")
            return []
        
        # Format the full address
        full_address = f"{address_info['address']}, {address_info['city']}, {address_info['state']}"
        print(f"  Setting address for all devices: {full_address}")
        
        updated_devices = []
        
        try:
            # Get all devices in the network
            devices = self.dashboard.networks.getNetworkDevices(network_id)
            
            for device in devices:
                serial = device['serial']
                device_name = device.get('name', serial)
                
                try:
                    # Update the device address
                    self.dashboard.devices.updateDevice(serial=serial, address=full_address)
                    updated_devices.append(f"{device_name} ({serial})")
                    print(f"  Updated address for: {device_name} ({serial})")
                    time.sleep(0.5)  # Rate limiting
                except Exception as e:
                    error_msg = f"Address update failed for {device_name}: {e}"
                    print(f"  ERROR: {error_msg}")
                    if hasattr(self, 'results'):
                        self.results['errors'].append(error_msg)
        
        except Exception as e:
            error_msg = f"Failed to get network devices for address update: {e}"
            print(f"  ERROR: {error_msg}")
            if hasattr(self, 'results'):
                self.results['errors'].append(error_msg)
        
        return updated_devices

    def get_subnet_base(self, network_id: str) -> Optional[str]:
        """Helper method to determine the subnet base from existing assignments or VLAN settings"""
        try:
            # Try to get subnet from VLAN 1 settings
            vlan_settings = self.dashboard.appliance.getNetworkApplianceVlan(network_id, "1")
            subnet = vlan_settings.get('subnet', '')
            
            if subnet and '/' in subnet:
                # Extract base IP from subnet (e.g., "192.168.1.0/24" -> "192.168.1")
                base_ip = '.'.join(subnet.split('/')[0].split('.')[:-1])
                return base_ip
            
            # Fallback: check existing fixed IP assignments
            assignments = vlan_settings.get('fixedIpAssignments', {})
            if assignments:
                # Get first IP to determine subnet
                first_ip = list(assignments.values())[0].get('ip', '')
                if first_ip:
                    return '.'.join(first_ip.split('.')[:-1])
            
            # No fallback - return None if we can't determine subnet
            return None
            
        except Exception as e:
            print(f"  Warning: Could not determine subnet: {e}")
            return None

    def check_and_add_switch_assignments(self, network_id: str, switch_names: List[str] = None):
        """Step 5.5: Check if switches have fixed IP assignments, add them if missing"""
        print("Step 5.5: Checking and adding switch IP assignments...")
        created_switch_assignments = []
        
        try:
            # Get current network devices
            devices = self.dashboard.networks.getNetworkDevices(network_id)
            switches = [d for d in devices if d.get('model', '').upper().startswith(('MS120', 'MS130'))]
            
            if not switches:
                print("  No switches found in network")
                return created_switch_assignments
            
            # Get current VLAN 1 fixed IP assignments
            dhcp_settings = self.dashboard.appliance.getNetworkApplianceVlan(network_id, "1")
            assignments = dhcp_settings.get('fixedIpAssignments', {})
            
            # Get subnet base for IP assignment
            subnet_base = self.get_subnet_base(network_id)
            
            if not subnet_base:
                print("  Cannot determine network subnet - skipping switch IP assignments for safety")
                return created_switch_assignments
            
            print(f"  Detected subnet base: {subnet_base}")
            
            # Create mapping of current SW1/SW2 to new names
            switch_mapping = {}
            
            if switch_names:
                # Map current switches to new names based on SW1/SW2 designation
                for switch in switches:
                    current_name = switch.get('name', '').upper()
                    
                    # Determine if this is SW1 or SW2 based on current name
                    if 'SW1' in current_name:
                        # Find the new name that contains SW1
                        for new_name in switch_names:
                            if 'SW1' in new_name.upper():
                                switch_mapping[switch['serial']] = {
                                    'new_name': new_name,
                                    'ip_ending': '.93',
                                    'switch_designation': 'SW1'
                                }
                                print(f"  Mapped: Current SW1 ({current_name}) → {new_name}")
                                break
                    elif 'SW2' in current_name:
                        # Find the new name that contains SW2
                        for new_name in switch_names:
                            if 'SW2' in new_name.upper():
                                switch_mapping[switch['serial']] = {
                                    'new_name': new_name,
                                    'ip_ending': '.89',
                                    'switch_designation': 'SW2'
                                }
                                print(f"  Mapped: Current SW2 ({current_name}) → {new_name}")
                                break
            
            # If no mappings found, fall back to default behavior
            if not switch_mapping:
                print("  No SW1/SW2 patterns found in current names or spreadsheet names")
                print("  Using default assignment order")
                default_names = switch_names if switch_names else ["SW1", "SW2"]
                default_ips = [".93", ".89"]
                
                for i, switch in enumerate(switches[:2]):
                    switch_mapping[switch['serial']] = {
                        'new_name': default_names[i] if i < len(default_names) else f"SW{i+1}",
                        'ip_ending': default_ips[i],
                        'switch_designation': f'SW{i+1}'
                    }
            
            # Check which switches need IP assignments and process them
            switches_needing_updates = []
            for switch in switches:
                switch_mac = switch.get('mac', '').lower()
                switch_serial = switch.get('serial', '')
                has_assignment = False
                
                # Check if switch already has a fixed IP assignment
                for mac, assignment in assignments.items():
                    if mac.lower().replace(':', '') == switch_mac.replace(':', ''):
                        has_assignment = True
                        print(f"  Switch {switch.get('name', switch_serial)} already has IP assignment: {assignment.get('ip')}")
                        break
                
                # Add to update list if no IP assignment or if we have a name mapping
                if not has_assignment or switch_serial in switch_mapping:
                    switches_needing_updates.append(switch)
            
            if not switches_needing_updates:
                print("  All switches already configured properly")
                return created_switch_assignments
            
            # Process switches that need updates
            for switch in switches_needing_updates:
                switch_mac = switch.get('mac', '')
                switch_serial = switch.get('serial', '')
                
                if switch_serial not in switch_mapping:
                    continue
                
                mapping = switch_mapping[switch_serial]
                new_name = mapping['new_name']
                ip_ending = mapping['ip_ending']
                switch_ip = f"{subnet_base}{ip_ending}"
                
                try:
                    if switch_mac:
                        # Format MAC address with colons
                        raw_mac = switch_mac.lower().replace(':', '')
                        formatted_mac = ':'.join(raw_mac[j:j+2] for j in range(0, len(raw_mac), 2))
                        
                        # Create/update fixed IP assignment
                        assignments[formatted_mac] = {'ip': switch_ip, 'name': new_name}
                        
                        # Rename the actual device
                        try:
                            self.dashboard.devices.updateDevice(serial=switch_serial, name=new_name)
                            print(f"  Renamed switch device: {new_name} ({switch_serial})")
                            time.sleep(0.5)  # Rate limiting
                        except Exception as device_error:
                            error_msg = f"Switch device rename failed for {switch_serial}: {device_error}"
                            print(f"  WARNING: {error_msg}")
                            if hasattr(self, 'results'):
                                self.results['errors'].append(error_msg)
                        
                        created_switch_assignments.append(f"{new_name} -> {switch_ip} (MAC: {formatted_mac})")
                        print(f"  Updated switch assignment: {new_name} -> {switch_ip} (MAC: {formatted_mac})")
                
                except Exception as e:
                    error_msg = f"Switch assignment failed for {switch_serial}: {e}"
                    print(f"  ERROR: {error_msg}")
                    if hasattr(self, 'results'):
                        self.results['errors'].append(error_msg)
            
            # Update VLAN with new/updated assignments
            if created_switch_assignments:
                self.dashboard.appliance.updateNetworkApplianceVlan(network_id, "1", fixedIpAssignments=assignments)
                time.sleep(1)
                print(f"  Updated {len(created_switch_assignments)} switch assignments")
            
        except Exception as e:
            error_msg = f"Failed to check/add switch assignments: {e}"
            print(f"  ERROR: {error_msg}")
            if hasattr(self, 'results'):
                self.results['errors'].append(error_msg)
        
        return created_switch_assignments

    def create_ap_assignments(self, network_id: str, ip_assignments: List[Dict]):
        """Step 6: Create AP IP assignments"""
        if not ip_assignments:
            return []
        
        print("Step 6: Creating AP assignments...")
        created_assignments = []
        try:
            # Only work with VLAN 1
            dhcp_settings = self.dashboard.appliance.getNetworkApplianceVlan(network_id, "1")
            assignments = dhcp_settings.get('fixedIpAssignments', {})
            
            for i, assignment in enumerate(ip_assignments, 1):
                try:
                    device_info = self.dashboard.devices.getDevice(assignment['serial'])
                    if device_info and 'mac' in device_info:
                        # Format MAC address with colons (bc:33:40:49:78:40)
                        raw_mac = device_info['mac'].lower().replace(':', '')
                        formatted_mac = ':'.join(raw_mac[j:j+2] for j in range(0, len(raw_mac), 2))
                        
                        # Use AP1, AP2, etc. for naming (changed from CW1, CW2)
                        ap_name = f"AP{i}"
                        assignments[formatted_mac] = {'ip': assignment['ip'], 'name': ap_name}
                        created_assignments.append(f"{ap_name} -> {assignment['ip']} (MAC: {formatted_mac})")
                        print(f"  Added AP: {ap_name} -> {assignment['ip']} (MAC: {formatted_mac})")
                except Exception as e:
                    error_msg = f"Failed AP assignment for {assignment['serial']}: {e}"
                    print(f"  ERROR: {error_msg}")
                    if hasattr(self, 'results'):
                        self.results['errors'].append(error_msg)
            
            self.dashboard.appliance.updateNetworkApplianceVlan(network_id, "1", fixedIpAssignments=assignments)
            time.sleep(1)
        except Exception as e:
            error_msg = f"Error processing VLAN 1 for AP assignments: {e}"
            print(f"  ERROR: {error_msg}")
            if hasattr(self, 'results'):
                self.results['errors'].append(error_msg)
        
        return created_assignments

    def get_devices_to_remove(self, network_id: str):
        """Get list of devices that will be removed"""
        devices = self.dashboard.networks.getNetworkDevices(network_id)
        devices_to_remove = []
        
        for device in devices:
            model = device.get('model', '').upper()
            device_name = device.get('name', device.get('serial', 'Unknown'))
            
            # Only show MX64s and other old models (not MX67/MX68)
            if any(old_model in model for old_model in ['MX64', 'MR33', 'MR36', 'CW9162']):
                devices_to_remove.append(f"{model} - {device_name} ({device['serial']})")
        
        return devices_to_remove

    def complete_refresh(self, network_id: str, devices: List[Dict], address_info: Dict = None, ip_assignments: List[Dict] = None, switch_names: List[str] = None):
        """Execute complete 8-step refresh (switches stay in place, preserve static IP settings)"""
        print(f"Starting network refresh for {network_id}")
        
        # Track results for summary
        self.results = {
            'assignments_cleared': 0,
            'iboot_ranges_removed': 0,
            'static_ip_captured': False,
            'devices_removed': [],
            'devices_added': [],
            'mx67_wan_conversions': [],
            'mx67_static_ip_applied': [],  # New tracking for static IP preservation
            'mt40_updates': [],
            'addresses_updated': [],
            'switch_assignments_created': [],
            'ap_assignments_created': [],
            'errors': []
        }
        
        # Step 1: Clear non-switch assignments (switches stay in place)
        self.results['assignments_cleared'] = self.clear_non_switch_assignments(network_id)
        print("Waiting for dashboard sync after clearing assignments...")
        time.sleep(10)
        
        # Step 2: Remove iBoot ranges
        self.results['iboot_ranges_removed'] = self.remove_iboot_ranges(network_id)
        print("Waiting for dashboard sync after removing iBoot ranges...")
        time.sleep(10)
        
        # Step 2.5: Capture static IP settings from MX64 devices before removal
        static_ip_settings = self.capture_mx64_static_ip_settings(network_id)
        self.results['static_ip_captured'] = bool(static_ip_settings)
        
        # Step 3: Remove old devices
        removed_devices = self.remove_old_devices(network_id)
        self.results['devices_removed'] = removed_devices
        print("Waiting for dashboard sync after removing devices...")
        time.sleep(10)
        
        # Step 4: Add new devices
        added_devices = self.add_new_devices(network_id, devices)
        self.results['devices_added'] = added_devices
        print("Waiting for dashboard sync after adding devices...")
        time.sleep(10)
        
        # Step 4.6: Convert MX67 port 2 to WAN (only for newly added MX67s)
        mx67_conversions = self.convert_mx67_port2_to_wan(network_id)
        self.results['mx67_wan_conversions'] = mx67_conversions
        if mx67_conversions:
            print("Waiting for dashboard sync after MX67 WAN conversions...")
            time.sleep(10)
        
        # Step 4.7: Apply static IP settings to MX67 devices (if captured from MX64)
        mx67_static_ip_applied = self.apply_static_ip_to_mx67(network_id, static_ip_settings)
        self.results['mx67_static_ip_applied'] = mx67_static_ip_applied
        if mx67_static_ip_applied:
            print("Waiting for dashboard sync after applying static IP settings...")
            time.sleep(10)
        
        # Step 4.8: Update existing MT40 names if specified (was Step 4.7)
        mt40_updates = [d for d in devices if d.get('update_existing')]
        if mt40_updates:
            mt40_name = mt40_updates[0]['name']  # Get the name from spreadsheet
            updated_mt40s = self.update_existing_mt40s(network_id, mt40_name)
            self.results['mt40_updates'] = updated_mt40s
            print("Waiting for dashboard sync after updating MT40 names...")
            time.sleep(5)
        
        # Step 5: Update device addresses
        if address_info:
            updated_addresses = self.update_device_addresses(network_id, address_info)
            self.results['addresses_updated'] = updated_addresses
            print("Waiting for dashboard sync after updating addresses...")
            time.sleep(10)
        
        # Step 5.5: Check and add switch IP assignments if missing
        switch_assignments = self.check_and_add_switch_assignments(network_id, switch_names)
        self.results['switch_assignments_created'] = switch_assignments
        print("Waiting for dashboard sync after creating switch assignments...")
        time.sleep(5)
        
        # Step 6: Create AP assignments
        if ip_assignments:
            self.results['ap_assignments_created'] = self.create_ap_assignments(network_id, ip_assignments)
            print("Waiting for dashboard sync after creating AP assignments...")
            time.sleep(10)
        
        print("Network refresh complete!")
        return self.results

def load_store_data(file_path: str) -> Dict:
    """Load store data from spreadsheet"""
    df = pd.read_excel(file_path) if file_path.endswith('.xlsx') else pd.read_csv(file_path)
    stores = {}
    
    for _, row in df.iterrows():
        store_num = str(row.get('storeNumber', '')).strip()
        network_id = str(row.get('NetworkID', '')).strip()
        
        if not store_num or not network_id:
            continue
        
        # Extract address information
        address_info = {
            'address': str(row.get('Address', '')).strip() if pd.notna(row.get('Address')) else '',
            'city': str(row.get('City', '')).strip() if pd.notna(row.get('City')) else '',
            'state': str(row.get('State', '')).strip() if pd.notna(row.get('State')) else ''
        }
        
        # Only include address info if all components are present
        if not all(address_info.values()):
            address_info = None
        
        # Extract switch names
        switch_names = []
        switch_name_a = str(row.get('SW-A Name', '')).strip() if pd.notna(row.get('SW-A Name')) else ''
        switch_name_b = str(row.get('SW-B Name', '')).strip() if pd.notna(row.get('SW-B Name')) else ''
        
        if switch_name_a and switch_name_a != 'nan':
            switch_names.append(switch_name_a)
        if switch_name_b and switch_name_b != 'nan':
            switch_names.append(switch_name_b)
        
        devices = []
        device_pairs = [('mx_a', 'MX-A SN', 'MX-A Name'), ('mx_b', 'MX-B SN', 'MX-B Name'),
                       ('cw_a', 'CW-A SN', 'CW-A Name'), ('cw_b', 'CW-B SN', 'CW-B Name'),
                       ('mt40', 'MT40 SN', 'MT40 Name')]
        
        for device_type, sn_col, name_col in device_pairs:
            serial = str(row.get(sn_col, '')).strip()
            name = str(row.get(name_col, '')).strip()
            
            if device_type == 'mt40':
                # Special handling for MT40: create entry even if no serial (for existing device updates)
                if serial and serial != 'nan':
                    # New MT40 to be added
                    device = {
                        'serial': serial, 
                        'name': name or f"{DEVICE_TYPES[device_type]}-{store_num}", 
                        'device_type': DEVICE_TYPES[device_type]
                    }
                    devices.append(device)
                elif name and name != 'nan':
                    # Existing MT40 to be renamed (no serial provided)
                    device = {
                        'serial': None,  # Indicates existing device
                        'name': name, 
                        'device_type': DEVICE_TYPES[device_type],
                        'update_existing': True
                    }
                    devices.append(device)
            else:
                # Handle other devices as before
                if serial and serial != 'nan':
                    device = {'serial': serial, 'name': name or f"{DEVICE_TYPES[device_type]}-{store_num}", 'device_type': DEVICE_TYPES[device_type]}
                    
                    if device_type in ['cw_a', 'cw_b']:
                        ip = str(row.get(f'{device_type.upper().replace("_", "-")} IP', '')).strip()
                        if ip and ip != 'nan':
                            device['ip_address'] = ip
                    
                    devices.append(device)
        
        if devices:
            store_data = {
                'store_number': store_num, 
                'network_id': network_id, 
                'devices': devices
            }
            if address_info:
                store_data['address_info'] = address_info
            if switch_names:
                store_data['switch_names'] = switch_names
            stores[store_num] = store_data
    
    return stores

def print_terminal_summary(store_num: str, network_id: str, results: Dict):
    """Print summary to terminal"""
    print("\n" + "="*80)
    print(f"NETWORK REFRESH SUMMARY - STORE {store_num}")
    print("="*80)
    print(f"Network ID: {network_id}")
    print(f"Timestamp: {time.strftime('%Y-%m-%d %H:%M:%S')}")
    print("\nSTEP RESULTS:")
    print(f"  ✓ Non-switch assignments cleared: {results['assignments_cleared']}")
    print(f"  ✓ iBoot ranges removed: {results['iboot_ranges_removed']}")
    print(f"  ✓ Static IP settings captured: {'Yes' if results['static_ip_captured'] else 'No'}")
    print(f"  ✓ Old devices removed: {len(results['devices_removed'])}")
    print(f"  ✓ New devices added: {len(results['devices_added'])}")
    print(f"  ✓ MX67 port 2 converted to WAN: {len(results['mx67_wan_conversions'])}")
    print(f"  ✓ MX67 static IP settings applied: {len(results['mx67_static_ip_applied'])}")
    print(f"  ✓ Existing MT40s updated: {len(results['mt40_updates'])}")
    print(f"  ✓ Device addresses updated: {len(results['addresses_updated'])}")
    print(f"  ✓ Switch assignments created: {len(results['switch_assignments_created'])}")
    print(f"  ✓ AP assignments created: {len(results['ap_assignments_created'])}")
    
    if results['errors']:
        print(f"\n⚠ ERRORS ENCOUNTERED: {len(results['errors'])}")
        for error in results['errors']:
            print(f"    - {error}")
    else:
        print(f"\n🎉 SUCCESS: Network refresh completed without errors!")
    
    print("="*80)

def create_summary_file(store_num: str, network_id: str, results: Dict, devices: List[Dict], address_info: Dict = None):
    """Create detailed summary text file"""
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    filename = f"network_refresh_summary_store_{store_num}_{timestamp}.txt"
    
    with open(filename, 'w') as f:
        f.write("MERAKI NETWORK REFRESH SUMMARY\n")
        f.write("=" * 50 + "\n\n")
        f.write(f"Store Number: {store_num}\n")
        f.write(f"Network ID: {network_id}\n")
        f.write(f"Date/Time: {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"Script Version: Network Refresh Complete Tool v2.0\n\n")
        
        f.write("WORKFLOW EXECUTED:\n")
        f.write("1. Clear non-switch fixed IP assignments (switches preserved)\n")
        f.write("2. Remove iBoot reserved IP ranges\n")
        f.write("2.5. Capture static IP settings from MX64 devices\n")
        f.write("3. Remove old devices (MX64, MR33, MR36, CW9162 - MX67/MX68 devices are preserved)\n")
        f.write("4. Add new devices and update names\n")
        f.write("4.6. Convert port 2 to WAN on newly added MX67 devices\n")
        f.write("4.7. Apply captured static IP settings to MX67 devices\n")
        f.write("4.8. Update existing MT40 device names\n")
        f.write("5. Update device addresses\n")
        f.write("5.5. Check and add switch IP assignments if missing (with intelligent SW1/SW2 mapping)\n")
        f.write("6. Create fixed IP assignments for new access points\n\n")
        
        f.write("RESULTS SUMMARY:\n")
        f.write("-" * 20 + "\n")
        f.write(f"Non-switch assignments cleared: {results['assignments_cleared']}\n")
        f.write(f"iBoot ranges removed: {results['iboot_ranges_removed']}\n")
        f.write(f"Static IP settings captured: {'Yes' if results['static_ip_captured'] else 'No'}\n")
        f.write(f"Old devices removed: {len(results['devices_removed'])}\n")
        f.write(f"New devices added: {len(results['devices_added'])}\n")
        f.write(f"MX67 port 2 converted to WAN: {len(results['mx67_wan_conversions'])}\n")
        f.write(f"MX67 static IP settings applied: {len(results['mx67_static_ip_applied'])}\n")
        f.write(f"Existing MT40s updated: {len(results['mt40_updates'])}\n")
        f.write(f"Device addresses updated: {len(results['addresses_updated'])}\n")
        f.write(f"Switch assignments created: {len(results['switch_assignments_created'])}\n")
        f.write(f"AP assignments created: {len(results['ap_assignments_created'])}\n\n")
        
        if address_info:
            f.write("ADDRESS INFORMATION:\n")
            f.write("-" * 19 + "\n")
            f.write(f"Address: {address_info['address']}, {address_info['city']}, {address_info['state']}\n\n")
        
        if results['devices_removed']:
            f.write("DEVICES REMOVED:\n")
            f.write("-" * 15 + "\n")
            for device in results['devices_removed']:
                f.write(f"  - {device}\n")
            f.write("\n")
        
        if results['devices_added']:
            f.write("DEVICES ADDED:\n")
            f.write("-" * 13 + "\n")
            for device in results['devices_added']:
                f.write(f"  - {device}\n")
            f.write("\n")
        
        if results['mx67_wan_conversions']:
            f.write("MX67 PORT 2 WAN CONVERSIONS:\n")
            f.write("-" * 28 + "\n")
            for device in results['mx67_wan_conversions']:
                f.write(f"  - {device}\n")
            f.write("\n")
        
        if results['mx67_static_ip_applied']:
            f.write("MX67 STATIC IP SETTINGS APPLIED:\n")
            f.write("-" * 32 + "\n")
            for device in results['mx67_static_ip_applied']:
                f.write(f"  - {device}\n")
            f.write("\n")
        
        if results['mt40_updates']:
            f.write("EXISTING MT40 DEVICES UPDATED:\n")
            f.write("-" * 30 + "\n")
            for device in results['mt40_updates']:
                f.write(f"  - {device}\n")
            f.write("\n")
        
        if results['addresses_updated']:
            f.write("DEVICE ADDRESSES UPDATED:\n")
            f.write("-" * 24 + "\n")
            for device in results['addresses_updated']:
                f.write(f"  - {device}\n")
            f.write("\n")
        
        if results['switch_assignments_created']:
            f.write("SWITCH FIXED IP ASSIGNMENTS CREATED:\n")
            f.write("-" * 35 + "\n")
            for assignment in results['switch_assignments_created']:
                f.write(f"  - {assignment}\n")
            f.write("\n")
        
        if results['ap_assignments_created']:
            f.write("AP FIXED IP ASSIGNMENTS CREATED:\n")
            f.write("-" * 32 + "\n")
            for assignment in results['ap_assignments_created']:
                f.write(f"  - {assignment}\n")
            f.write("\n")
        
        f.write("NEW DEVICE DETAILS:\n")
        f.write("-" * 18 + "\n")
        for device in devices:
            if not device.get('update_existing'):  # Only show devices that were actually added
                f.write(f"  Device: {device['name']}\n")
                f.write(f"    Serial: {device.get('serial', 'N/A')}\n")
                f.write(f"    Type: {device['device_type']}\n")
                if 'ip_address' in device:
                    f.write(f"    IP: {device['ip_address']}\n")
                # Add note for MX67 devices about WAN conversion and static IP
                if 'MX67' in device['device_type']:
                    f.write(f"    Note: Port 2 converted to WAN")
                    if results['mx67_static_ip_applied']:
                        f.write(f", Static IP settings applied from MX64")
                    f.write(f"\n")
                f.write("\n")
        
        if results['errors']:
            f.write("ERRORS ENCOUNTERED:\n")
            f.write("-" * 18 + "\n")
            for error in results['errors']:
                f.write(f"  - {error}\n")
            f.write("\n")
        else:
            f.write("STATUS: SUCCESS - No errors encountered\n\n")
        
        f.write("END OF REPORT\n")
    
    return filename

def main():
    """Main execution"""
    if not API_KEY:
        print("ERROR: API_KEY not found in .env file")
        print("Please create a .env file with your API_KEY, ORG_ID, and DATA_FILE_PATH")
        print("Example .env file contents:")
        print("API_KEY=your_api_key_here")
        print("ORG_ID=your_org_id_here")
        print("DATA_FILE_PATH=C:\\path\\to\\your\\Network Refresh Tool - Data Sheet.xlsx")
        return
    
    if not ORG_ID:
        print("WARNING: ORG_ID not found in .env file")
        print("Will use the first organization found in your account")
    
    # Load data - try .env file path first, then fall back to file dialog
    file_path = None
    
    if DATA_FILE_PATH and os.path.exists(DATA_FILE_PATH):
        file_path = DATA_FILE_PATH
        print(f"Using data file from .env: {DATA_FILE_PATH}")
    else:
        if DATA_FILE_PATH:
            print(f"WARNING: DATA_FILE_PATH specified in .env but file not found: {DATA_FILE_PATH}")
        print("Opening file dialog to select spreadsheet...")
        
        # Try to create a root window for the file dialog
        try:
            root = tk.Tk()
            root.withdraw()  # Hide the root window
            file_path = filedialog.askopenfilename(
                title="Select Network Refresh Data Spreadsheet", 
                filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv"), ("All files", "*.*")]
            )
            root.destroy()
        except Exception as e:
            print(f"Error opening file dialog: {e}")
            print("Please ensure DATA_FILE_PATH is correctly set in your .env file")
            return
    
    if not file_path:
        print("No file selected. Exiting.")
        return
    
    stores = load_store_data(file_path)
    if not stores:
        print("No stores found")
        return
    
    # Select store
    print(f"Available stores: {list(stores.keys())}")
    store_num = input("Enter store number: ").strip()
    
    if store_num not in stores:
        print("Store not found")
        return
    
    store_data = stores[store_num]
    devices = store_data['devices']
    address_info = store_data.get('address_info')
    switch_names = store_data.get('switch_names', [])
    ip_assignments = [{'name': d['name'], 'serial': d['serial'], 'ip': d['ip_address']} 
                     for d in devices if 'ip_address' in d and d.get('serial')]
    
    print(f"Processing {len(devices)} device entries for store {store_num}")
    
    # Show what will be removed and added
    manager = NetworkRefreshManager(API_KEY, ORG_ID)
    devices_to_remove = manager.get_devices_to_remove(store_data['network_id'])
    devices_to_add = [f"{d['device_type']} ({d['serial']})" for d in devices if d.get('serial') and not d.get('update_existing')]
    mt40_updates = [d for d in devices if d.get('update_existing')]
    
    print("\n" + "="*60)
    print("DEVICES TO BE REMOVED:")
    if devices_to_remove:
        for device in devices_to_remove:
            print(f"  - {device}")
    else:
        print("  - None")
    
    print("\nDEVICES TO BE ADDED:")
    if devices_to_add:
        for device in devices_to_add:
            print(f"  - {device}")
    else:
        print("  - None")
    
    if mt40_updates:
        print(f"\nEXISTING MT40 DEVICES TO BE UPDATED:")
        for mt40 in mt40_updates:
            print(f"  - Name will be set to: {mt40['name']}")
    
    if address_info:
        print(f"\nADDRESS TO BE SET FOR ALL DEVICES:")
        print(f"  - {address_info['address']}, {address_info['city']}, {address_info['state']}")
    else:
        print("\nNOTE: No address information found in spreadsheet")
    
    if ip_assignments:
        print(f"\nIP ASSIGNMENTS TO BE CREATED:")
        for i, assignment in enumerate(ip_assignments, 1):
            print(f"  - AP{i}: {assignment['ip']} ({assignment['serial']})")
    
    print(f"\nSWITCH NAMES FROM SPREADSHEET:")
    if switch_names:
        for i, name in enumerate(switch_names):
            ip_ending = ".93" if i == 0 else ".89"
            print(f"  - {name}: IP ending in {ip_ending}")
    else:
        print("  - No switch names found in spreadsheet, will use SW1/SW2 as defaults")
    
    print("\nSTATIC IP PRESERVATION:")
    print("  - Script will capture static IP settings from MX64 devices before removal")
    print("  - If static IP settings are found, they will be applied to new MX67 devices")
    print("  - MX67 port 2 will be converted to WAN on newly added devices")
    print("\nNOTE: Switch IP assignments will be checked and added if missing")
    print("Switch assignments only created if network subnet can be safely determined.")
    print("Existing switch assignments will be preserved.")
    print("="*60)
    
    # Confirm
    if input("Proceed? (y/n): ").lower() != 'y':
        return
    
    # Execute refresh
    results = manager.complete_refresh(store_data['network_id'], devices, address_info, ip_assignments, switch_names)
    
    # Print terminal summary
    print_terminal_summary(store_num, store_data['network_id'], results)
    
    # Create detailed text file summary
    summary_filename = create_summary_file(store_num, store_data['network_id'], results, devices, address_info)
    
    print(f"\nDetailed summary saved to: {summary_filename}")

if __name__ == "__main__":
    main()