#!/usr/bin/env python3
"""
MX67 Staging Network Management Tool

This tool manages MX67 devices on staging networks for firmware sync purposes.
Devices are temporarily added to networks for firmware updates, then removed.
"""

import requests
import json
import time
import argparse
import sys
import os
from typing import List, Dict, Optional
from dataclasses import dataclass
from dotenv import load_dotenv

# Static list of staging networks
STAGING_NETWORKS = {
    "CAN Network Refresh Staging 01": "N_769552586326941404",
    "CAN Network Refresh Staging 02": "N_769552586326941409", 
    "CAN Network Refresh Staging 03": "N_769552586326941410",
    "CAN Network Refresh Staging 04": "N_769552586326941411",
    "CAN Network Refresh Staging 05": "N_769552586326941405",
    "CAN Network Refresh Staging 06": "N_769552586326941406",
    "CAN Network Refresh Staging 07": "N_769552586326941407",
    "CAN Network Refresh Staging 08": "N_769552586326941408",
    "CAN Network Refresh Staging 09": "N_769552586326941412",
    "CAN Network Refresh Staging 10": "N_769552586326941413",
}


@dataclass
class Device:
    serial: str
    model: str
    name: Optional[str] = None


class MerakiStagingManager:
    def __init__(self, api_key: str, org_id: str, base_url: str = "https://api.meraki.com/api/v1"):
        self.api_key = api_key
        self.org_id = org_id
        self.base_url = base_url
        self.headers = {
            "X-Cisco-Meraki-API-Key": api_key,
            "Content-Type": "application/json"
        }
    
    def _make_request(self, method: str, endpoint: str, data: Dict = None) -> Dict:
        """Make API request with error handling"""
        url = f"{self.base_url}{endpoint}"
        
        try:
            if method.upper() == "GET":
                response = requests.get(url, headers=self.headers)
            elif method.upper() == "POST":
                response = requests.post(url, headers=self.headers, json=data)
            elif method.upper() == "DELETE":
                response = requests.delete(url, headers=self.headers)
            elif method.upper() == "PUT":
                response = requests.put(url, headers=self.headers, json=data)
            else:
                raise ValueError(f"Unsupported HTTP method: {method}")
            
            response.raise_for_status()
            return response.json() if response.content else {}
            
        except requests.exceptions.RequestException as e:
            print(f"API request failed: {e}")
            if hasattr(e, 'response') and e.response is not None:
                print(f"Response: {e.response.text}")
            raise
    
    def get_organizations(self) -> List[Dict]:
        """Get all organizations"""
        return self._make_request("GET", "/organizations")
    
    def get_networks(self, org_id: str, tags: List[str] = None) -> List[Dict]:
        """Get networks, optionally filtered by tags"""
        networks = self._make_request("GET", f"/organizations/{org_id}/networks")
        
        if tags:
            # Filter networks by tags
            filtered_networks = []
            for network in networks:
                network_tags = network.get('tags', [])
                if any(tag in network_tags for tag in tags):
                    filtered_networks.append(network)
            return filtered_networks
        
        return networks
    
    def get_network_devices(self, network_id: str) -> List[Dict]:
        """Get all devices in a network"""
        return self._make_request("GET", f"/networks/{network_id}/devices")
    
    def get_organization_devices(self, org_id: str) -> List[Dict]:
        """Get all devices in organization"""
        return self._make_request("GET", f"/organizations/{org_id}/devices")
    
    def claim_device(self, network_id: str, serial: str) -> Dict:
        """Claim a device to a network"""
        data = {"serials": [serial]}
        return self._make_request("POST", f"/networks/{network_id}/devices/claim", data)
    
    def remove_device(self, network_id: str, serial: str) -> Dict:
        """Remove a device from a network"""
        return self._make_request("POST", f"/networks/{network_id}/devices/{serial}/remove")
    
    def update_device(self, network_id: str, serial: str, **kwargs) -> Dict:
        """Update device settings"""
        return self._make_request("PUT", f"/networks/{network_id}/devices/{serial}", kwargs)
    
    def get_mx67_devices(self, available_only: bool = True) -> List[Device]:
        """Get MX67 devices, optionally only unclaimed ones"""
        devices = self.get_organization_devices(self.org_id)
        mx67_devices = []
        
        for device in devices:
            if device.get('model', '').startswith('MX67'):
                # If available_only, skip devices that are already in networks
                if available_only and device.get('networkId'):
                    continue
                
                mx67_devices.append(Device(
                    serial=device['serial'],
                    model=device['model'],
                    name=device.get('name')
                ))
        
        return mx67_devices
    
    def add_mx67_to_network(self, network_id: str, serial: str, device_name: str = None) -> bool:
        """Add MX67 device to staging network"""
        try:
            print(f"Adding MX67 {serial} to network {network_id}...")
            
            # Claim the device
            self.claim_device(network_id, serial)
            
            # Update device name if provided
            if device_name:
                self.update_device(network_id, serial, name=device_name)
            
            print(f"Successfully added MX67 {serial}")
            return True
            
        except Exception as e:
            print(f"Failed to add MX67 {serial}: {e}")
            return False
    
    def remove_mx67_from_network(self, network_id: str, serial: str) -> bool:
        """Remove MX67 device from staging network"""
        try:
            print(f"Removing MX67 {serial} from network {network_id}...")
            
            self.remove_device(network_id, serial)
            
            print(f"Successfully removed MX67 {serial}")
            return True
            
        except Exception as e:
            print(f"Failed to remove MX67 {serial}: {e}")
            return False
    
    def remove_mx67_batch(self, network_id: str, mx67_serials: List[str]) -> Dict:
        """
        Remove multiple MX67s from network after firmware sync
        
        Args:
            network_id: Target staging network
            mx67_serials: List of MX67 serial numbers
        """
        results = {
            'removed': [],
            'failed_to_remove': []
        }
        
        print(f"Removing {len(mx67_serials)} MX67 devices from staging network")
        
        # Remove devices from network
        for serial in mx67_serials:
            if self.remove_mx67_from_network(network_id, serial):
                results['removed'].append(serial)
            else:
                results['failed_to_remove'].append(serial)
        
        return results
    
    def check_staging_network_capacity(self) -> Dict:
        """Check current device count in each staging network"""
        network_status = {}
        
        for network_name, network_id in STAGING_NETWORKS.items():
            try:
                devices = self.get_network_devices(network_id)
                mx67_count = len([d for d in devices if d.get('model', '').startswith('MX67')])
                network_status[network_name] = {
                    'network_id': network_id,
                    'mx67_count': mx67_count,
                    'available_slots': max(0, 2 - mx67_count),
                    'devices': [d for d in devices if d.get('model', '').startswith('MX67')]
                }
            except Exception as e:
                network_status[network_name] = {
                    'network_id': network_id,
                    'mx67_count': 0,
                    'available_slots': 0,
                    'error': str(e),
                    'devices': []
                }
        
        return network_status
    
    def smart_batch_add(self, mx67_serials: List[str]) -> Dict:
        """
        Intelligently distribute MX67s across staging networks (max 2 per network)
        
        Args:
            mx67_serials: List of MX67 serial numbers (up to 20)
        """
        if len(mx67_serials) > 20:
            raise ValueError("Maximum of 20 devices allowed per batch")
        
        results = {
            'added': {},  # network_name: [serials]
            'failed_to_add': [],
            'network_assignments': {},  # serial: network_name
            'networks_with_existing_devices': {}
        }
        
        print(f"Processing {len(mx67_serials)} MX67 devices for staging network assignment")
        print("Checking staging network capacity...")
        
        # Check current network status
        network_status = self.check_staging_network_capacity()
        
        # Alert about networks with existing devices
        networks_with_devices = {name: info for name, info in network_status.items() 
                               if info['mx67_count'] > 0}
        
        if networks_with_devices:
            print("\n⚠️  WARNING: Found existing MX67 devices in staging networks:")
            for network_name, info in networks_with_devices.items():
                if 'error' not in info:
                    print(f"   {network_name}: {info['mx67_count']}/2 slots used")
                    for device in info['devices']:
                        print(f"     - {device['serial']} ({device.get('name', 'No name')})")
            
            results['networks_with_existing_devices'] = networks_with_devices
            
            response = input("\nDo you want to continue? Existing devices should be removed first. (y/N): ")
            if response.lower() not in ['y', 'yes']:
                print("Operation cancelled. Please remove existing devices first.")
                return results
        
        # Sort networks by available capacity (most available first)
        available_networks = [(name, info) for name, info in network_status.items() 
                            if info.get('available_slots', 0) > 0 and 'error' not in info]
        available_networks.sort(key=lambda x: x[1]['available_slots'], reverse=True)
        
        total_capacity = sum(info['available_slots'] for _, info in available_networks)
        
        if len(mx67_serials) > total_capacity:
            print(f"\n❌ Error: Requesting {len(mx67_serials)} devices but only {total_capacity} slots available")
            print("Available capacity:")
            for network_name, info in available_networks:
                print(f"   {network_name}: {info['available_slots']}/2 available")
            return results
        
        print(f"\n✓ Sufficient capacity found. Distributing devices across networks...")
        
        # Distribute devices across networks
        device_queue = mx67_serials.copy()
        
        while device_queue and available_networks:
            for network_name, network_info in available_networks[:]:
                if not device_queue:
                    break
                
                if network_info['available_slots'] > 0:
                    serial = device_queue.pop(0)
                    network_id = network_info['network_id']
                    
                    if self.add_mx67_to_network(network_id, serial):
                        if network_name not in results['added']:
                            results['added'][network_name] = []
                        results['added'][network_name].append(serial)
                        results['network_assignments'][serial] = network_name
                        network_info['available_slots'] -= 1
                    else:
                        results['failed_to_add'].append(serial)
                
                # Remove network from available list if full
                if network_info['available_slots'] == 0:
                    available_networks.remove((network_name, network_info))
        
        # Generate removal commands for each network
        removal_commands = []
        for network_name, serials in results['added'].items():
            if serials:
                removal_commands.append(
                    f'python mx67_tool.py batch-remove --network "{network_name}" --serials {" ".join(serials)}'
                )
        
        if results['added']:
            total_added = sum(len(serials) for serials in results['added'].values())
            print(f"\n✓ Successfully added {total_added} devices across {len(results['added'])} networks")
            print("Devices will now sync firmware and configuration automatically.")
            print("\nTo remove devices after firmware sync is complete, run these commands:")
            for cmd in removal_commands:
                print(f"  {cmd}")
        
        return results

    def remove_all_mx67s_from_staging(self) -> Dict:
        """
        Remove all MX67 devices from all staging networks
        """
        results = {
            'networks_processed': {},
            'total_removed': 0,
            'total_failed': 0
        }
        
        print("🧹 Removing ALL MX67 devices from ALL staging networks...")
        print("Checking all staging networks for MX67 devices...")
        
        network_status = self.check_staging_network_capacity()
        networks_with_devices = {name: info for name, info in network_status.items() 
                               if info.get('mx67_count', 0) > 0 and 'error' not in info}
        
        if not networks_with_devices:
            print("✓ No MX67 devices found in any staging networks.")
            return results
        
        print(f"\nFound MX67 devices in {len(networks_with_devices)} staging networks:")
        total_devices = 0
        for network_name, info in networks_with_devices.items():
            device_count = len(info['devices'])
            total_devices += device_count
            print(f"  📍 {network_name}: {device_count} devices")
            for device in info['devices']:
                print(f"     - {device['serial']} ({device.get('name', 'No name')})")
        
        print(f"\nTotal MX67 devices to remove: {total_devices}")
        
        # Confirm with user
        confirm = input(f"\n⚠️  Are you sure you want to remove ALL {total_devices} MX67 devices from ALL staging networks? (yes/no): ")
        if confirm.lower() not in ['yes', 'y']:
            print("Operation cancelled.")
            return results
        
        print("\nProceeding with removal...")
        
        # Remove devices from each network
        for network_name, info in networks_with_devices.items():
            network_id = info['network_id']
            serials_to_remove = [device['serial'] for device in info['devices']]
            
            print(f"\n📍 Processing {network_name}...")
            removal_results = self.remove_mx67_batch(network_id, serials_to_remove)
            
            results['networks_processed'][network_name] = removal_results
            results['total_removed'] += len(removal_results['removed'])
            results['total_failed'] += len(removal_results['failed_to_remove'])
        
        return results

    def list_staging_networks(self) -> None:
        """List configured staging networks"""
        if not STAGING_NETWORKS:
            print("No staging networks configured. Please update STAGING_NETWORKS in the script.")
            return
        
        print(f"\nConfigured staging networks ({len(STAGING_NETWORKS)}):")
        for name, network_id in STAGING_NETWORKS.items():
            print(f"  {name} -> {network_id}")


def load_environment_variables() -> tuple[str, str]:
    """Load API key and org ID from .env file"""
    load_dotenv()
    
    api_key = os.getenv('API_KEY')
    org_id = os.getenv('ORG_ID')
    
    if not api_key:
        raise ValueError("API_KEY not found in .env file")
    if not org_id:
        raise ValueError("ORG_ID not found in .env file")
    
    return api_key, org_id


def get_network_id_helper(network_input: str) -> str:
    """Helper function to resolve network ID"""
    # Check if it's a network name in our staging networks
    if network_input in STAGING_NETWORKS:
        return STAGING_NETWORKS[network_input]
    # Check if it's already a network ID (starts with 'N_')
    elif network_input.startswith('N_'):
        return network_input
    else:
        raise ValueError(f"Network '{network_input}' not found in configured staging networks")


def interactive_menu():
    """Interactive menu for users to select operations"""
    while True:
        print("\n" + "="*60)
        print("🔧 MX67 STAGING NETWORK MANAGEMENT TOOL")
        print("="*60)
        print("1. Check staging network capacity")
        print("2. List available MX67 devices")
        print("3. Smart batch add (distribute devices automatically)")
        print("4. Manual add to specific network")
        print("5. Remove device from network")
        print("6. Batch remove from network")
        print("7. 🧹 REMOVE ALL MX67s from ALL staging networks")
        print("8. List all staging networks")
        print("9. List devices in specific network")
        print("10. Exit")
        print("="*60)
        
        try:
            choice = input("Select an option (1-10): ").strip()
            
            if choice == "10":
                print("Goodbye!")
                break
            elif choice in ["1", "2", "3", "4", "5", "6", "7", "8", "9"]:
                return choice
            else:
                print("❌ Invalid choice. Please enter 1-10.")
        except KeyboardInterrupt:
            print("\n\nGoodbye!")
            break
    return None


def get_user_input():
    """Get user input for interactive operations"""
    
    # Load credentials
    try:
        api_key, org_id = load_environment_variables()
        print(f"✓ Loaded credentials from .env file (Org: {org_id})")
        manager = MerakiStagingManager(api_key, org_id)
    except Exception as e:
        print(f"❌ Error loading credentials: {e}")
        return
    
    while True:
        choice = interactive_menu()
        if choice is None:
            break
            
        try:
            if choice == "1":
                # Check capacity
                network_status = manager.check_staging_network_capacity()
                print("\nStaging Network Capacity Status:")
                print("=" * 60)
                
                for network_name, info in network_status.items():
                    if 'error' in info:
                        print(f"❌ {network_name}: Error - {info['error']}")
                    else:
                        status = "🟢 Available" if info['available_slots'] > 0 else "🔴 Full"
                        print(f"{status} {network_name}: {info['mx67_count']}/2 slots used ({info['available_slots']} available)")
                        
                        if info['devices']:
                            for device in info['devices']:
                                print(f"    - {device['serial']} ({device.get('name', 'No name')})")
                
                total_available = sum(info.get('available_slots', 0) for info in network_status.values())
                print(f"\nTotal available slots across all networks: {total_available}")
            
            elif choice == "2":
                # List available MX67s
                devices = manager.get_mx67_devices()
                print(f"\n📋 Found {len(devices)} available MX67 devices:")
                if devices:
                    for i, device in enumerate(devices, 1):
                        print(f"  {i:2d}. {device.serial} - {device.model} ({device.name or 'No name'})")
                else:
                    print("  No available MX67 devices found.")
            
            elif choice == "3":
                # Smart batch add
                print("\nEnter MX67 serial numbers (up to 20):")
                print("Enter one serial number per line, press Enter on blank line when done:")
                serials = []
                while True:
                    serial_input = input(f"Serial {len(serials)+1} (or press Enter to finish): ").strip()
                    if not serial_input:
                        break
                    serials.append(serial_input)
                    if len(serials) >= 20:
                        print("Maximum 20 devices reached.")
                        break
                
                if not serials:
                    print("❌ No serials entered.")
                    continue
                
                results = manager.smart_batch_add(serials)
                
                print("\n" + "="*60)
                print("SMART BATCH ADD RESULTS")
                print("="*60)
                
                if results['networks_with_existing_devices']:
                    print("⚠️  Networks with existing devices were found (see warnings above)")
                
                if results['added']:
                    total_added = sum(len(serials) for serials in results['added'].values())
                    print(f"Successfully distributed {total_added} devices across {len(results['added'])} networks:")
                    
                    for network_name, serials in results['added'].items():
                        print(f"\n📍 {network_name}:")
                        for serial in serials:
                            print(f"    ✓ {serial}")
                
                if results['failed_to_add']:
                    print(f"\nFailed to add: {len(results['failed_to_add'])}")
                    for serial in results['failed_to_add']:
                        print(f"    ✗ {serial}")
            
            elif choice == "4":
                # Manual add to specific network
                manager.list_staging_networks()
                network_input = input("\nEnter network name or ID: ").strip()
                if not network_input:
                    print("❌ No network specified.")
                    continue
                
                try:
                    network_id = get_network_id_helper(network_input)
                except ValueError as e:
                    print(f"❌ {e}")
                    continue
                
                serials_input = input("Enter MX67 serial numbers (max 2):\nEnter one per line, press Enter on blank line when done:\n").strip()
                if serials_input:
                    # If they entered something on the first line, use the old method as fallback
                    serials = serials_input.split()
                else:
                    # Use the new line-by-line method
                    serials = []
                    while True:
                        serial_input = input(f"Serial {len(serials)+1} (or press Enter to finish): ").strip()
                        if not serial_input:
                            break
                        serials.append(serial_input)
                        if len(serials) >= 2:
                            print("Maximum 2 devices reached.")
                            break
                
                if not serials:
                    print("❌ No serials entered.")
                    continue
                if len(serials) > 2:
                    print("❌ Maximum 2 devices allowed per network.")
                    continue
                
                # Check current capacity
                devices = manager.get_network_devices(network_id)
                mx67_count = len([d for d in devices if d.get('model', '').startswith('MX67')])
                
                if mx67_count + len(serials) > 2:
                    print(f"❌ Network already has {mx67_count}/2 MX67 devices. Cannot add {len(serials)} more.")
                    print("Current MX67 devices:")
                    for device in devices:
                        if device.get('model', '').startswith('MX67'):
                            print(f"  - {device['serial']} ({device.get('name', 'No name')})")
                    continue
                
                results = {'added': [], 'failed_to_add': []}
                for serial in serials:
                    if manager.add_mx67_to_network(network_id, serial):
                        results['added'].append(serial)
                    else:
                        results['failed_to_add'].append(serial)
                
                print("\n" + "="*50)
                print("MANUAL ADD RESULTS")
                print("="*50)
                print(f"Successfully added: {len(results['added'])}")
                for serial in results['added']:
                    print(f"  ✓ {serial}")
                
                if results['failed_to_add']:
                    print(f"\nFailed to add: {len(results['failed_to_add'])}")
                    for serial in results['failed_to_add']:
                        print(f"  ✗ {serial}")
            
            elif choice == "5":
                # Remove single device
                manager.list_staging_networks()
                network_input = input("\nEnter network name or ID: ").strip()
                if not network_input:
                    print("❌ No network specified.")
                    continue
                
                try:
                    network_id = get_network_id_helper(network_input)
                except ValueError as e:
                    print(f"❌ {e}")
                    continue
                
                # Show current devices in network
                devices = manager.get_network_devices(network_id)
                mx67_devices = [d for d in devices if d.get('model', '').startswith('MX67')]
                
                if not mx67_devices:
                    print("❌ No MX67 devices found in this network.")
                    continue
                
                print("\nCurrent MX67 devices in network:")
                for i, device in enumerate(mx67_devices, 1):
                    print(f"  {i}. {device['serial']} ({device.get('name', 'No name')})")
                
                serial = input("\nEnter MX67 serial number to remove: ").strip()
                if not serial:
                    print("❌ No serial entered.")
                    continue
                
                if manager.remove_mx67_from_network(network_id, serial):
                    print(f"✓ Successfully removed {serial}")
                else:
                    print(f"❌ Failed to remove {serial}")
            
            elif choice == "6":
                # Batch remove
                manager.list_staging_networks()
                network_input = input("\nEnter network name or ID: ").strip()
                if not network_input:
                    print("❌ No network specified.")
                    continue
                
                try:
                    network_id = get_network_id_helper(network_input)
                except ValueError as e:
                    print(f"❌ {e}")
                    continue
                
                # Show current devices in network
                devices = manager.get_network_devices(network_id)
                mx67_devices = [d for d in devices if d.get('model', '').startswith('MX67')]
                
                if not mx67_devices:
                    print("❌ No MX67 devices found in this network.")
                    continue
                
                print("\nCurrent MX67 devices in network:")
                for i, device in enumerate(mx67_devices, 1):
                    print(f"  {i}. {device['serial']} ({device.get('name', 'No name')})")
                
                serials_input = input("\nEnter MX67 serial numbers to remove:\nEnter one per line, press Enter on blank line when done:\n").strip()
                if serials_input:
                    # If they entered something on the first line, use the old method as fallback
                    serials = serials_input.split()
                else:
                    # Use the new line-by-line method
                    serials = []
                    while True:
                        serial_input = input(f"Serial {len(serials)+1} (or press Enter to finish): ").strip()
                        if not serial_input:
                            break
                        serials.append(serial_input)
                
                if not serials:
                    print("❌ No serials entered.")
                    continue
                results = manager.remove_mx67_batch(network_id, serials)
                
                print("\n" + "="*50)
                print("BATCH REMOVE RESULTS")
                print("="*50)
                print(f"Successfully removed: {len(results['removed'])}")
                for serial in results['removed']:
                    print(f"  ✓ {serial}")
                
                if results['failed_to_remove']:
                    print(f"\nFailed to remove: {len(results['failed_to_remove'])}")
                    for serial in results['failed_to_remove']:
                        print(f"  ✗ {serial}")
            
            elif choice == "7":
                # Remove ALL MX67s from ALL staging networks
                print("\n🧹 REMOVE ALL MX67s FROM ALL STAGING NETWORKS")
                print("="*60)
                print("⚠️  WARNING: This will remove ALL MX67 devices from ALL staging networks!")
                print("This operation cannot be undone.")
                
                results = manager.remove_all_mx67s_from_staging()
                
                if results['total_removed'] > 0 or results['total_failed'] > 0:
                    print("\n" + "="*60)
                    print("🧹 REMOVE ALL RESULTS")
                    print("="*60)
                    print(f"Total devices removed: {results['total_removed']}")
                    print(f"Total devices failed: {results['total_failed']}")
                    
                    for network_name, network_results in results['networks_processed'].items():
                        if network_results['removed'] or network_results['failed_to_remove']:
                            print(f"\n📍 {network_name}:")
                            for serial in network_results['removed']:
                                print(f"    ✓ {serial}")
                            for serial in network_results['failed_to_remove']:
                                print(f"    ✗ {serial}")
            
            elif choice == "8":
                # List staging networks
                manager.list_staging_networks()
            
            elif choice == "9":
                # List devices in specific network
                manager.list_staging_networks()
                network_input = input("\nEnter network name or ID: ").strip()
                if not network_input:
                    print("❌ No network specified.")
                    continue
                
                try:
                    network_id = get_network_id_helper(network_input)
                except ValueError as e:
                    print(f"❌ {e}")
                    continue
                
                devices = manager.get_network_devices(network_id)
                print(f"\n📋 Found {len(devices)} devices in network:")
                for device in devices:
                    model_indicator = "🔧" if device.get('model', '').startswith('MX67') else "📱"
                    print(f"  {model_indicator} {device['serial']} - {device['model']} ({device.get('name', 'No name')})")
        
        except Exception as e:
            print(f"❌ Error: {e}")
        
        input("\nPress Enter to continue...")


def main():
    parser = argparse.ArgumentParser(description="MX67 Staging Network Management Tool")
    parser.add_argument("--interactive", "-i", action="store_true", 
                       help="Run in interactive menu mode")
    
    subparsers = parser.add_subparsers(dest="command", help="Available commands")
    
    # List command
    list_parser = subparsers.add_parser("list", help="List available resources")
    list_parser.add_argument("--type", choices=["networks", "mx67s", "devices", "staging", "capacity"], 
                           default="staging", help="Resource type to list")
    list_parser.add_argument("--network-id", help="Network ID (for listing devices)")
    
    # Add command
    add_parser = subparsers.add_parser("add", help="Add MX67 to network")
    add_parser.add_argument("--network", required=True, 
                          help="Staging network name or ID (from configured networks)")
    add_parser.add_argument("--serial", required=True, help="MX67 serial number")
    add_parser.add_argument("--name", help="Device name")
    
    # Remove command
    remove_parser = subparsers.add_parser("remove", help="Remove MX67 from network")
    remove_parser.add_argument("--network", required=True, 
                             help="Staging network name or ID")
    remove_parser.add_argument("--serial", required=True, help="MX67 serial number")
    
    # Smart batch add command
    batch_add_parser = subparsers.add_parser("batch-add", help="Smart batch add: distribute MX67s across staging networks")
    batch_add_parser.add_argument("--serials", nargs="+", required=True, 
                                help="MX67 serial numbers (up to 20)")
    
    # Manual add command for specific network
    manual_add_parser = subparsers.add_parser("manual-add", help="Manually add MX67s to specific network")
    manual_add_parser.add_argument("--network", required=True, 
                                help="Staging network name or ID")
    manual_add_parser.add_argument("--serials", nargs="+", required=True, 
                                help="MX67 serial numbers (max 2 per network)")
    
    # Batch remove command
    batch_remove_parser = subparsers.add_parser("batch-remove", help="Remove multiple MX67s after sync")
    batch_remove_parser.add_argument("--network", required=True, 
                                   help="Staging network name or ID")
    batch_remove_parser.add_argument("--serials", nargs="+", required=True, 
                                   help="MX67 serial numbers")
    
    # Remove all command
    remove_all_parser = subparsers.add_parser("remove-all", help="Remove ALL MX67s from ALL staging networks")
    remove_all_parser.add_argument("--force", action="store_true", 
                                 help="Skip confirmation prompt")
    
    args = parser.parse_args()
    
    # If no command provided or interactive flag used, run interactive mode
    if not args.command or args.interactive:
        get_user_input()
        return
    
    # Command line mode - execute specific commands
    try:
        api_key, org_id = load_environment_variables()
        print(f"Loaded credentials from .env file (Org: {org_id})")
        manager = MerakiStagingManager(api_key, org_id)
        
        if args.command == "list":
            handle_list_command(args, manager, org_id)
        elif args.command == "add":
            handle_add_command(args, manager)
        elif args.command == "remove":
            handle_remove_command(args, manager)
        elif args.command == "batch-add":
            handle_batch_add_command(args, manager)
        elif args.command == "manual-add":
            handle_manual_add_command(args, manager)
        elif args.command == "batch-remove":
            handle_batch_remove_command(args, manager)
        elif args.command == "remove-all":
            handle_remove_all_command(args, manager)
            
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)


def handle_list_command(args, manager, org_id):
    """Handle list command"""
    if args.type == "staging":
        manager.list_staging_networks()
    
    elif args.type == "capacity":
        network_status = manager.check_staging_network_capacity()
        print("\nStaging Network Capacity Status:")
        print("=" * 60)
        
        for network_name, info in network_status.items():
            if 'error' in info:
                print(f"❌ {network_name}: Error - {info['error']}")
            else:
                status = "🟢 Available" if info['available_slots'] > 0 else "🔴 Full"
                print(f"{status} {network_name}: {info['mx67_count']}/2 slots used ({info['available_slots']} available)")
                
                if info['devices']:
                    for device in info['devices']:
                        print(f"    - {device['serial']} ({device.get('name', 'No name')})")
        
        total_available = sum(info.get('available_slots', 0) for info in network_status.values())
        print(f"\nTotal available slots across all networks: {total_available}")
    
    elif args.type == "networks":
        networks = manager.get_networks(org_id)
        print(f"\nFound {len(networks)} networks:")
        for net in networks:
            tags_str = ", ".join(net.get('tags', []))
            print(f"  {net['id']} - {net['name']} (tags: {tags_str})")
    
    elif args.type == "mx67s":
        devices = manager.get_mx67_devices()
        print(f"\nFound {len(devices)} available MX67 devices:")
        for device in devices:
            print(f"  {device.serial} - {device.model} ({device.name or 'No name'})")
    
    elif args.type == "devices" and args.network_id:
        devices = manager.get_network_devices(args.network_id)
        print(f"\nFound {len(devices)} devices in network:")
        for device in devices:
            print(f"  {device['serial']} - {device['model']} ({device.get('name', 'No name')})")


def handle_add_command(args, manager):
    """Handle add command"""
    network_id = get_network_id_helper(args.network)
    manager.add_mx67_to_network(network_id, args.serial, args.name)


def handle_remove_command(args, manager):
    """Handle remove command"""
    network_id = get_network_id_helper(args.network)
    manager.remove_mx67_from_network(network_id, args.serial)


def handle_batch_add_command(args, manager):
    """Handle batch-add command"""
    results = manager.smart_batch_add(args.serials)
    
    print("\n" + "="*60)
    print("SMART BATCH ADD RESULTS")
    print("="*60)
    
    if results['networks_with_existing_devices']:
        print("⚠️  Networks with existing devices were found (see warnings above)")
    
    if results['added']:
        total_added = sum(len(serials) for serials in results['added'].values())
        print(f"Successfully distributed {total_added} devices across {len(results['added'])} networks:")
        
        for network_name, serials in results['added'].items():
            print(f"\n📍 {network_name}:")
            for serial in serials:
                print(f"    ✓ {serial}")
    
    if results['failed_to_add']:
        print(f"\nFailed to add: {len(results['failed_to_add'])}")
        for serial in results['failed_to_add']:
            print(f"    ✗ {serial}")


def handle_manual_add_command(args, manager):
    """Handle manual-add command"""
    if len(args.serials) > 2:
        print("❌ Error: Maximum 2 devices allowed per network")
        sys.exit(1)
    
    network_id = get_network_id_helper(args.network)
    
    # Check current capacity
    devices = manager.get_network_devices(network_id)
    mx67_count = len([d for d in devices if d.get('model', '').startswith('MX67')])
    
    if mx67_count + len(args.serials) > 2:
        print(f"❌ Error: Network already has {mx67_count}/2 MX67 devices. Cannot add {len(args.serials)} more.")
        print("Current MX67 devices:")
        for device in devices:
            if device.get('model', '').startswith('MX67'):
                print(f"  - {device['serial']} ({device.get('name', 'No name')})")
        sys.exit(1)
    
    results = {'added': [], 'failed_to_add': []}
    for serial in args.serials:
        if manager.add_mx67_to_network(network_id, serial):
            results['added'].append(serial)
        else:
            results['failed_to_add'].append(serial)
    
    print("\n" + "="*50)
    print("MANUAL ADD RESULTS")
    print("="*50)
    print(f"Successfully added: {len(results['added'])}")
    for serial in results['added']:
        print(f"  ✓ {serial}")
    
    if results['failed_to_add']:
        print(f"\nFailed to add: {len(results['failed_to_add'])}")
        for serial in results['failed_to_add']:
            print(f"  ✗ {serial}")
    
    if results['added']:
        network_name = next((name for name, id in STAGING_NETWORKS.items() if id == network_id), network_id)
        print(f"\nTo remove devices after firmware sync, use:")
        print(f'python mx67_tool.py batch-remove --network "{network_name}" --serials {" ".join(results["added"])}')


def handle_batch_remove_command(args, manager):
    """Handle batch-remove command"""
    network_id = get_network_id_helper(args.network)
    results = manager.remove_mx67_batch(network_id, args.serials)
    
    print("\n" + "="*50)
    print("BATCH REMOVE RESULTS")
    print("="*50)
    print(f"Successfully removed: {len(results['removed'])}")
    for serial in results['removed']:
        print(f"  ✓ {serial}")
    
    if results['failed_to_remove']:
        print(f"\nFailed to remove: {len(results['failed_to_remove'])}")
        for serial in results['failed_to_remove']:
            print(f"  ✗ {serial}")


def handle_remove_all_command(args, manager):
    """Handle remove-all command"""
    if args.force:
        # For force mode, we need to bypass the confirmation
        print("🧹 Removing ALL MX67 devices from ALL staging networks...")
        print("Checking all staging networks for MX67 devices...")
        
        network_status = manager.check_staging_network_capacity()
        networks_with_devices = {name: info for name, info in network_status.items() 
                               if info.get('mx67_count', 0) > 0 and 'error' not in info}
        
        if not networks_with_devices:
            print("✓ No MX67 devices found in any staging networks.")
            return
        
        results = {
            'networks_processed': {},
            'total_removed': 0,
            'total_failed': 0
        }
        
        print("\nProceeding with removal...")
        
        # Remove devices from each network
        for network_name, info in networks_with_devices.items():
            network_id = info['network_id']
            serials_to_remove = [device['serial'] for device in info['devices']]
            
            print(f"\n📍 Processing {network_name}...")
            removal_results = manager.remove_mx67_batch(network_id, serials_to_remove)
            
            results['networks_processed'][network_name] = removal_results
            results['total_removed'] += len(removal_results['removed'])
            results['total_failed'] += len(removal_results['failed_to_remove'])
    else:
        results = manager.remove_all_mx67s_from_staging()
    
    if results['total_removed'] > 0 or results['total_failed'] > 0:
        print("\n" + "="*60)
        print("🧹 REMOVE ALL RESULTS")
        print("="*60)
        print(f"Total devices removed: {results['total_removed']}")
        print(f"Total devices failed: {results['total_failed']}")
        
        for network_name, network_results in results['networks_processed'].items():
            if network_results['removed'] or network_results['failed_to_remove']:
                print(f"\n📍 {network_name}:")
                for serial in network_results['removed']:
                    print(f"    ✓ {serial}")
                for serial in network_results['failed_to_remove']:
                    print(f"    ✗ {serial}")


if __name__ == "__main__":
    main()