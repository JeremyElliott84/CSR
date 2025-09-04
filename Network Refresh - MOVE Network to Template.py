#!/usr/bin/env python3
"""
Script to move one or more Meraki networks to a new configuration template while preserving existing VLAN subnets,
interface IPs, and DHCP fixed IP assignments (extracted from each VLAN).
Automatically adjusts VLAN 1 subnet based on the target template's VLAN structure by querying the template network's VLANs directly:
  - If template network includes both VLAN 1 and VLAN 4: restores original /27 on VLAN 1
  - If template network includes VLAN 1 but not VLAN 4: calculates and applies a /26 merge of VLANs 1 & 4

Includes 20-second pauses only after unbind and bind steps to prevent API rate limiting and allow changes to finalize.

Now includes spreadsheet integration to lookup network and template IDs by store number.

Requirements:
    pip install meraki python-dotenv pandas openpyxl

Usage:
    1. Create a .env file with your API key, organization ID, and data file path:
       API_KEY=your_api_key_here
       ORG_ID=your_org_id_here
       DATA_FILE_PATH=C:\\path\\to\\your\\Network Refresh Tool - Data Sheet.xlsx
    2. Run: python move_network_template.py [--base-url BASE_URL]
"""

# =============================================================================
# CONFIGURATION SECTION - UPDATE THESE VALUES
# =============================================================================

# Spreadsheet Configuration - Column Mappings
SPREADSHEET_COLUMNS = {
    'store_number': 'storeNumber',  # Column A
    'network_id': 'NetworkID',      # Column B  
    'template_id': 'New Template ID' # Column C
}

# Switch Port Profile Configuration
# SW1 = Switch 1 -> Port Profile A (e.g., "US Tier A - MS120-A")
# SW2 = Switch 2 -> Port Profile B (e.g., "US Tier A - 120-B")
PORT_PROFILE_MAPPINGS = {
    'MS120': {
        'SW1': 'US Tier A - MS120-A',    # Switch 1 -> Profile A
        'SW2': 'US Tier A - 120-B'       # Switch 2 -> Profile B
    },
    'MS130': {
        'SW1': 'US Tier A - MS130-A',    # Switch 1 -> Profile A
        'SW2': 'US Tier A - MS130-B'     # Switch 2 -> Profile B
    }
}

# =============================================================================
# END CONFIGURATION SECTION
# =============================================================================

# Configuration
VLAN_IDS = [1,2,3,4,5,7,999]
PAUSE_SECONDS = 20

import sys
import time
import meraki
import ipaddress
import argparse
import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from dotenv import load_dotenv
import re

# Load environment variables from .env file
load_dotenv()

def get_env_variables():
    """Get API key, organization ID, and data file path from environment variables."""
    api_key = os.getenv('API_KEY')
    org_id = os.getenv('ORG_ID')
    data_file_path = os.getenv('DATA_FILE_PATH')
    
    if not api_key:
        print("Error: API_KEY not found in .env file")
        sys.exit(1)
    
    if not org_id:
        print("Error: ORG_ID not found in .env file")
        sys.exit(1)
    
    return api_key, org_id, data_file_path

def get_file_path(env_file_path):
    """
    Get the file path either from environment variable or user selection
    
    Args:
        env_file_path (str): File path from environment variable (can be None)
    
    Returns:
        str: File path to the spreadsheet
    """
    if env_file_path and os.path.exists(env_file_path):
        print(f"Using file path from .env: {env_file_path}")
        return env_file_path
    elif env_file_path:
        print(f"Warning: DATA_FILE_PATH from .env file does not exist: {env_file_path}")
        print("Falling back to file selection dialog...")
    
    # Hide the main tkinter window
    root = tk.Tk()
    root.withdraw()
    
    print("Please select the network refresh spreadsheet file...")
    
    # Open file dialog
    file_path = filedialog.askopenfilename(
        title="Select Network Refresh Spreadsheet",
        filetypes=[
            ("Excel files", "*.xlsx *.xls"),
            ("CSV files", "*.csv"),
            ("All files", "*.*")
        ]
    )
    
    if not file_path:
        print("No file selected. Exiting.")
        return None
    
    print(f"Selected file: {file_path}")
    return file_path

def load_spreadsheet_data(file_path: str) -> dict:
    """
    Load store data from network refresh spreadsheet
    
    Args:
        file_path (str): Path to the spreadsheet file
        
    Returns:
        dict: Dictionary with store numbers as keys and store data as values
    """
    try:
        # Determine file type and load accordingly
        if file_path.endswith('.csv'):
            df = pd.read_csv(file_path)
        else:
            df = pd.read_excel(file_path)
        
        print(f"Loaded {len(df)} rows from spreadsheet")
        print(f"Columns found: {list(df.columns)}")
        
        # Check for required columns (with flexible whitespace matching)
        missing_columns = []
        column_mapping = {}
        
        # Create a mapping of stripped column names to actual column names
        actual_columns = {col.strip(): col for col in df.columns}
        
        for key, expected_col in SPREADSHEET_COLUMNS.items():
            # Try exact match first
            if expected_col in df.columns:
                column_mapping[key] = expected_col
            # Try stripped match
            elif expected_col.strip() in actual_columns:
                column_mapping[key] = actual_columns[expected_col.strip()]
                print(f"Mapped '{expected_col}' to '{actual_columns[expected_col.strip()]}' (whitespace difference)")
            else:
                missing_columns.append(expected_col)
        
        if missing_columns:
            print(f"Warning: Missing expected columns: {missing_columns}")
            print("Available columns:")
            for i, col in enumerate(df.columns):
                print(f"  {i+1}. {col}")
            
            # Allow user to map missing columns
            for key, expected_col in SPREADSHEET_COLUMNS.items():
                if key not in column_mapping:
                    print(f"\nWhich column contains '{expected_col}' data? Enter column name or number (or 'skip'):")
                    user_input = input().strip()
                    
                    if user_input.lower() == 'skip':
                        continue
                    elif user_input.isdigit():
                        col_index = int(user_input) - 1
                        if 0 <= col_index < len(df.columns):
                            column_mapping[key] = df.columns[col_index]
                        else:
                            print(f"Invalid column number for {key}")
                    else:
                        if user_input in df.columns:
                            column_mapping[key] = user_input
                        else:
                            print(f"Column '{user_input}' not found for {key}")
        
        # Process data by store
        stores_data = {}
        
        for _, row in df.iterrows():
            # Skip rows without required data
            store_number = row.get(column_mapping.get('store_number', ''), '')
            network_id = row.get(column_mapping.get('network_id', ''), '')
            template_id = row.get(column_mapping.get('template_id', ''), '')
            
            if pd.isna(store_number) or pd.isna(network_id) or pd.isna(template_id):
                continue
                
            if not str(store_number).strip() or not str(network_id).strip() or not str(template_id).strip():
                continue
            
            store_number = str(store_number).strip()
            network_id = str(network_id).strip()
            template_id = str(template_id).strip()
            
            stores_data[store_number] = {
                'store_number': store_number,
                'network_id': network_id,
                'template_id': template_id
            }
        
        print(f"Found {len(stores_data)} stores with complete data in spreadsheet")
        return stores_data
        
    except Exception as e:
        print(f"Error loading spreadsheet: {str(e)}")
        return {}

def prompt_store_number(stores_data: dict):
    """
    Prompt user to enter a single store number and return corresponding network/template data
    
    Args:
        stores_data (dict): Dictionary of store data loaded from spreadsheet
        
    Returns:
        tuple: (network_id, template_id, store_number)
    """
    print(f"\nFound {len(stores_data)} stores in spreadsheet.")
    available_stores = sorted(stores_data.keys())
    
    # Show examples of available stores
    if len(available_stores) > 10:
        print("Example stores:", ", ".join(available_stores[:10]) + "...")
    else:
        print("Available stores:", ", ".join(available_stores))
    
    while True:
        store_input = input("\nEnter store number to move: ").strip()
        
        if store_input in stores_data:
            store_data = stores_data[store_input]
            print(f"Selected Store {store_input}:")
            print(f"  Network ID: {store_data['network_id']}")
            print(f"  Template ID: {store_data['template_id']}")
            return (store_data['network_id'], store_data['template_id'], store_data['store_number'])
        else:
            print(f"Store '{store_input}' not found.")
            if len(available_stores) <= 20:
                print(f"Available stores: {', '.join(available_stores)}")
            else:
                print(f"Available stores: {', '.join(available_stores[:20])}...")
                print(f"(and {len(available_stores) - 20} more...)")

def main():
    parser = argparse.ArgumentParser(
        description="Move Meraki network(s) to new templates based on store numbers from spreadsheet."
    )
    parser.add_argument(
        "--api-key", dest="api_key",
        help="Optional: override the API key from .env file."
    )
    parser.add_argument(
        "--org-id", dest="org_id",
        help="Optional: override the organization ID from .env file."
    )
    parser.add_argument(
        "--base-url", dest="base_url", default="https://api.meraki.com/api/v1",
        help="Base URL for the Meraki API."
    )
    args = parser.parse_args()

    # Get API key, org ID, and data file path from .env file or command line arguments
    env_api_key, env_org_id, env_data_file_path = get_env_variables()
    api_key = args.api_key or env_api_key
    org_id = args.org_id or env_org_id

    print(f"Using organization ID: {org_id}")
    
    dashboard = meraki.DashboardAPI(
        api_key=api_key,
        base_url=args.base_url,
        log_file_prefix=False
    )

    # Load spreadsheet data
    print("\n=== Loading Network Refresh Spreadsheet ===")
    file_path = get_file_path(env_data_file_path)
    if not file_path:
        sys.exit(1)
        
    stores_data = load_spreadsheet_data(file_path)
    if not stores_data:
        print("No store data found in spreadsheet. Exiting.")
        sys.exit(1)

    # Get store number and corresponding network/template data
    network_id, template_network_id, store_number = prompt_store_number(stores_data)

    # Confirmation prompt for the selected store
    print(f"\n=== Confirmation ===")
    print(f"You are about to move:")
    print(f"  Store {store_number}: Network {network_id} -> Template {template_network_id}")
    
    confirm = input(f"\nProceed with moving this network? (y/n): ").strip().lower()
    if confirm != 'y':
        print("Operation cancelled by user.")
        sys.exit(0)

    # Process the selected network
    print(f"\n=== Processing Store {store_number}: Network {network_id} -> Template {template_network_id} ===")

    # 1. Snapshot VLANs and fixed IP assignments
    vlan_snapshot = []
    print("Snapshotting source network VLANs and fixed IP assignments...")
    for vid in VLAN_IDS:
        try:
            v = dashboard.appliance.getNetworkApplianceVlan(network_id, vid)
            vlan_snapshot.append({
                'id': v['id'],
                'name': v['name'],
                'subnet': v['subnet'],
                'applianceIp': v['applianceIp'],
                'groupPolicyId': v.get('groupPolicyId'),
                'fixedIpAssignments': v.get('fixedIpAssignments', {})
            })
            print(f" - VLAN {vid}: subnet={v['subnet']} ip={v['applianceIp']}")
            for mac, entry in v.get('fixedIpAssignments', {}).items():
                print(f"   · fixed IP {mac} -> {entry.get('ip')} ({entry.get('name')})")
        except Exception:
            print(f" [!] VLAN {vid} not found in source; skipping.")

    # 1b. Unbind network from current template
    try:
        print(f"Unbinding network {network_id} from current template (retain configs)...")
        dashboard.networks.unbindNetwork(network_id, retainConfigs=True)
        print("Unbound successfully.")
    except Exception as e:
        print(f"Error unbinding network: {e}")
        print("Exiting due to unbind failure.")
        sys.exit(1)
    time.sleep(PAUSE_SECONDS)

    # 2. Determine VLAN structure on template network
    print("Checking template network VLANs...")
    template_vlans = []
    for vid in [1,4]:
        try:
            tv = dashboard.appliance.getNetworkApplianceVlan(template_network_id, vid)
            template_vlans.append(tv['id'])
            print(f" - Template VLAN {vid} exists: subnet={tv['subnet']}")
        except Exception:
            print(f" [!] Template VLAN {vid} not found; skipping.")
    has_v1 = 1 in template_vlans
    has_v4 = 4 in template_vlans
    if has_v1 and not has_v4:
        v1 = next((v for v in vlan_snapshot if v['id'] == 1), None)
        v4 = next((v for v in vlan_snapshot if v['id'] == 4), None)
        if v1 and v4:
            net1 = ipaddress.ip_network(v1['subnet'], strict=False)
            net4 = ipaddress.ip_network(v4['subnet'], strict=False)
            start = min(net1.network_address, net4.network_address)
            merged = ipaddress.ip_network(f"{start}/26", strict=False)
            print(f"Applying merged VLAN1 /26 {merged} based on source data")
            v1['subnet'] = str(merged)
            vlan_snapshot = [v for v in vlan_snapshot if v['id'] != 4]
    elif has_v1 and has_v4:
        print("Template includes both VLAN1 and VLAN4; preserving original VLAN1 subnet.")
    else:
        print("Template missing VLAN1; cannot apply VLAN structure correctly.")

    # 3. Bind to new template network
    try:
        print(f"Binding {network_id} to template network {template_network_id}...")
        dashboard.networks.bindNetwork(network_id, template_network_id)
        print(f"Bound to template network {template_network_id}")
    except Exception as e:
        print(f"Error binding network: {e}")
        print("Exiting due to bind failure.")
        sys.exit(1)
    time.sleep(PAUSE_SECONDS)

    # 4. Restore VLANs and only VLAN 1 fixed IPs on moved network
    print("Restoring VLAN settings and VLAN 1 fixed IP assignments on moved network...")
    for v in vlan_snapshot:
        try:
            params = {
                'name': v['name'],
                'subnet': v['subnet'],
                'applianceIp': v['applianceIp']
            }
            if v.get('groupPolicyId'):
                params['groupPolicyId'] = v['groupPolicyId']
            if v['id'] == 1 and v['fixedIpAssignments']:
                params['fixedIpAssignments'] = v['fixedIpAssignments']
            dashboard.appliance.updateNetworkApplianceVlan(network_id, v['id'], **params)
            if v['id'] == 1:
                print(f" VLAN {v['id']} restored with fixed IPs: {list(v['fixedIpAssignments'].keys())}")
            else:
                print(f" VLAN {v['id']} restored.")
        except Exception as e:
            print(f"Error updating VLAN {v['id']}: {e}")
    
    print(f"✓ Completed processing Store {store_number}")

    print(f"\n🎉 Successfully processed Store {store_number}!")


if __name__ == '__main__':
    main()