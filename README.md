# Carters Network Refresh Project

## Overview
This repository contains Python scripts for the Carters network refresh project, focusing on device staging, management, and network template migrations.

## Project Description
The scripts provided for the Network Refresh Project automates various network operations including device staging, management tasks, and the migration of network configurations to new templates. This is aimed to reduce overall install time.

## Files in this Repository

### Core Scripts
- **`Network Refresh - CSR Device Staging.py`** - Handles the staging and initial configuration of CSR network devices
- **`Network Refresh - MOVE Network to Template.py`** - Automates the migration of network configurations to standardized templates
- **`Network Refresh - Device Management.py`** - Manages existing network devices and their configurations  

### Configuration
- **`.env`** - Environment variables and configuration settings (not tracked in git for security)
- This file must be in the same folder as the scripts.
- The API key will be delivered by Jeremy Elliot @ Carters

## Prerequisites
- Python 3.7 or higher
- Required Python packages (install via pip):
  ```bash
  pip install -r requirements.txt
  ```
- Network access to target devices
- Appropriate credentials and permissions

## Setup Instructions

1. **Clone the repository:**
   ```bash
   git clone https://github.com/JeremyElliott84/CSR.git
   cd CSR
   ```

2. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Configure environment variables:**
   - Copy `.env.example` to `.env` (if available)
   - Update the `.env` file with your specific network credentials and settings

4. **Verify network connectivity:**
   - Ensure you have access to the target network devices
   - Test connectivity before running the scripts

## Usage

### Device Staging
```bash
python "Network Refresh - CSR Device Staging.py"
```
Use this script to stage new Carters MX67 routers with initial configurations.

### Network Template Migration
```bash
python "Network Refresh - MOVE Network to Template.py"
```
Migrate existing network to the new template.

### Device Management  
```bash
python "Network Refresh - Device Management.py"
```
Removes old network devices, Adds new network devices, removes old DHCP reservations, creates new DHCP reservations, renames devices


## Important Notes

⚠️ **Security Notice**: 
- Never commit credentials or sensitive information to the repository
- Use environment variables or secure credential management
- Ensure `.env` files are in `.gitignore`

⚠️ **Network Operations Warning**:
- Test all scripts in a non-production environment first
- Always backup current configurations before making changes
- Verify network connectivity and permissions before execution


## Support
For questions or issues related to this project please contact Jeremy Elliott @ Carters OshKosh Inc.
jeremy.elliot@carters.com
762.232.1100 Office


---
*Last updated: September 2025*
