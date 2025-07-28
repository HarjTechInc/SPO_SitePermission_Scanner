# HarjTech SharePoint Online Permission Audit Tool

## Overview
This PowerShell script helps SharePoint administrators audit a SharePoint Online site for permission inheritance and content exposure.  The script connects to a specified site, scans all **document libraries** and their contents (files and folders), and generates two outputs:

* **CSV report:** detailed list of every file and folder with its library name, full path, item type (file/folder), whether it has unique permissions (broken inheritance), whether it is shared with “Everyone” or “Everyone except external users”, and the file type (extension).  
* **Summary report:** a text file summarising the total number of files and folders, counts of items with unique permissions, counts by file type, and a list of libraries that continue to inherit permissions from the parent site.

The script uses **PnP.PowerShell** (Microsoft’s recommended cmdlets for SharePoint Online) and supports interactive sign‑in.  It is branded and provided by **HarjTech** as a complimentary utility.  For more advanced governance, migration or automation projects, please visit **[www.harjtech.com](https://www.harjtech.com)**.

## Prerequisites
1. **PowerShell 7.x** or **Windows PowerShell 5.1** installed on your machine.
2. **PnP.PowerShell** module installed.  If the script cannot find the module, it will attempt to install it for the current user.
3. SharePoint Online administrative or site‑owner permissions to the site you are auditing.
4. Internet connectivity (for authentication and data retrieval).

## How It Works
* The script prompts for a **SiteUrl** parameter, which should be the full URL of the SharePoint Online site (e.g., `https://contoso.sharepoint.com/sites/ProjectX`).
* It connects to the site using `Connect-PnPOnline` with interactive login.
* It retrieves all document libraries (base template 101) and iterates through every item in each library.  Items are differentiated by the `FileSystemObjectType` property; value `0` indicates a file and `1` indicates a folder【879391428515485†L318-L320】.
* For each item the script records:
  - **UniquePermissions:** whether `HasUniqueRoleAssignments` is `true` (broken inheritance)【879391428515485†L318-L320】.  
  - **ExposedToEveryone:** whether any role assignment includes a member whose title contains “Everyone”.  
  - **FileType:** the file extension for files (e.g., docx, pdf), used to build a file‑type distribution.
* Counts of files, folders, unique permissions and file types are aggregated, and libraries that still inherit permissions (i.e., `HasUniqueRoleAssignments` is `false`) are listed.
* All detailed information is exported to a CSV file, and a human‑readable summary is saved to a text file.  The script prints a brief summary and a marketing message upon completion.

## Usage
1. **Extract the contents** of the downloaded ZIP file to a local folder.
2. Open **PowerShell** and change to the extracted folder.
3. Run the script with the required `SiteUrl` parameter.  For example:
   ```powershell
   .\HarjTech_SharePoint_Permission_Scanner.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/ProjectX"
   ```
4. Sign in with your SharePoint Online credentials when prompted.
5. When the script completes, locate the generated `SPPermissionsReport-<timestamp>.csv` and `SPPermissionsSummary-<timestamp>.txt` files in the same folder.

## Output Details
**CSV columns:**
- **Library:** Name of the document library containing the item.
- **ItemPath:** Server relative URL to the item.
- **ItemName:** File or folder name.
- **ItemType:** `File` or `Folder`.
- **UniquePermissions:** `True` if the item breaks permission inheritance.
- **ExposedToEveryone:** `True` if the item is shared with the built‑in “Everyone” group.
- **FileType:** File extension without the dot (blank for folders).

**Summary report:** lists totals, counts by file type, and libraries that do not have custom permissions.

## Disclaimer
This script is provided **as‑is** without warranties of any kind.  HarjTech is not responsible for any damages or loss resulting from the use of this script.  Use at your own risk and test in a non‑production environment first.

## About HarjTech
HarjTech specialises in **SharePoint, Microsoft Teams and Power Platform** solutions for organisations looking to modernise workflows and enhance collaboration.  Our services include:

* SharePoint migration and clean‑up
* Power Platform quick‑start and automation
* Teams governance and adoption
* Document management and digital forms
* Managed systems and support

To learn more or request a custom engagement, visit **[www.harjtech.com](https://www.harjtech.com)** or email **contact@harjtech.com**.
