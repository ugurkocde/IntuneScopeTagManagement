# Intune Scope Tag Management

![screenshot](https://github.com/ugurkocde/IntuneScopeTagManagement/assets/43906965/fbf1cfac-2450-4af9-9414-c8fc61fa7fb2)


# Overview
This PowerShell script provides a graphical user interface (GUI) for managing Intune Scope Tags. It allows users to connect to Microsoft Graph, retrieve and display managed devices and their scope tags, assign scope tags to devices, and export the data to a CSV file.

# Features
Graph API Integration: Connects to Microsoft Graph to fetch and manage Intune data.
Manage Scope Tags: View and assign scope tags to managed devices.
Search Functionality: Filter devices by name or ID.
Data Export: Export the list of devices and their associated scope tags to a CSV file.
User-Friendly Interface: A graphical interface for ease of use.

# Work in progress
- Support for Policies
- Support for Apps

# Requirements
1. PowerShell 5.1 or higher.
2. Microsoft Graph PowerShell SDK.
3. Microsoft Graph Permissions: The user must have appropriate permissions configured in Microsoft Graph to manage Intune devices and scope tags.

   The script requires the following specific permissions (scopes) in Microsoft Graph:
- 3.1 **DeviceManagementManagedDevices.ReadWrite.All**: This permission is necessary to read and write managed device information.
- 3.2 **DeviceManagementRBAC.Read.All**: This scope is required to read role-based access control (RBAC) settings, which include scope tags.


To set up these permissions, an Azure AD application with delegated permissions for the above scopes is required. The user executing the script must be granted these permissions in Azure AD and consented to these delegated permissions.

# Installation
1. Ensure that PowerShell 5.1 or higher is installed on your system.
2. Install the Microsoft Graph PowerShell SDK using the following command:
`Install-Module Microsoft.Graph -Scope CurrentUser`
3. Download the script file to your local machine.

# Usage
1. Open PowerShell and navigate to the directory where the script is located.
2. Run the script:
`.\IntuneScopeTagManagement.ps1`
3. Use the GUI to connect to Microsoft Graph, manage scope tags, and export data as needed.

## Contributing
Feel free to fork the repository and submit pull requests. For major changes, please open an issue first to discuss what you would like to change.
