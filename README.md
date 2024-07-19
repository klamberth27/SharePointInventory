# SharePoint Sites Inventory Script

This PowerShell script processes users from a CSV file and fetches information about their SharePoint site memberships. It exports the results to a new CSV file.

## Prerequisites

1. **PowerShell**: Ensure you have PowerShell installed.
2. **SharePoint Online Management Shell**: This script uses the `Microsoft.Online.Sharepoint.PowerShell` module. The script will automatically install it if it's not already installed.

## Instructions

1. **Edit the Script**:
   - Open the script file.
   - Locate the following line:
     ```powershell
     $tenantUrl = "https://m365x72261057-admin.sharepoint.com"
     ```
   - Replace `"https://m365x72261057-admin.sharepoint.com"` with your own SharePoint tenant admin URL. It should look something like this:
     ```powershell
     $tenantUrl = "https://your-tenant-admin.sharepoint.com"
     ```

2. **Prepare the CSV File**:
   - Ensure you have a `Users.csv` file in the same directory as the script.
   - The CSV file should have the following structure:
     ```
     UserEmail
     user1@example.com
     user2@example.com
     ```

3. **Run the Script**:
   - Open PowerShell as an administrator.
   - Navigate to the directory containing the script.
   - Run the script:
     ```powershell
     .\YourScriptName.ps1
     ```
   - Authenticate using Admin account
   - The script will connect to your SharePoint tenant, process the users, and export the results to `CMSSPSites.csv` in the same directory.

4. **Check the Output**:
   - The results will be saved in `CMSSPSites.csv`.

## Troubleshooting

- **No CSV file to read**: Ensure the `Users.csv` file is present and correctly formatted.
- **Permission Issues**: Ensure you have the necessary permissions to connect to the SharePoint tenant and access the sites.
- **Module Installation**: If the script fails to install the module, try manually installing it:
  ```powershell
  Install-Module -Name Microsoft.Online.Sharepoint.PowerShell -Force
  ```

## Additional Notes

- This script is designed to be run in an environment with access to the SharePoint Online Management Shell.
- Modify the script as needed to suit your requirements.

---
