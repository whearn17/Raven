# Raven v0.4 - Audit Log Retriever

Raven is a PowerShell script used to fetch audit log records from Microsoft Purview within a specified date range and exports them to a CSV file.

## Pre-requisites

To run this script, you need:

* PowerShell 5.1 or later.
* Exchange Online PowerShell V3 module.

## Usage

1. Open PowerShell.
2. Navigate to the folder containing the Raven script.
3. Run the script with the mandatory parameter, UserPrincipalName. For example:

```
.\Raven.ps1 -UserPrincipalName john.doe@domain.com
```

1. When prompted, enter the start date (as a number of days ago between 0 and 90, defaults to 90). For instance, to retrieve logs from 7 days ago to now, enter 7.
2. When prompted, enter the directory to save the audit logs. For example, C:\AuditLogs.
3. The script will start running, connecting to Exchange Online, checking if audit logging is enabled, retrieving audit logs, and finally saving them to the specified directory.

## Note:

* If you receive errors regarding 'Connect-ExchangeOnline' or 'Search-UnifiedAuditLog', please ensure that you have the Exchange Online PowerShell V3 module installed and are running the script with an account that has necessary permissions.

## License

Raven is open-source software, licensed under MIT License.
