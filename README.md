# Exchange-Server-Toolkit
A collection of PowerShell scripts for Microsoft Exchange Server covering configuration reporting, health checks, automation, and administrative tasks.
# Exchange Server Configuration Detailed Report

This PowerShell script generates a **detailed Exchange Server configuration report** in **HTML format**.

The report collects configuration details from your on-premises Microsoft Exchange environment and exports them into a structured HTML file for review, documentation, auditing, or compliance purposes.

The script must be executed from **Exchange Management Shell (EMS)**..


# Features

-   Collects Exchange Server configuration details
-   Generates structured HTML output
-   Easy to review and share
Useful for:
	-   Health checks
	-   Documentation
	-   Migration preparation
	-   Auditing
	-   Compliance reviews

## Requirements

-   On-Premises Microsoft Exchange Server
-   Exchange Management Shell (EMS)
-   Appropriate administrative permissions
-   PowerShell 5.1 (recommended)

## How to Run the Script

**Open Exchange Management Shell**
Run the script **ONLY** from:
Exchange Management Shell
Do NOT use normal PowerShell unless Exchange modules are manually loaded.

## Execute the Script
**.\Exchange-Config-Detailed-Report.ps1**
If script execution is restricted:
**Set-ExecutionPolicy RemoteSigned -Scope Process**
Then run the script again

## Output
After execution:
-   An **HTML file** will be generated.
-   Default output location:

## The report includes:
-   Server Information
-   Database Configuration
-   Mailbox Details
-   Send/Receive Connectors
-   Virtual Directories
-   Transport Configuration
-   Certificates
-   DAG Information (if applicable)
And more….
