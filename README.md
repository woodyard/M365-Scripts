# M365-Scripts

PowerShell scripts for administering Microsoft 365 and Microsoft Entra ID tenants. Each script is self-contained: prerequisites, parameters, and examples are documented in the script's comment-based help block (`Get-Help .\Script-Name.ps1 -Full`).

## Structure

Scripts are organised by the Microsoft 365 workload they target. Folders are created on demand - if a category has no scripts in it yet, the folder isn't there.

| Folder | Scope |
|---|---|
| [`Identity/`](Identity/) | Microsoft Entra ID: users, groups, sync, MFA, Conditional Access |
| `Exchange/` | Exchange Online: mailboxes, archives, distribution lists, mail flow |
| `SharePoint/` | SharePoint Online and OneDrive: sites, permissions, lists |
| `Teams/` | Microsoft Teams: teams, channels, policies, phone system |
| [`Intune/`](Intune/) | Intune / Endpoint: devices, apps, detection and remediation scripts |
| `Security/` | Purview / Security & Compliance: DLP, retention, eDiscovery, audit |
| `Licensing/` | License assignment, group-based licensing, reporting |

## Usage

Each script documents its own prerequisites and required Graph / module permissions at the top of the file. As a general rule:

- Read the comment-based help before running anything. `Get-Help .\Path\To\Script.ps1 -Full`.
- Most scripts that talk to Microsoft Graph require the `Microsoft.Graph` PowerShell module. Install with `Install-Module Microsoft.Graph -Scope CurrentUser`.
- Scripts that make changes support `-WhatIf` where possible. Use it first.

## Contributing

If you're adding a new script, drop it into the folder that matches the workload, or create a new top-level folder following the same pattern if none fits. Keep the comment-based help block up to date.

## License

[MIT](LICENSE) - use, adapt, redistribute freely.
