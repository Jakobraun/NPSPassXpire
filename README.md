# Password Expiration Report

A portable Windows PowerShell application that queries Microsoft Graph for staff accounts whose passwords are expected to expire soon. It produces a seven-day CSV report and groups email addresses for one-, three-, and seven-day notifications. 
This is designed to work at Norfolk Public Schools where a password expiration policy of 180 days is still in place in 2026 for staff accounts.

## Requirements

- Windows with Windows PowerShell 5.1 and WPF
- A Microsoft Entra ID account permitted to read users through Microsoft Graph
- Access to PowerShell Gallery on the first run
- Microsoft Edge or Google Chrome for the automatic private device-login window

The first run installs `Microsoft.Graph.Authentication` and `Microsoft.Graph.Users` for the current Windows user.

## Setup

1. Create `staff.csv` in the application folder. Keep the headers exactly as shown below.
2. Add the required staff names beneath the header row.
3. Review the organization-specific values in `Password Expiration Worker.ps1`, particularly the staff domains and `$maxAgeDays`.
4. Run `Launch Password Expiration Report.cmd`.
5. Select **Connect & Run** and complete Microsoft device authentication.

If `staff.csv` is missing or contains no names, the application stops before Microsoft sign-in and does not query directory accounts.

Generated reports are named `PwdExpire_MM-dd_HH-mm.csv` and are intentionally excluded from Git.

## Staff CSV format

```csv
First Name,Last Name
Alex,Rivera
Taylor,Morgan
```

The application accepts fewer than 500 lines and a maximum file size of 128 KB. Names are validated before use.

## Privacy and security

This application handles directory and staff information. Do not commit:

- `staff.csv`, which contains the real staff list
- generated `PwdExpire_*.csv` reports, which contain names, UPNs, email addresses, and expiration information
- user-specific settings files
- Outlook `.oft` templates unless their content and embedded metadata have been reviewed

These files are excluded by `.gitignore`. No credentials are intentionally written by the application, but Microsoft Graph authentication and report handling should still follow the organization's security policies.

## Current limitations
- Site matching is name-based rather than UPN- or email-based, so duplicate or unusually formatted names may match incorrectly.
- Multiple simultaneous instances use the same temporary status files and may interfere with one another.
- The scripts are not digitally signed. The launcher uses `ExecutionPolicy Bypass` for the process.

## Repository contents

- `Password Expiration App.ps1` — WPF interface, CSV validation, and staff matching
- `Password Expiration Worker.ps1` — Graph authentication, account query, and report generation
- `Launch Password Expiration Report.cmd` — Windows launcher
- `email templates/README.md` — guidance for locally maintained Outlook templates

## Development check

PowerShell syntax can be checked without running the application:

```powershell
$tokens = $null
$errors = $null
[void][System.Management.Automation.Language.Parser]::ParseFile(
    '.\Password Expiration App.ps1',
    [ref]$tokens,
    [ref]$errors
)
$errors
```

Repeat the check for `Password Expiration Worker.ps1`.

## License

This project is licensed under the GNU General Public License v3.0. See [LICENSE](LICENSE) for the full terms.
