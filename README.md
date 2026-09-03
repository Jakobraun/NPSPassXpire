# Password Expiration Report

A portable Windows PowerShell application that queries Microsoft Graph for staff accounts whose passwords expired within the last 100 days or are expected to expire soon. It produces a seven-day CSV report and groups email addresses for already-expired, one-, three-, and seven-day notifications.
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

Before each run, the application filters the source table and writes unique valid names to `staff_sanitized.csv`. It accepts either `First Name` and `Last Name` columns or an asset export with a `Location` column. In an asset export, values beginning with `Staff:` are converted from full names to first and last names; `Room:` and `Student:` rows are excluded. Valid two-column first-name/last-name rows can also be appended to an asset export. Non-English and unsupported characters are removed, and consecutive whitespace is reduced to one space. For example, `Sarah ðŸ˜ƒ` and `Connor ðŸš€` become `Sarah` and `Connor`. Names are formatted with an initial capital and lowercase remaining letters, with each hyphenated segment capitalized, then sorted by last name. Rows with missing names, unsafe spreadsheet formulas, names longer than 80 characters, malformed CSV data, or reserved placeholder values are skipped. Reserved values are matched case-insensitively: `NULL`, `none`, `TRUE`, `FALSE`, and `na`. The account query and site matching use only the sanitized file.

Generated reports are named `PwdExpire_MM-dd_HH-mm.csv` and are intentionally excluded from Git.

## Staff CSV format

```csv
First Name,Last Name
Alex,Rivera
Taylor,Morgan
```

The application accepts up to 999 non-empty rows and a maximum file size of 1 MB. Blank and comma-only rows are ignored. The table must include `First Name` and `Last Name` headers, or a `Location` column containing `Staff:` values; other columns are ignored. Sanitized names contain only English letters, single spaces, periods, apostrophes, and hyphens.

## Privacy and security

This application handles directory and staff information. Do not commit:

- `staff.csv`, which contains the real staff list
- `staff_sanitized.csv`, which contains the filtered working staff list
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
