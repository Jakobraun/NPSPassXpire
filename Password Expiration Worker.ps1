param(
    [Parameter(Mandatory)] [string]$ResultPath,
    [Parameter(Mandatory)] [string]$StatusPath,
    [Parameter(Mandatory)] [ValidateRange(1,365)] [int]$SearchDays
)

$ErrorActionPreference = 'Stop'
$appDirectory = Split-Path -Parent $PSCommandPath
$timestamp = Get-Date -Format 'MM-dd_HH-mm'
$baseName = "PwdExpire_$timestamp"
$csvPath = Join-Path $appDirectory ($baseName + '.csv')
$suffix = 2
while (Test-Path -LiteralPath $csvPath) {
    $csvPath = Join-Path $appDirectory ("{0}_{1}.csv" -f $baseName, $suffix)
    $suffix++
}

function Initialize-GraphDependencies {
    $requiredModules = @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Users')
    $missingModules = @($requiredModules | Where-Object {
        -not (Get-Module -ListAvailable -Name $_)
    })

    if ($missingModules.Count -gt 0) {
        Write-Output "First-run setup: installing $($missingModules -join ', ') for the current user..."

        # Windows PowerShell 5.1 defaults can predate the TLS version required by PSGallery.
        [Net.ServicePointManager]::SecurityProtocol =
            [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

        if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
            Install-PackageProvider -Name NuGet -Scope CurrentUser -Force -Confirm:$false | Out-Null
        }

        $repository = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
        if (-not $repository) {
            Register-PSRepository -Default
            $repository = Get-PSRepository -Name PSGallery
        }

        # Trust only for the unattended install and restore the user's prior setting afterward.
        $originalPolicy = $repository.InstallationPolicy
        try {
            if ($originalPolicy -ne 'Trusted') {
                Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
            }
            foreach ($moduleName in $missingModules) {
                Install-Module -Name $moduleName -Repository PSGallery -Scope CurrentUser `
                    -Force -AllowClobber -Confirm:$false
            }
        } finally {
            if ($originalPolicy -and $originalPolicy -ne 'Trusted') {
                Set-PSRepository -Name PSGallery -InstallationPolicy $originalPolicy -ErrorAction SilentlyContinue
            }
        }
    }

    foreach ($moduleName in $requiredModules) {
        Import-Module $moduleName -Force -ErrorAction Stop
    }
}

try {
    Initialize-GraphDependencies

    Write-Output 'Follow the Microsoft device sign-in instructions below.'
    Connect-MgGraph -Scopes 'User.Read.All' -UseDeviceCode -ContextScope Process -NoWelcome
    'AUTHENTICATED' | Set-Content -LiteralPath $StatusPath -Encoding ASCII
    Write-Output 'Authenticated. Checking password expiration dates...'
    'CHECKING' | Set-Content -LiteralPath $StatusPath -Encoding ASCII

    $maxAgeDays = 180
    $nowUtc = [DateTime]::UtcNow
    $staffDomain = '@nps.k12.va.us'
    $studentDomain = '@npsk12.net'

    $reportRows = @(Get-MgUser -All -Property DisplayName,GivenName,Surname,UserPrincipalName,Mail,UserType,AccountEnabled,LastPasswordChangeDateTime,OnPremisesLastPasswordChangeDateTime,PasswordPolicies |
        Where-Object {
            $_.AccountEnabled -eq $true -and
            $_.UserType -eq 'Member' -and
            $_.UserPrincipalName -like "*$staffDomain" -and
            $_.UserPrincipalName -notlike "*$studentDomain" -and
            $_.UserPrincipalName -notlike '*#EXT#*'
        } |
        ForEach-Object {
            $lastChange = $_.LastPasswordChangeDateTime
            if (-not $lastChange) { $lastChange = $_.OnPremisesLastPasswordChangeDateTime }
            if ($lastChange) {
                $expiration = $lastChange.AddDays($maxAgeDays)
                $daysLeft = [math]::Floor(($expiration - $nowUtc).TotalDays)
                if ($daysLeft -ge -100 -and $daysLeft -le $SearchDays) {
                    [pscustomobject]@{
                        Name = $_.DisplayName
                        FirstName = $_.GivenName
                        LastName = $_.Surname
                        UPN = $_.UserPrincipalName
                        Email = $_.Mail
                        Expires = $expiration.ToString('yyyy-MM-dd')
                        DaysLeft = $daysLeft
                    }
                }
            }
        } |
        Sort-Object DaysLeft)

    if ($reportRows.Count -gt 0) {
        $reportRows | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8
    } else {
        '"Name","FirstName","LastName","UPN","Email","Expires","DaysLeft"' | Set-Content -LiteralPath $csvPath -Encoding UTF8
    }

    $rows = @(Import-Csv -LiteralPath $csvPath)
    $result = [ordered]@{ Success = $true; Count = $rows.Count; CsvPath = $csvPath; Error = $null }
    Write-Output "Report created with $($rows.Count) account(s)."
} catch {
    $result = [ordered]@{ Success = $false; Count = 0; CsvPath = $csvPath; Error = $_.Exception.Message }
    Write-Output "Run failed: $($_.Exception.Message)"
} finally {
    if ((Get-Command Get-MgContext -ErrorAction SilentlyContinue) -and
        (Get-MgContext -ErrorAction SilentlyContinue)) {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
    $result | ConvertTo-Json | Set-Content -LiteralPath $ResultPath -Encoding UTF8
}
