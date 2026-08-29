Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class WindowFocus {
    [DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
}
"@

$desktop = Split-Path -Parent $PSCommandPath
$workerPath = Join-Path $desktop 'Password Expiration Worker.ps1'
$resultPath = Join-Path $env:TEMP 'NPS-PasswordExpiration-result.json'
$authStatusPath = Join-Path $env:TEMP 'NPS-PasswordExpiration-status.txt'
$workerOutputPath = Join-Path $env:TEMP 'NPS-PasswordExpiration-output.txt'
$workerErrorPath = Join-Path $env:TEMP 'NPS-PasswordExpiration-error.txt'
$requiredStaffPath = Join-Path $desktop 'staff.csv'
$sanitizedStaffPath = Join-Path $desktop 'staff_sanitized.csv'
$noStaffMessage = 'No staff list found. No matches will be generated. Please add names to the staff.csv file.'

if (-not (Test-Path -LiteralPath $workerPath -PathType Leaf)) {
    [System.Windows.MessageBox]::Show(
        "Required application file is missing:`r`n$workerPath",
        'Password Expiration Report',
        'OK',
        'Error'
    ) | Out-Null
    exit 1
}

[xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Password Expiration Report" MinHeight="650" Width="900" SizeToContent="Height"
        WindowStartupLocation="CenterScreen" ResizeMode="CanMinimize"
        Background="#F4F7FB" FontFamily="Segoe UI">
  <Grid Margin="32">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <TextBlock Text="Password Expiration Report" FontSize="28" FontWeight="SemiBold" Foreground="#172033"/>
    <TextBlock Grid.Row="1" Margin="0,8,0,24" Text="Engineers: sign in with your ADM credentials to find staff accounts expiring soon."
               FontSize="14" Foreground="#5B6577" TextWrapping="Wrap"/>

    <Border Grid.Row="3" Background="White" CornerRadius="10" Padding="22" BorderBrush="#DDE3EC" BorderThickness="1">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="1"/>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="1"/>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <StackPanel>
          <TextBlock Text="TOTAL EXPIRED / EXPIRING" FontSize="11" Height="30" FontWeight="SemiBold" Foreground="#64748B"/>
          <TextBlock Name="TotalText" Text="-" FontSize="34" FontWeight="Bold" Foreground="#0F5FA8" Margin="0,3,0,0"/>
        </StackPanel>
        <Border Grid.Column="1" Background="#DDE3EC" Margin="12,0"/>
        <StackPanel Grid.Column="2">
          <TextBlock Text="SITE STAFF IMPORTED" FontSize="11" Height="30" FontWeight="SemiBold" Foreground="#64748B"/>
          <TextBlock Name="ImportedTotalText" Text="-" FontSize="34" FontWeight="Bold" Foreground="#18864B" Margin="0,3,0,0"/>
        </StackPanel>
        <Border Grid.Column="3" Background="#DDE3EC" Margin="12,0"/>
        <StackPanel Grid.Column="4">
          <TextBlock Text="EXPIRED / EXPIRING AT SITE(S)" FontSize="11" Height="30" FontWeight="SemiBold" Foreground="#64748B"/>
          <TextBlock Name="SchoolTotalText" Text="-" FontSize="34" FontWeight="Bold" Foreground="#7C3AED" Margin="0,3,0,0"/>
        </StackPanel>
        <StackPanel Grid.Column="5" Margin="18,0,0,0" VerticalAlignment="Center">
          <Button Name="RunButton" Content="Connect &amp; Run" Width="150" Height="42"
                  Background="#1261A0" Foreground="White" FontWeight="SemiBold"
                  BorderThickness="0" Cursor="Hand"/>
        </StackPanel>
      </Grid>
    </Border>

    <Border Grid.Row="2" Margin="0,0,0,16" Background="White" CornerRadius="8" Padding="14"
            BorderBrush="#DDE3EC" BorderThickness="1">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <TextBlock Name="LogText" Text="Ready. Select Connect &amp; Run to begin."
                   Foreground="#42526A" FontSize="13" TextWrapping="Wrap" Padding="0,0,16,0"/>
        <StackPanel Name="ActivityPanel" Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center"
                    Visibility="Collapsed">
          <Ellipse Name="Dot1" Width="9" Height="9" Fill="#1473C9" Opacity="1" Margin="3,0"/>
          <Ellipse Name="Dot2" Width="9" Height="9" Fill="#1473C9" Opacity="0.25" Margin="3,0"/>
          <Ellipse Name="Dot3" Width="9" Height="9" Fill="#1473C9" Opacity="0.25" Margin="3,0,12,0"/>
          <TextBlock Name="TimerText" Text="00:00" Foreground="#0F5FA8" FontFamily="Consolas" FontSize="14" FontWeight="SemiBold"/>
        </StackPanel>
      </Grid>
    </Border>

    <Border Grid.Row="4" Margin="0,16,0,0" Background="White" CornerRadius="8" Padding="14"
            BorderBrush="#DDE3EC" BorderThickness="1">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="12"/>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="12"/>
          <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        <StackPanel>
          <TextBlock Text="&lt;1 DAY UNTIL EXPIRATION" Foreground="#B42318" FontWeight="Bold" FontSize="11" Margin="0,0,0,7"/>
          <TextBox Name="Email1DayText" IsReadOnly="True" MinLines="2" Padding="8" TextWrapping="Wrap"
                   VerticalScrollBarVisibility="Disabled" Background="#FFF1F0" BorderBrush="#F04438"/>
        </StackPanel>
        <StackPanel Grid.Column="2">
          <TextBlock Text="&lt;3 DAYS UNTIL EXPIRATION" Foreground="#B54708" FontWeight="Bold" FontSize="11" Margin="0,0,0,7"/>
          <TextBox Name="Email3DayText" IsReadOnly="True" MinLines="2" Padding="8" TextWrapping="Wrap"
                   VerticalScrollBarVisibility="Disabled" Background="#FFFAEB" BorderBrush="#F79009"/>
        </StackPanel>
        <StackPanel Grid.Column="4">
          <TextBlock Text="&lt;7 DAYS UNTIL EXPIRATION" Foreground="#16794A" FontWeight="Bold" FontSize="11" Margin="0,0,0,7"/>
          <TextBox Name="Email7DayText" IsReadOnly="True" MinLines="2" Padding="8" TextWrapping="Wrap"
                   VerticalScrollBarVisibility="Disabled" Background="#ECFDF3" BorderBrush="#12B76A"/>
        </StackPanel>
      </Grid>
    </Border>
    <Border Grid.Row="5" Margin="0,16,0,0" Background="White" CornerRadius="8" Padding="14"
            BorderBrush="#DDE3EC" BorderThickness="1">
      <StackPanel>
        <TextBlock Text="ALREADY EXPIRED" Foreground="#B42318" FontWeight="Bold"
                   FontSize="11" Margin="0,0,0,7"/>
        <Grid>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="110"/>
          </Grid.ColumnDefinitions>
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>
          <TextBlock Text="EMAIL ADDRESS" Foreground="#667085" FontWeight="SemiBold" FontSize="10" Margin="8,0,0,4"/>
          <TextBlock Grid.Column="1" Text="DAYS EXPIRED" Foreground="#667085" FontWeight="SemiBold" FontSize="10" Margin="20,0,0,4"/>
          <TextBox Name="ExpiredEmailText" Grid.Row="1" IsReadOnly="True" MinLines="2" Padding="8"
                   TextWrapping="NoWrap" VerticalScrollBarVisibility="Disabled" FontFamily="Consolas" FontSize="12"
                   Background="#FFF1F0" BorderBrush="#F04438"/>
          <Border Grid.Row="1" Grid.Column="1" Margin="12,0,0,0" Padding="8" Background="#FFF1F0"
                  BorderBrush="#F04438" BorderThickness="1" IsHitTestVisible="False">
            <TextBlock Name="ExpiredDaysText" TextWrapping="NoWrap" FontFamily="Consolas" FontSize="12"/>
          </Border>
        </Grid>
      </StackPanel>
    </Border>
    <TextBlock Name="StatusText" Grid.Row="6" Visibility="Collapsed"/>
  </Grid>
</Window>
'@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)
$runButton = $window.FindName('RunButton')
$totalText = $window.FindName('TotalText')
$importedTotalText = $window.FindName('ImportedTotalText')
$schoolTotalText = $window.FindName('SchoolTotalText')
$email1DayText = $window.FindName('Email1DayText')
$email3DayText = $window.FindName('Email3DayText')
$email7DayText = $window.FindName('Email7DayText')
$expiredEmailText = $window.FindName('ExpiredEmailText')
$expiredDaysText = $window.FindName('ExpiredDaysText')
$logText = $window.FindName('LogText')
$statusText = $window.FindName('StatusText')
$activityPanel = $window.FindName('ActivityPanel')
$timerText = $window.FindName('TimerText')
$dots = @($window.FindName('Dot1'), $window.FindName('Dot2'), $window.FindName('Dot3'))
$script:process = $null
$script:checking = $false
$script:dotIndex = 0
$script:stopwatch = [Diagnostics.Stopwatch]::new()
$script:siteStaff = @()
$script:expiringRows = @()
$script:deviceCode = $null
$script:browserProcess = $null
$script:browserName = $null
$script:staffListError = $null
$script:staffImportSummary = $null
$script:timer = [Windows.Threading.DispatcherTimer]::new()
$script:timer.Interval = [TimeSpan]::FromMilliseconds(500)

function Open-PrivateDeviceLogin {
    $edgePaths = @(
        (Join-Path ${env:ProgramFiles(x86)} 'Microsoft\Edge\Application\msedge.exe'),
        (Join-Path $env:ProgramFiles 'Microsoft\Edge\Application\msedge.exe')
    )
    foreach ($browser in $edgePaths) {
        if (Test-Path -LiteralPath $browser) {
            $script:browserName = 'msedge'
            return (Start-Process -FilePath $browser -ArgumentList '--inprivate','--new-window','--start-minimized','https://microsoft.com/devicelogin' -WindowStyle Minimized -PassThru)
        }
    }
    $chromePaths = @(
        (Join-Path $env:ProgramFiles 'Google\Chrome\Application\chrome.exe'),
        (Join-Path ${env:ProgramFiles(x86)} 'Google\Chrome\Application\chrome.exe')
    )
    foreach ($browser in $chromePaths) {
        if (Test-Path -LiteralPath $browser) {
            $script:browserName = 'chrome'
            return (Start-Process -FilePath $browser -ArgumentList '--incognito','--new-window','--start-minimized','https://microsoft.com/devicelogin' -WindowStyle Minimized -PassThru)
        }
    }
    return $null
}

function Show-PrivateDeviceLogin {
    $candidates = @()
    if ($script:browserProcess) { $candidates += $script:browserProcess }
    if ($script:browserName) {
        $candidates += @(Get-Process -Name $script:browserName -ErrorAction SilentlyContinue | Sort-Object StartTime -Descending)
    }
    foreach ($candidate in $candidates) {
        try {
            $candidate.Refresh()
            if ($candidate.MainWindowHandle -ne [IntPtr]::Zero) {
                [void][WindowFocus]::ShowWindowAsync($candidate.MainWindowHandle, 9)
                [void][WindowFocus]::SetForegroundWindow($candidate.MainWindowHandle)
                return
            }
        } catch { }
    }
}

$script:timer.Add_Tick({
    if (-not $script:process) { return }
    if (-not $script:process.HasExited) {
        if (-not $script:deviceCode -and (Test-Path -LiteralPath $workerOutputPath)) {
            $workerOutput = Get-Content -LiteralPath $workerOutputPath -Raw -ErrorAction SilentlyContinue
            if ($workerOutput -match 'First-run setup: installing') {
                $logText.Text = 'First-run setup: installing the required Microsoft Graph components. This can take a few minutes.'
            }
            $codeMatch = [regex]::Match($workerOutput, '(?i)enter\s+(?:the\s+)?code\s+([A-Z0-9-]{6,16})')
            if (-not $codeMatch.Success) {
                $codeMatch = [regex]::Match($workerOutput, '(?i)code[:\s]+([A-Z0-9-]{6,16})')
            }
            if ($codeMatch.Success) {
                $script:deviceCode = $codeMatch.Groups[1].Value.ToUpperInvariant()
                try { [Windows.Clipboard]::SetText($script:deviceCode) } catch { }
                $logText.Text = "Device code: $($script:deviceCode)  (copied to clipboard)`r`nPaste it into the private Microsoft sign-in window."
                Show-PrivateDeviceLogin
            }
        }
        if (Test-Path -LiteralPath $authStatusPath) {
            $state = (Get-Content -LiteralPath $authStatusPath -Raw).Trim()
            if ($state -in @('AUTHENTICATED','CHECKING')) {
                if (-not $script:checking) {
                    $script:stopwatch.Restart()
                }
                $script:checking = $true
                $statusText.Foreground = '#18864B'
                $statusText.FontWeight = 'SemiBold'
                $statusText.Text = 'Authentication succeeded'
                $activityPanel.Visibility = 'Visible'
                for ($dot = 0; $dot -lt 3; $dot++) {
                    $dots[$dot].Opacity = if ($dot -eq $script:dotIndex) { 1 } else { 0.25 }
                }
                $script:dotIndex = ($script:dotIndex + 1) % 3
                $elapsed = $script:stopwatch.Elapsed
                $timerText.Text = '{0:00}:{1:00}' -f [math]::Floor($elapsed.TotalMinutes), $elapsed.Seconds
            }
        }
        return
    }
    $script:timer.Stop()
    $script:stopwatch.Stop()
    $activityPanel.Visibility = 'Collapsed'
    try {
        if (-not (Test-Path -LiteralPath $resultPath)) { throw 'The worker did not return a result.' }
        $result = Get-Content -LiteralPath $resultPath -Raw | ConvertFrom-Json
        if ($result.Success) {
            $totalText.Text = [string]$result.Count
            $script:expiringRows = @(Import-Csv -LiteralPath $result.CsvPath)
            Update-SchoolMatches
            $statusText.Text = "Complete. Report saved to $($result.CsvPath)"
            $statusText.Foreground = '#18864B'
            $logText.Text = "Finished: $($result.Count) expired or expiring account(s). Report: $($result.CsvPath)"
        } else {
            throw $result.Error
        }
    } catch {
        $totalText.Text = '!'
        $statusText.Text = 'Run failed. Review the error below.'
        $logText.Text = 'FAILED: ' + $_.Exception.Message
    } finally {
        $runButton.IsEnabled = $true
    }
})

$runButton.Add_Click({
    try {
        if ($script:process -and -not $script:process.HasExited) { return }
        if (-not (Import-RequiredStaffList)) {
            $totalText.Text = '-'
            $schoolTotalText.Text = '-'
            $email1DayText.Text = ''
            $email3DayText.Text = ''
            $email7DayText.Text = ''
            $expiredEmailText.Text = ''
            $expiredDaysText.Text = ''
            $logText.Text = $noStaffMessage
            [System.Windows.MessageBox]::Show(
                $noStaffMessage,
                'Password Expiration Report',
                'OK',
                'Warning'
            ) | Out-Null
            return
        }
        $staffImportSummary = $script:staffImportSummary
        $runButton.IsEnabled = $false
        $script:checking = $false
        $script:dotIndex = 0
        $script:deviceCode = $null
        $script:browserProcess = $null
        $script:browserName = $null
        $script:expiringRows = @()
        $schoolTotalText.Text = '-'
        $email1DayText.Text = ''
        $email3DayText.Text = ''
        $email7DayText.Text = ''
        $expiredEmailText.Text = ''
        $expiredDaysText.Text = ''
        $script:stopwatch.Restart()
        $totalText.Text = '-'
        $statusText.Text = 'Complete Microsoft sign-in in the private browser window.'
        $logText.Text = "$staffImportSummary`r`nWaiting for the Microsoft device code..."

        if (Test-Path -LiteralPath $resultPath) { Remove-Item -LiteralPath $resultPath -Force }
        if (Test-Path -LiteralPath $authStatusPath) { Remove-Item -LiteralPath $authStatusPath -Force }
        if (Test-Path -LiteralPath $workerOutputPath) { Remove-Item -LiteralPath $workerOutputPath -Force }
        if (Test-Path -LiteralPath $workerErrorPath) { Remove-Item -LiteralPath $workerErrorPath -Force }
        $arguments = '-NoProfile -ExecutionPolicy Bypass -File "{0}" -ResultPath "{1}" -StatusPath "{2}" -SearchDays 7' -f $workerPath, $resultPath, $authStatusPath
        $script:browserProcess = Open-PrivateDeviceLogin
        $script:process = Start-Process -FilePath 'powershell.exe' -ArgumentList $arguments -WindowStyle Hidden -PassThru `
            -RedirectStandardOutput $workerOutputPath -RedirectStandardError $workerErrorPath
        if (-not $script:browserProcess) {
            $logText.Text = 'No supported private browser was found. Open https://microsoft.com/devicelogin manually.'
        }
        $script:timer.Start()
    } catch {
        $script:stopwatch.Stop()
        $activityPanel.Visibility = 'Collapsed'
        $totalText.Text = '!'
        $statusText.Text = 'Could not start PowerShell.'
        $logText.Text = 'FAILED: ' + $_.Exception.Message
        $runButton.IsEnabled = $true
    }
})

function ConvertTo-NameKey([string]$Name) {
    if ([string]::IsNullOrWhiteSpace($Name)) { return '' }
    return ([regex]::Replace($Name.ToLowerInvariant(), '[^\p{L}\p{M}]', ''))
}

function Add-FirstLastNameKeys($KeySet, [string]$FirstName, [string]$LastName) {
    if ([string]::IsNullOrWhiteSpace($FirstName) -or [string]::IsNullOrWhiteSpace($LastName)) { return }
    $firstParts = @($FirstName -split '\s+' | Where-Object { $_ })
    $lastParts = @($LastName -split '\s+' | Where-Object { $_ })
    if ($firstParts.Count -eq 0 -or $lastParts.Count -eq 0) { return }

    [void]$KeySet.Add((ConvertTo-NameKey ($FirstName + $LastName)))
    [void]$KeySet.Add((ConvertTo-NameKey ($LastName + $FirstName)))
    [void]$KeySet.Add((ConvertTo-NameKey ($firstParts[0] + $lastParts[-1])))
    [void]$KeySet.Add((ConvertTo-NameKey ($lastParts[-1] + $firstParts[0])))
}

function Update-SchoolMatches {
    if ($script:siteStaff.Count -eq 0 -or $script:expiringRows.Count -eq 0) {
        $schoolTotalText.Text = if ($script:siteStaff.Count -gt 0 -and $totalText.Text -eq '0') { '0' } else { '-' }
        $email1DayText.Text = ''
        $email3DayText.Text = ''
        $email7DayText.Text = ''
        $expiredEmailText.Text = ''
        $expiredDaysText.Text = ''
        return
    }

    $siteNameKeys = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($staff in $script:siteStaff) {
        Add-FirstLastNameKeys -KeySet $siteNameKeys -FirstName ([string]$staff.FirstName) -LastName ([string]$staff.LastName)
    }

    $expiredAccounts = [System.Collections.Generic.Dictionary[string,int]]::new([StringComparer]::OrdinalIgnoreCase)
    $oneDayEmails = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $threeDayEmails = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $sevenDayEmails = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $matchedAccounts = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($account in $script:expiringRows) {
        $displayName = [string]$account.Name
        $candidateKeys = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        [void]$candidateKeys.Add((ConvertTo-NameKey $displayName))
        Add-FirstLastNameKeys -KeySet $candidateKeys -FirstName ([string]$account.FirstName) -LastName ([string]$account.LastName)
        if ($displayName -match '^(?<Last>[^,]+),\s*(?<First>.+)$') {
            Add-FirstLastNameKeys -KeySet $candidateKeys -FirstName $Matches.First -LastName $Matches.Last
        } else {
            $parts = @($displayName -split '\s+' | Where-Object { $_ })
            if ($parts.Count -ge 2) {
                Add-FirstLastNameKeys -KeySet $candidateKeys -FirstName $parts[0] -LastName $parts[-1]
            }
        }
        $isMatch = $false
        foreach ($key in $candidateKeys) {
            if ($siteNameKeys.Contains($key)) { $isMatch = $true; break }
        }
        if ($isMatch) {
            $accountKey = if ($account.UPN) { [string]$account.UPN } else { $displayName }
            [void]$matchedAccounts.Add($accountKey)
            $email = if ($account.Email) { [string]$account.Email } else { [string]$account.UPN }
            if ($email -match '^[^\s@]+@[^\s@]+\.[^\s@]+$') {
                $daysLeft = 999
                [void][int]::TryParse([string]$account.DaysLeft, [ref]$daysLeft)
                if ($daysLeft -lt 0) {
                    $daysExpired = -$daysLeft
                    if (-not $expiredAccounts.ContainsKey($email) -or $daysExpired -gt $expiredAccounts[$email]) {
                        $expiredAccounts[$email] = $daysExpired
                    }
                } elseif ($daysLeft -le 1) {
                    [void]$oneDayEmails.Add($email)
                } elseif ($daysLeft -le 3) {
                    [void]$threeDayEmails.Add($email)
                } elseif ($daysLeft -le 7) {
                    [void]$sevenDayEmails.Add($email)
                }
            }
        }
    }

    $schoolTotalText.Text = [string]$matchedAccounts.Count
    $expiredEntries = @($expiredAccounts.GetEnumerator() | Sort-Object `
        @{ Expression = { $_.Value }; Descending = $true },
        @{ Expression = { $_.Key }; Ascending = $true })
    $expiredEmailText.Text = (@($expiredEntries | ForEach-Object { $_.Key })) -join ";`r`n"
    $expiredDaysText.Text = (@($expiredEntries | ForEach-Object { [string]$_.Value })) -join "`r`n"
    $email1DayText.Text = (@($oneDayEmails) | Sort-Object) -join '; '
    $email3DayText.Text = (@($threeDayEmails) | Sort-Object) -join '; '
    $email7DayText.Text = (@($sevenDayEmails) | Sort-Object) -join '; '
}

function ConvertTo-SafeStaffName([string]$Value) {
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }

    $candidate = $Value.Trim()
    if ($candidate -match '^[=+@]' -or $candidate -match '^-[^A-Za-z]') { return $null }

    $candidate = [regex]::Replace($candidate, '\s+', ' ')
    $candidate = [regex]::Replace($candidate, "[^A-Za-z .'-]", '')
    $candidate = [regex]::Replace($candidate, ' {2,}', ' ').Trim()

    if ([string]::IsNullOrWhiteSpace($candidate)) { return $null }
    if ($candidate.Length -gt 80) { return $null }
    if (@('null', 'none', 'true', 'false', 'na') -contains $candidate.ToLowerInvariant()) { return $null }
    if ($candidate -notmatch "^[A-Za-z][A-Za-z .'-]*$") { return $null }

    $formattedWords = foreach ($word in @($candidate -split ' ' | Where-Object { $_ })) {
        $formattedSegments = foreach ($segment in @($word -split '-')) {
            if ($segment.Length -eq 0) {
                return $null
            } elseif ($segment.Length -eq 1) {
                $segment.ToUpperInvariant()
            } else {
                $segment.Substring(0, 1).ToUpperInvariant() + $segment.Substring(1).ToLowerInvariant()
            }
        }
        $formattedSegments -join '-'
    }
    return $formattedWords -join ' '
}

function ConvertTo-SanitizedStaffCsv([string]$SourcePath, [string]$DestinationPath) {
    $file = Get-Item -LiteralPath $SourcePath -ErrorAction Stop
    if ($file.Extension -ine '.csv') { throw 'Only .csv files are accepted.' }
    if ($file.Length -gt 131072) { throw 'The CSV is too large. Maximum file size is 128 KB.' }

    Add-Type -AssemblyName Microsoft.VisualBasic | Out-Null
    $parser = [Microsoft.VisualBasic.FileIO.TextFieldParser]::new($file.FullName)
    $safeRows = [System.Collections.Generic.List[object]]::new()
    $seenNames = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $skippedRows = 0
    $cleanedRows = 0
    $excludedRows = 0
    $duplicateRows = 0
    $lineCount = 0
    try {
        $parser.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
        $parser.SetDelimiters(',')
        $parser.HasFieldsEnclosedInQuotes = $true

        if ($parser.EndOfData) {
            $headers = @('First Name', 'Last Name')
        } else {
            try {
                $headers = @($parser.ReadFields())
                $lineCount++
            } catch [Microsoft.VisualBasic.FileIO.MalformedLineException] {
                throw 'The CSV header row is malformed.'
            }
        }

        $normalizedHeaders = @($headers | ForEach-Object { ([string]$_).Trim() })
        $firstNameIndex = -1
        $lastNameIndex = -1
        $locationIndex = -1
        for ($headerIndex = 0; $headerIndex -lt $normalizedHeaders.Count; $headerIndex++) {
            if ($firstNameIndex -lt 0 -and $normalizedHeaders[$headerIndex] -ieq 'First Name') {
                $firstNameIndex = $headerIndex
            }
            if ($lastNameIndex -lt 0 -and $normalizedHeaders[$headerIndex] -ieq 'Last Name') {
                $lastNameIndex = $headerIndex
            }
            if ($locationIndex -lt 0 -and $normalizedHeaders[$headerIndex] -ieq 'Location') {
                $locationIndex = $headerIndex
            }
        }

        if ($firstNameIndex -ge 0 -and $lastNameIndex -ge 0) {
            $inputFormat = 'NameColumns'
            $largestRequiredIndex = [Math]::Max($firstNameIndex, $lastNameIndex)
        } elseif ($locationIndex -ge 0) {
            $inputFormat = 'AssetLocation'
            $largestRequiredIndex = $locationIndex
        } else {
            throw 'The CSV must contain First Name and Last Name headers, or a Location column containing Staff: names.'
        }

        while (-not $parser.EndOfData) {
            if ($lineCount -ge 499) { throw 'The CSV must contain fewer than 500 lines.' }
            $lineCount++
            try {
                $fields = @($parser.ReadFields())
            } catch [Microsoft.VisualBasic.FileIO.MalformedLineException] {
                $skippedRows++
                continue
            }

            $wasCleaned = $false
            if ($inputFormat -eq 'AssetLocation') {
                $isTwoColumnNameRow = $fields.Count -ge 2
                if ($isTwoColumnNameRow -and $fields.Count -gt 2) {
                    for ($extraIndex = 2; $extraIndex -lt $fields.Count; $extraIndex++) {
                        if (-not [string]::IsNullOrWhiteSpace([string]$fields[$extraIndex])) {
                            $isTwoColumnNameRow = $false
                            break
                        }
                    }
                }

                $hasLocation = $fields.Count -gt $locationIndex
                $location = if ($hasLocation) { ([string]$fields[$locationIndex]).Trim() } else { '' }
                if ($location -match '(?i)^(Room|Student)\s*:') {
                    $excludedRows++
                    continue
                } elseif ($location -match '(?i)^Staff\s*:\s*(?<FullName>.+)$') {
                    $rawFullName = $Matches.FullName
                    $safeFullName = ConvertTo-SafeStaffName $rawFullName
                    if (-not $safeFullName) {
                        $skippedRows++
                        continue
                    }
                    $nameParts = @($safeFullName -split ' ' | Where-Object { $_ })
                    if ($nameParts.Count -lt 2) {
                        $skippedRows++
                        continue
                    }

                    $rawFirst = $nameParts[0]
                    $rawLast = $nameParts[-1]
                    $first = ConvertTo-SafeStaffName $rawFirst
                    $last = ConvertTo-SafeStaffName $rawLast
                    $wasCleaned = $safeFullName -cne $rawFullName
                } elseif ($isTwoColumnNameRow) {
                    $rawFirst = [string]$fields[0]
                    $rawLast = [string]$fields[1]
                    if ($rawFirst.Trim() -ieq 'First Name' -and $rawLast.Trim() -ieq 'Last Name') {
                        continue
                    }
                    $first = ConvertTo-SafeStaffName $rawFirst
                    $last = ConvertTo-SafeStaffName $rawLast
                    $wasCleaned = $first -and $last -and ($first -cne $rawFirst -or $last -cne $rawLast)
                } else {
                    $skippedRows++
                    continue
                }
            } else {
                if ($fields.Count -le $largestRequiredIndex) {
                    $skippedRows++
                    continue
                }
                $rawFirst = [string]$fields[$firstNameIndex]
                $rawLast = [string]$fields[$lastNameIndex]
                $first = ConvertTo-SafeStaffName $rawFirst
                $last = ConvertTo-SafeStaffName $rawLast
                $wasCleaned = $first -and $last -and ($first -cne $rawFirst -or $last -cne $rawLast)
            }

            if (-not $first -or -not $last) {
                $skippedRows++
                continue
            }
            if ($wasCleaned) {
                $cleanedRows++
            }

            $nameKey = $first + [char]0 + $last
            if (-not $seenNames.Add($nameKey)) {
                $duplicateRows++
                continue
            }

            [void]$safeRows.Add([pscustomobject]@{
                'First Name' = $first
                'Last Name' = $last
            })
        }
    } finally {
        $parser.Close()
    }

    $sortedRows = @($safeRows | Sort-Object @{ Expression = { $_.'Last Name' } }, @{ Expression = { $_.'First Name' } })
    $temporaryPath = $DestinationPath + '.tmp'
    try {
        if ($sortedRows.Count -gt 0) {
            $sortedRows | Export-Csv -LiteralPath $temporaryPath -NoTypeInformation -Encoding UTF8
        } else {
            '"First Name","Last Name"' | Set-Content -LiteralPath $temporaryPath -Encoding UTF8
        }
        Move-Item -LiteralPath $temporaryPath -Destination $DestinationPath -Force
    } finally {
        if (Test-Path -LiteralPath $temporaryPath) {
            Remove-Item -LiteralPath $temporaryPath -Force
        }
    }

    return [pscustomobject]@{
        ValidCount = $safeRows.Count
        SkippedCount = $skippedRows
        CleanedCount = $cleanedRows
        ExcludedCount = $excludedRows
        DuplicateCount = $duplicateRows
        InputFormat = $inputFormat
        Path = $DestinationPath
    }
}

function Set-SiteStaff([string]$Path) {
    $sanitization = ConvertTo-SanitizedStaffCsv -SourcePath $Path -DestinationPath $sanitizedStaffPath
    $script:siteStaff = @(Import-Csv -LiteralPath $sanitization.Path | ForEach-Object {
        [pscustomobject]@{
            FirstName = [string]$_.'First Name'
            LastName = [string]$_.'Last Name'
        }
    })
    $importedTotalText.Text = [string]$script:siteStaff.Count
    $script:staffImportSummary = "Using staff_sanitized.csv: $($sanitization.ValidCount) unique staff member(s); cleaned $($sanitization.CleanedCount) row(s); excluded $($sanitization.ExcludedCount) room/student row(s); removed $($sanitization.DuplicateCount) duplicate(s); skipped $($sanitization.SkippedCount) invalid row(s)."
    Update-SchoolMatches
}

function Import-RequiredStaffList {
    $script:staffListError = $null
    $script:staffImportSummary = $null
    if (-not (Test-Path -LiteralPath $requiredStaffPath -PathType Leaf)) {
        if (Test-Path -LiteralPath $sanitizedStaffPath) {
            Remove-Item -LiteralPath $sanitizedStaffPath -Force
        }
        $script:siteStaff = @()
        $importedTotalText.Text = '0'
        Update-SchoolMatches
        $logText.Text = $noStaffMessage
        return $false
    }

    if (Test-Path -LiteralPath $sanitizedStaffPath) {
        Remove-Item -LiteralPath $sanitizedStaffPath -Force
    }

    try {
        Set-SiteStaff -Path $requiredStaffPath
    } catch {
        $script:siteStaff = @()
        $script:staffListError = $_.Exception.Message
        $importedTotalText.Text = '!'
        Update-SchoolMatches
        $logText.Text = 'STAFF.CSV ERROR: ' + $script:staffListError
        return $false
    }

    if ($script:siteStaff.Count -eq 0) {
        $logText.Text = "$noStaffMessage`r`n$($script:staffImportSummary)"
        return $false
    }

    return $true
}

if (Import-RequiredStaffList) {
    $logText.Text = "$($script:staffImportSummary) Select Connect & Run to begin."
}

$window.Add_Closing({
    if ($script:process -and -not $script:process.HasExited) {
        $script:process.Kill()
    }
})

[void]$window.ShowDialog()
