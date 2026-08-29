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
        Title="Password Expiration Report" Height="690" Width="900"
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
          <TextBlock Text="TOTAL ACCOUNTS EXPIRING" FontSize="11" Height="30" FontWeight="SemiBold" Foreground="#64748B"/>
          <TextBlock Name="TotalText" Text="-" FontSize="34" FontWeight="Bold" Foreground="#0F5FA8" Margin="0,3,0,0"/>
        </StackPanel>
        <Border Grid.Column="1" Background="#DDE3EC" Margin="12,0"/>
        <StackPanel Grid.Column="2">
          <TextBlock Text="SITE STAFF IMPORTED" FontSize="11" Height="30" FontWeight="SemiBold" Foreground="#64748B"/>
          <TextBlock Name="ImportedTotalText" Text="-" FontSize="34" FontWeight="Bold" Foreground="#18864B" Margin="0,3,0,0"/>
        </StackPanel>
        <Border Grid.Column="3" Background="#DDE3EC" Margin="12,0"/>
        <StackPanel Grid.Column="4">
          <TextBlock Text="EXPIRING AT YOUR SITE(S)" FontSize="11" Height="30" FontWeight="SemiBold" Foreground="#64748B"/>
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
          <TextBox Name="Email1DayText" IsReadOnly="True" Height="90" Padding="8" TextWrapping="Wrap"
                   VerticalScrollBarVisibility="Auto" Background="#FFF1F0" BorderBrush="#F04438"/>
        </StackPanel>
        <StackPanel Grid.Column="2">
          <TextBlock Text="&lt;3 DAYS UNTIL EXPIRATION" Foreground="#B54708" FontWeight="Bold" FontSize="11" Margin="0,0,0,7"/>
          <TextBox Name="Email3DayText" IsReadOnly="True" Height="90" Padding="8" TextWrapping="Wrap"
                   VerticalScrollBarVisibility="Auto" Background="#FFFAEB" BorderBrush="#F79009"/>
        </StackPanel>
        <StackPanel Grid.Column="4">
          <TextBlock Text="&lt;7 DAYS UNTIL EXPIRATION" Foreground="#16794A" FontWeight="Bold" FontSize="11" Margin="0,0,0,7"/>
          <TextBox Name="Email7DayText" IsReadOnly="True" Height="90" Padding="8" TextWrapping="Wrap"
                   VerticalScrollBarVisibility="Auto" Background="#ECFDF3" BorderBrush="#12B76A"/>
        </StackPanel>
      </Grid>
    </Border>
    <TextBlock Name="StatusText" Visibility="Collapsed"/>
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
            $logText.Text = "Finished: $($result.Count) account(s) expiring. Report: $($result.CsvPath)"
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
        $script:stopwatch.Restart()
        $totalText.Text = '-'
        $statusText.Text = 'Complete Microsoft sign-in in the private browser window.'
        $logText.Text = 'Waiting for the Microsoft device code...'

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

function Update-SchoolMatches {
    if ($script:siteStaff.Count -eq 0 -or $script:expiringRows.Count -eq 0) {
        $schoolTotalText.Text = if ($script:siteStaff.Count -gt 0 -and $totalText.Text -eq '0') { '0' } else { '-' }
        $email1DayText.Text = ''
        $email3DayText.Text = ''
        $email7DayText.Text = ''
        return
    }

    $siteNameKeys = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($staff in $script:siteStaff) {
        [void]$siteNameKeys.Add((ConvertTo-NameKey ($staff.FirstName + $staff.LastName)))
        [void]$siteNameKeys.Add((ConvertTo-NameKey ($staff.LastName + $staff.FirstName)))
    }

    $oneDayEmails = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $threeDayEmails = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $sevenDayEmails = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $matchedAccounts = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($account in $script:expiringRows) {
        $displayName = [string]$account.Name
        $candidateKeys = [System.Collections.Generic.List[string]]::new()
        [void]$candidateKeys.Add((ConvertTo-NameKey $displayName))
        $parts = @($displayName -split '[,\s]+' | Where-Object { $_ })
        if ($parts.Count -ge 2) {
            [void]$candidateKeys.Add((ConvertTo-NameKey ($parts[0] + $parts[-1])))
            [void]$candidateKeys.Add((ConvertTo-NameKey ($parts[-1] + $parts[0])))
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
                if ($daysLeft -ge 0 -and $daysLeft -le 1) {
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
    $email1DayText.Text = (@($oneDayEmails) | Sort-Object) -join '; '
    $email3DayText.Text = (@($threeDayEmails) | Sort-Object) -join '; '
    $email7DayText.Text = (@($sevenDayEmails) | Sort-Object) -join '; '
}

function Read-SafeSiteStaff([string]$Path) {
    $file = Get-Item -LiteralPath $Path -ErrorAction Stop
    if ($file.Extension -ine '.csv') { throw 'Only .csv files are accepted.' }
    if ($file.Length -gt 131072) { throw 'The CSV is too large. Maximum file size is 128 KB.' }

    Add-Type -AssemblyName Microsoft.VisualBasic
    $parser = [Microsoft.VisualBasic.FileIO.TextFieldParser]::new($file.FullName)
    try {
        $parser.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
        $parser.SetDelimiters(',')
        $parser.HasFieldsEnclosedInQuotes = $true
        $records = [System.Collections.Generic.List[object]]::new()
        while (-not $parser.EndOfData) {
            $fields = @($parser.ReadFields())
            if ($fields.Count -ne 2) { throw 'Every row must contain exactly two columns.' }
            [void]$records.Add($fields)
            if ($records.Count -ge 500) { throw 'The CSV must contain fewer than 500 lines.' }
        }
    } finally {
        $parser.Close()
    }

    if ($records.Count -lt 2) { throw 'The CSV must contain a header and at least one staff member.' }
    if ($records[0][0].Trim() -ine 'First Name' -or $records[0][1].Trim() -ine 'Last Name') {
        throw 'The two headers must be First Name and Last Name.'
    }

    $safeRows = [System.Collections.Generic.List[object]]::new()
    for ($index = 1; $index -lt $records.Count; $index++) {
        $first = $records[$index][0].Trim()
        $last = $records[$index][1].Trim()
        foreach ($value in @($first, $last)) {
            if ([string]::IsNullOrWhiteSpace($value)) { throw "Blank name found on line $($index + 1)." }
            if ($value.Length -gt 80) { throw "Name longer than 80 characters found on line $($index + 1)." }
            if ($value -match '^[=+@]' -or $value -match '^-[^A-Za-z]') { throw "Unsafe spreadsheet formula found on line $($index + 1)." }
            if ($value -notmatch "^[\p{L}\p{M}][\p{L}\p{M} .'-]*$") { throw "Invalid characters found on line $($index + 1)." }
        }
        [void]$safeRows.Add([pscustomobject]@{ FirstName = $first; LastName = $last })
    }
    return @($safeRows)
}

function Set-SiteStaff([string]$Path) {
    $script:siteStaff = @(Read-SafeSiteStaff -Path $Path)
    $importedTotalText.Text = [string]$script:siteStaff.Count
    Update-SchoolMatches
}

try {
    if (-not (Test-Path -LiteralPath $requiredStaffPath -PathType Leaf)) {
        throw "Required file not found: $requiredStaffPath"
    }
    Set-SiteStaff -Path $requiredStaffPath
    $logText.Text = "Loaded staff.csv: $($script:siteStaff.Count) staff member(s). Select Connect & Run to begin."
} catch {
    $script:siteStaff = @()
    $importedTotalText.Text = '!'
    $logText.Text = 'STAFF.CSV ERROR: ' + $_.Exception.Message
}

$window.Add_Closing({
    if ($script:process -and -not $script:process.HasExited) {
        $script:process.Kill()
    }
})

[void]$window.ShowDialog()
