# PowerShell script encoding setting
$OutputEncoding = [System.Text.Encoding]::UTF8
chcp 65001 > $null

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create main form
$form = New-Object System.Windows.Forms.Form
$form.Text = 'WebDriver Synchronization Tool'
$form.Size = New-Object System.Drawing.Size(600,500)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false

# Log textbox
$logTextBox = New-Object System.Windows.Forms.TextBox
$logTextBox.Location = New-Object System.Drawing.Point(10,200)
$logTextBox.Size = New-Object System.Drawing.Size(565,200)
$logTextBox.Multiline = $true
$logTextBox.ScrollBars = 'Vertical'
$logTextBox.ReadOnly = $true
$form.Controls.Add($logTextBox)

# Progress bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10,410)
$progressBar.Size = New-Object System.Drawing.Size(565,20)
$form.Controls.Add($progressBar)

# Browser selection group box
$groupBox = New-Object System.Windows.Forms.GroupBox
$groupBox.Location = New-Object System.Drawing.Point(10,10)
$groupBox.Size = New-Object System.Drawing.Size(565,180)
$groupBox.Text = "WebDriver Selection"
$form.Controls.Add($groupBox)

# Description label
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,30)
$label.Size = New-Object System.Drawing.Size(530,20)
$label.Text = 'Select the WebDriver to install:'
$groupBox.Controls.Add($label)

# Chrome button
$chromeButton = New-Object System.Windows.Forms.Button
$chromeButton.Location = New-Object System.Drawing.Point(180,70)
$chromeButton.Size = New-Object System.Drawing.Size(200,30)
$chromeButton.Text = 'Chrome WebDriver'
$groupBox.Controls.Add($chromeButton)

# Edge button
$edgeButton = New-Object System.Windows.Forms.Button
$edgeButton.Location = New-Object System.Drawing.Point(180,120)
$edgeButton.Size = New-Object System.Drawing.Size(200,30)
$edgeButton.Text = 'Edge WebDriver'
$groupBox.Controls.Add($edgeButton)

# Write-Host replacement function
function Write-Log {
    param(
        [string]$Message,
        [string]$Color = "Black"
    )
    $logTextBox.AppendText("$Message`r`n")
    $logTextBox.Select($logTextBox.Text.Length, 0)
    $logTextBox.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

# Version comparison function
function Compare-Versions {
    param (
        [string]$Version1,
        [string]$Version2
    )
    
    $v1Parts = $Version1.Split('.')
    $v2Parts = $Version2.Split('.')
    
    for ($i = 0; $i -lt [Math]::Min(3, [Math]::Min($v1Parts.Length, $v2Parts.Length)); $i++) {
        if ($v1Parts[$i] -ne $v2Parts[$i]) {
            return $false
        }
    }
    return $true
}

# WebDriver installation function
function Install-WebDriver {
    param([string]$Choice)
    
    # Set variables based on selection
    switch ($Choice) {
        "1" {
            $browserName = "Chrome"
            $driverName = "chromedriver"
            $browserPath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
        }
        "2" {
            $browserName = "Edge"
            $driverName = "msedgedriver"
            $browserPath = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
        }
    }
    
    $progressBar.Value = 5
    Write-Log "=============================================="
    Write-Log "Starting $browserName & Selenium WebDriver synchronization."
    Write-Log "=============================================="

    # Check SeleniumBasic installation path
    $progressBar.Value = 10
    $LocalPath = Join-Path $env:LOCALAPPDATA "SeleniumBasic"
    $ProgramPath = "C:\Program Files\SeleniumBasic"
    $DriverPath = $null

    if (Test-Path $LocalPath) {
        $DriverPath = $LocalPath
        Write-Log "SeleniumBasic found in user local path."
        Write-Log "Installation path: $DriverPath"
    }
    elseif (Test-Path $ProgramPath) {
        $DriverPath = $ProgramPath
        Write-Log "SeleniumBasic found in Program Files path."
        Write-Log "Installation path: $DriverPath"
    }
    else {
        Write-Log "SeleniumBasic is not installed!"
        Write-Log "Please install SeleniumBasic in one of the following paths:"
        Write-Log "1. $LocalPath"
        Write-Log "2. $ProgramPath"
        return
    }

    $progressBar.Value = 20
    $DriverExePath = Join-Path $DriverPath "$driverName.exe"

    # Check browser version
    if (-not (Test-Path $browserPath)) {
        Write-Log "$browserName is not installed."
        return
    }

    $BrowserVersion = (Get-Item $browserPath).VersionInfo.FileVersion
    $MajorVersion = $BrowserVersion.Split(".")[0]

    Write-Log "Current $browserName version: $BrowserVersion"

    $progressBar.Value = 30
    # Terminate existing WebDriver processes
    $driverProcesses = Get-Process -Name $driverName -ErrorAction SilentlyContinue
    if ($driverProcesses) {
        Write-Log "Terminating existing $driverName processes..."
        Stop-Process -Name $driverName -Force
    }

    $progressBar.Value = 40
    # Check existing WebDriver version
    $CurrentDriverVersion = ""
    if (Test-Path $DriverExePath) {
        try {
            $CurrentDriverVersion = & $DriverExePath --version
            $CurrentDriverVersion = $CurrentDriverVersion -replace "$driverName ", ""
            Write-Log "Current installed $driverName version: $CurrentDriverVersion"
        }
        catch {
            Write-Log "Failed to check existing $driverName version"
        }
    }

    $progressBar.Value = 50
    # Get WebDriver version information
    $LatestVersion = $null
    if ($Choice -eq "1") {
        $url = "https://googlechromelabs.github.io/chrome-for-testing/LATEST_RELEASE_" + $MajorVersion
        try {
            Write-Log "Checking ChromeDriver version information..."
            $response = Invoke-WebRequest -Uri $url -UseBasicParsing
            if ($response.StatusCode -eq 200) {
                $LatestVersion = [System.Text.Encoding]::UTF8.GetString($response.Content).Trim()
                Write-Log "ChromeDriver version to download: $LatestVersion"
            }
        }
        catch {
            Write-Log "Failed to get ChromeDriver version information."
            Write-Log "Error details: $($_.Exception.Message)"
            Write-Log "URL: $url"
            return
        }
    }
    else {
        $LatestVersion = $BrowserVersion
        $VersionForUrl = $LatestVersion -replace '[^\d\.]'
        Write-Log "EdgeDriver version to download: $LatestVersion"
    }

    $progressBar.Value = 60
    # Version comparison and download
    if (-not $CurrentDriverVersion -or -not (Compare-Versions -Version1 $CurrentDriverVersion -Version2 $LatestVersion)) {
        Write-Log "Downloading new version of $driverName..."
        
        $TempZip = Join-Path $DriverPath "${driverName}_win64.zip"
        
        Remove-Item -Path $DriverExePath -Force -ErrorAction SilentlyContinue
        Remove-Item -Path $TempZip -Force -ErrorAction SilentlyContinue
        
        if ($Choice -eq "1") {
            $downloadUrl = "https://edgedl.me.gvt1.com/edgedl/chrome/chrome-for-testing/$LatestVersion/win64/chromedriver-win64.zip"
        }
        else {
            $downloadUrl = "https://msedgedriver.azureedge.net/$VersionForUrl/edgedriver_win64.zip"
        }
        
        $progressBar.Value = 70
        try {
            Write-Log "Downloading... ($downloadUrl)"
            Invoke-WebRequest -Uri $downloadUrl -OutFile $TempZip
        }
        catch {
            Write-Log "Failed to download $driverName."
            Write-Log "Error details: $($_.Exception.Message)"
            Write-Log "Download URL: $downloadUrl"
            return
        }
        
        $progressBar.Value = 80
        Write-Log "Extracting $driverName..."
        
        try {
            Expand-Archive -Path $TempZip -DestinationPath $DriverPath -Force
            
            if ($Choice -eq "1") {
                $driverWin64Path = Join-Path $DriverPath "chromedriver-win64"
                if (Test-Path $driverWin64Path) {
                    Get-ChildItem -Path $driverWin64Path | Move-Item -Destination $DriverPath -Force
                    Remove-Item -Path $driverWin64Path -Recurse -Force
                }
            }
        }
        catch {
            Write-Log "Error occurred during extraction."
            Write-Log $_.Exception.Message
            return
        }
        
        Remove-Item -Path $TempZip -Force
        Write-Log "$driverName update completed"
    }
    else {
        Write-Log "$driverName is already at a compatible version."
    }

    $progressBar.Value = 90
    # Add WebDriver path to PATH environment variable
    $CurrentPath = [Environment]::GetEnvironmentVariable("PATH", "Machine")
    if ($CurrentPath -notlike "*$DriverPath*") {
        Write-Log "Adding $driverName path to PATH environment variable..."
        [Environment]::SetEnvironmentVariable("PATH", "$CurrentPath;$DriverPath", "Machine")
        Write-Log "Successfully added $driverName path to PATH environment variable"
    }

    $progressBar.Value = 100
    Write-Log "=============================================="
    Write-Log "Setup completed successfully!"
    Write-Log "$browserName version: $BrowserVersion"
    Write-Log "$driverName version: $LatestVersion"
    Write-Log "$driverName path: $DriverPath"
    Write-Log "=============================================="
    
    # Re-enable buttons
    $chromeButton.Enabled = $true
    $edgeButton.Enabled = $true
}

# Button click events
$chromeButton.Add_Click({
    $chromeButton.Enabled = $false
    $edgeButton.Enabled = $false
    Install-WebDriver -Choice "1"
})

$edgeButton.Add_Click({
    $chromeButton.Enabled = $false
    $edgeButton.Enabled = $false
    Install-WebDriver -Choice "2"
})

# Check and request admin privileges
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
$isAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    $result = [System.Windows.Forms.MessageBox]::Show(
        "This tool requires administrator privileges. Would you like to restart with administrator privileges?",
        "Administrator Privileges Required",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question)
    
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        # Restart with admin privileges
        $processStartInfo = New-Object System.Diagnostics.ProcessStartInfo
        $processStartInfo.FileName = "powershell.exe"
        $processStartInfo.Arguments = "-File `"$($myinvocation.mycommand.definition)`""
        $processStartInfo.Verb = "runas"
        [System.Diagnostics.Process]::Start($processStartInfo)
    }
    exit
}

[System.Windows.Forms.Application]::Run($form)