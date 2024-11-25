# PowerShell ��ũ��Ʈ ���ڵ� ����
$OutputEncoding = [System.Text.Encoding]::UTF8
chcp 65001 > $null

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# ���� �� �Լ� (�Ҽ��� �� �ڸ������� ��)
function Compare-Versions {
    param (
        [string]$Version1,
        [string]$Version2
    )
    
    $v1Parts = $Version1.Split('.')
    $v2Parts = $Version2.Split('.')
    
    # �Ҽ��� �� �ڸ������� ��
    for ($i = 0; $i -lt [Math]::Min(3, [Math]::Min($v1Parts.Length, $v2Parts.Length)); $i++) {
        if ($v1Parts[$i] -ne $v2Parts[$i]) {
            return $false
        }
    }
    return $true
}

# �ڵ����� ������ ���� ��û
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {  
    $arguments = "& '" + $myinvocation.mycommand.definition + "'"
    Start-Process powershell -Verb runAs -ArgumentList $arguments -WindowStyle Normal
    Break
}

# ���� ����
$host.UI.RawUI.WindowTitle = "Chrome & Selenium WebDriver ����ȭ ����"

# ȭ�� Ŭ����
Clear-Host

Write-Host "=============================================="
Write-Host "Chrome & Selenium WebDriver ����ȭ�� �����մϴ�."
Write-Host "=============================================="
Write-Host ""

# SeleniumBasic ��ġ ��� Ȯ��
$LocalPath = Join-Path $env:LOCALAPPDATA "SeleniumBasic"
$ProgramPath = "C:\Program Files\SeleniumBasic"

# ��ġ ��� ����
$DriverPath = $null

if (Test-Path $LocalPath) {
    $DriverPath = $LocalPath
    Write-Host "SeleniumBasic�� ����� ���� ��ο��� �߰ߵǾ����ϴ�." -ForegroundColor Green
    Write-Host "��ġ ���: $DriverPath" -ForegroundColor Cyan
}
elseif (Test-Path $ProgramPath) {
    $DriverPath = $ProgramPath
    Write-Host "SeleniumBasic�� Program Files ��ο��� �߰ߵǾ����ϴ�." -ForegroundColor Green
    Write-Host "��ġ ���: $DriverPath" -ForegroundColor Cyan
}
else {
    Write-Host "SeleniumBasic�� ��ġ�Ǿ� ���� �ʽ��ϴ�!" -ForegroundColor Red
    Write-Host "SeleniumBasic�� ���� ��ġ�ϼ���." -ForegroundColor Yellow
    Write-Host "���� ��� �� �ϳ��� SeleniumBasic�� ��ġ���ּ���:" -ForegroundColor Yellow
    Write-Host "1. $LocalPath" -ForegroundColor Yellow
    Write-Host "2. $ProgramPath" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Enter Ű�� ���� �����ϼ���..."
    Read-Host
    exit
}

$DriverExePath = Join-Path $DriverPath "chromedriver.exe"

# ũ�� ������ ���� Ȯ��
$ChromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
if (-not (Test-Path $ChromePath)) {
    Write-Host "Chrome�� ��ġ�Ǿ� ���� �ʽ��ϴ�." -ForegroundColor Red
    Write-Host "Enter Ű�� ���� �����ϼ���..."
    Read-Host
    exit
}

$ChromeVersion = (Get-Item $ChromePath).VersionInfo.FileVersion
$MajorVersion = $ChromeVersion.Split(".")[0]

Write-Host "���� Chrome ����: $ChromeVersion" -ForegroundColor Green
Write-Host ""

# ���� ChromeDriver ���μ��� ����
$chromeDriverProcesses = Get-Process -Name "chromedriver" -ErrorAction SilentlyContinue
if ($chromeDriverProcesses) {
    Write-Host "���� ChromeDriver ���μ��� ���� ��..." -ForegroundColor Yellow
    Stop-Process -Name "chromedriver" -Force
}

# ���� ChromeDriver ���� Ȯ��
$CurrentDriverVersion = ""
if (Test-Path $DriverExePath) {
    try {
        $CurrentDriverVersion = & $DriverExePath --version
        $CurrentDriverVersion = $CurrentDriverVersion -replace "ChromeDriver ", ""
        Write-Host "���� ��ġ�� ChromeDriver ����: $CurrentDriverVersion" -ForegroundColor Cyan
    }
    catch {
        Write-Host "���� ChromeDriver ���� Ȯ�� ����" -ForegroundColor Red
    }
}

# ChromeDriver ���� ���� ��������
$LatestVersion = $null
$url = "https://googlechromelabs.github.io/chrome-for-testing/LATEST_RELEASE_" + $MajorVersion

try {
    Write-Host "ChromeDriver ���� ���� Ȯ�� ��..." -ForegroundColor Yellow
    
    # �� ��û �� ���� ó��
    $response = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 10
    if ($response.StatusCode -eq 200) {
        # ���� ������ ���ڿ��� ��ȯ
        $responseString = [System.Text.Encoding]::UTF8.GetString($response.Content)
        $LatestVersion = $responseString.Trim()
        Write-Host "�ٿ�ε��� ChromeDriver ����: $LatestVersion" -ForegroundColor Green
    }
}
catch {
    Write-Host "ChromeDriver ���� ������ �������µ� �����߽��ϴ�." -ForegroundColor Red
    Write-Host "���� �� ����: $($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host "URL: $url" -ForegroundColor Yellow
    Write-Host "Enter Ű�� ���� �����ϼ���..."
    Read-Host
    exit
}

# ���� �� (�Ҽ��� �� �ڸ�������)
if (-not $CurrentDriverVersion -or -not (Compare-Versions -Version1 $CurrentDriverVersion -Version2 $LatestVersion)) {
    Write-Host ""
    Write-Host "���ο� ������ ChromeDriver �ٿ�ε� ��..." -ForegroundColor Yellow
    
    # �ӽ� zip ���� ���
    $TempZip = Join-Path $DriverPath "chromedriver_win32.zip"
    
    # ���� ���ϵ� ����
    Remove-Item -Path $DriverExePath -Force -ErrorAction SilentlyContinue
    Remove-Item -Path $TempZip -Force -ErrorAction SilentlyContinue
    
    # �� ���� �ٿ�ε�
    $downloadUrl = "https://edgedl.me.gvt1.com/edgedl/chrome/chrome-for-testing/$LatestVersion/win64/chromedriver-win64.zip"
    
    try {
        Write-Host "�ٿ�ε� ��... ($downloadUrl)" -ForegroundColor Yellow
        Invoke-WebRequest -Uri $downloadUrl -OutFile $TempZip -TimeoutSec 30
    }
    catch {
        Write-Host "ChromeDriver �ٿ�ε忡 �����߽��ϴ�." -ForegroundColor Red
        Write-Host "���� �� ����: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Host "�ٿ�ε� URL: $downloadUrl" -ForegroundColor Yellow
        Write-Host "Enter Ű�� ���� �����ϼ���..."
        Read-Host
        exit
    }
    
    # ���� ����
    Write-Host "ChromeDriver ���� ���� ��..." -ForegroundColor Yellow
    
    try {
        Expand-Archive -Path $TempZip -DestinationPath $DriverPath -Force
        
        # chromedriver-win64 ���� ó�� (���ο� ����)
        $chromedriverWin64Path = Join-Path $DriverPath "chromedriver-win64"
        if (Test-Path $chromedriverWin64Path) {
            Get-ChildItem -Path $chromedriverWin64Path | Move-Item -Destination $DriverPath -Force
            Remove-Item -Path $chromedriverWin64Path -Recurse -Force
        }
    }
    catch {
        Write-Host "���� ���� �� ������ �߻��߽��ϴ�." -ForegroundColor Red
        Write-Host $_.Exception.Message
        Write-Host "Enter Ű�� ���� �����ϼ���..."
        Read-Host
        exit
    }
    
    # �ӽ� ���� ����
    Remove-Item -Path $TempZip -Force
    
    Write-Host "ChromeDriver ������Ʈ �Ϸ�" -ForegroundColor Green
}
else {
    Write-Host "ChromeDriver�� �̹� ȣȯ�Ǵ� �����Դϴ�." -ForegroundColor Green
}

# ȯ�� ���� PATH�� ChromeDriver ��� �߰�
$CurrentPath = [Environment]::GetEnvironmentVariable("PATH", "Machine")
if ($CurrentPath -notlike "*$DriverPath*") {
    Write-Host ""
    Write-Host "ȯ�� ���� PATH�� ChromeDriver ��� �߰� ��..." -ForegroundColor Yellow
    [Environment]::SetEnvironmentVariable("PATH", "$CurrentPath;$DriverPath", "Machine")
    Write-Host "ȯ�� ���� PATH�� ChromeDriver ��� �߰� �Ϸ�" -ForegroundColor Green
}

Write-Host ""
Write-Host "=============================================="
Write-Host "������ ��� �Ϸ�Ǿ����ϴ�!" -ForegroundColor Green
Write-Host "Chrome ����: $ChromeVersion" -ForegroundColor Cyan
Write-Host "ChromeDriver ����: $LatestVersion" -ForegroundColor Cyan
Write-Host "ChromeDriver ���: $DriverPath" -ForegroundColor Cyan
Write-Host "=============================================="
Write-Host ""
Write-Host "Enter Ű�� ���� �����ϼ���..."
Read-Host