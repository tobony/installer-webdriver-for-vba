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
$host.UI.RawUI.WindowTitle = "Edge & Selenium WebDriver ����ȭ ����"

Clear-Host

Write-Host "=============================================="
Write-Host "Edge & Selenium WebDriver ����ȭ�� �����մϴ�."
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

$DriverExePath = Join-Path $DriverPath "msedgedriver.exe"

# Edge ������ ���� Ȯ��
$EdgePath = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
if (-not (Test-Path $EdgePath)) {
    Write-Host "Edge�� ��ġ�Ǿ� ���� �ʽ��ϴ�." -ForegroundColor Red
    Write-Host "Enter Ű�� ���� �����ϼ���..."
    Read-Host
    exit
}

$EdgeVersion = (Get-Item $EdgePath).VersionInfo.FileVersion
$MajorVersion = $EdgeVersion.Split(".")[0]

Write-Host "���� Edge ����: $EdgeVersion" -ForegroundColor Green
Write-Host ""

# ���� EdgeDriver ���μ��� ����
$edgeDriverProcesses = Get-Process -Name "msedgedriver" -ErrorAction SilentlyContinue
if ($edgeDriverProcesses) {
    Write-Host "���� EdgeDriver ���μ��� ���� ��..." -ForegroundColor Yellow
    Stop-Process -Name "msedgedriver" -Force
}

# ���� EdgeDriver ���� Ȯ��
$CurrentDriverVersion = ""
if (Test-Path $DriverExePath) {
    try {
        $CurrentDriverVersion = & $DriverExePath --version
        $CurrentDriverVersion = $CurrentDriverVersion -replace "MSEdgeDriver ", ""
        Write-Host "���� ��ġ�� EdgeDriver ����: $CurrentDriverVersion" -ForegroundColor Cyan
    }
    catch {
        Write-Host "���� EdgeDriver ���� Ȯ�� ����" -ForegroundColor Red
    }
}

# EdgeDriver ���� ���� ��������
$LatestVersion = $EdgeVersion  # Edge �������� ������ ���� ���
Write-Host "�ٿ�ε��� EdgeDriver ����: $LatestVersion" -ForegroundColor Green

try {
}
catch {
    Write-Host "EdgeDriver ���� ������ �������µ� �����߽��ϴ�." -ForegroundColor Red
    Write-Host "���� �� ����: $($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host "URL: $url" -ForegroundColor Yellow
    Write-Host "Enter Ű�� ���� �����ϼ���..."
    Read-Host
    exit
}

# ���� ��
if (-not $CurrentDriverVersion -or -not (Compare-Versions -Version1 $CurrentDriverVersion -Version2 $LatestVersion)) {
    Write-Host ""
    Write-Host "���ο� ������ EdgeDriver �ٿ�ε� ��..." -ForegroundColor Yellow
    
    $TempZip = Join-Path $DriverPath "edgedriver_win64.zip"
    
    Remove-Item -Path $DriverExePath -Force -ErrorAction SilentlyContinue
    Remove-Item -Path $TempZip -Force -ErrorAction SilentlyContinue
    
    # URL���� ����� ���� ���ڿ����� ���ʿ��� ���� ����
    $VersionForUrl = $LatestVersion -replace '[^\d\.]'
    $downloadUrl = "https://msedgedriver.azureedge.net/$VersionForUrl/edgedriver_win64.zip"
    
    try {
        Write-Host "�ٿ�ε� ��... ($downloadUrl)" -ForegroundColor Yellow
        Invoke-WebRequest -Uri $downloadUrl -OutFile $TempZip -TimeoutSec 30
    }
    catch {
        Write-Host "EdgeDriver �ٿ�ε忡 �����߽��ϴ�." -ForegroundColor Red
        Write-Host "���� �� ����: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Host "�ٿ�ε� URL: $downloadUrl" -ForegroundColor Yellow
        Write-Host "Enter Ű�� ���� �����ϼ���..."
        Read-Host
        exit
    }
    
    Write-Host "EdgeDriver ���� ���� ��..." -ForegroundColor Yellow
    
    try {
        Expand-Archive -Path $TempZip -DestinationPath $DriverPath -Force
    }
    catch {
        Write-Host "���� ���� �� ������ �߻��߽��ϴ�." -ForegroundColor Red
        Write-Host $_.Exception.Message
        Write-Host "Enter Ű�� ���� �����ϼ���..."
        Read-Host
        exit
    }
    
    Remove-Item -Path $TempZip -Force
    
    Write-Host "EdgeDriver ������Ʈ �Ϸ�" -ForegroundColor Green
}
else {
    Write-Host "EdgeDriver�� �̹� ȣȯ�Ǵ� �����Դϴ�." -ForegroundColor Green
}

# ȯ�� ���� PATH�� EdgeDriver ��� �߰�
$CurrentPath = [Environment]::GetEnvironmentVariable("PATH", "Machine")
if ($CurrentPath -notlike "*$DriverPath*") {
    Write-Host ""
    Write-Host "ȯ�� ���� PATH�� EdgeDriver ��� �߰� ��..." -ForegroundColor Yellow
    [Environment]::SetEnvironmentVariable("PATH", "$CurrentPath;$DriverPath", "Machine")
    Write-Host "ȯ�� ���� PATH�� EdgeDriver ��� �߰� �Ϸ�" -ForegroundColor Green
}

Write-Host ""
Write-Host "=============================================="
Write-Host "������ ��� �Ϸ�Ǿ����ϴ�!" -ForegroundColor Green
Write-Host "Edge ����: $EdgeVersion" -ForegroundColor Cyan
Write-Host "EdgeDriver ����: $LatestVersion" -ForegroundColor Cyan
Write-Host "EdgeDriver ���: $DriverPath" -ForegroundColor Cyan
Write-Host "=============================================="
Write-Host ""
Write-Host "Enter Ű�� ���� �����ϼ���..."
Read-Host