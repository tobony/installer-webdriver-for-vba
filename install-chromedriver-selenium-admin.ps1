# PowerShell 스크립트 인코딩 설정
$OutputEncoding = [System.Text.Encoding]::UTF8
chcp 65001 > $null

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# 버전 비교 함수 (소수점 세 자리까지만 비교)
function Compare-Versions {
    param (
        [string]$Version1,
        [string]$Version2
    )
    
    $v1Parts = $Version1.Split('.')
    $v2Parts = $Version2.Split('.')
    
    # 소수점 세 자리까지만 비교
    for ($i = 0; $i -lt [Math]::Min(3, [Math]::Min($v1Parts.Length, $v2Parts.Length)); $i++) {
        if ($v1Parts[$i] -ne $v2Parts[$i]) {
            return $false
        }
    }
    return $true
}

# 자동으로 관리자 권한 요청
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {  
    $arguments = "& '" + $myinvocation.mycommand.definition + "'"
    Start-Process powershell -Verb runAs -ArgumentList $arguments -WindowStyle Normal
    Break
}

# 제목 설정
$host.UI.RawUI.WindowTitle = "Chrome & Selenium WebDriver 동기화 도구"

# 화면 클리어
Clear-Host

Write-Host "=============================================="
Write-Host "Chrome & Selenium WebDriver 동기화를 시작합니다."
Write-Host "=============================================="
Write-Host ""

# SeleniumBasic 설치 경로 확인
$LocalPath = Join-Path $env:LOCALAPPDATA "SeleniumBasic"
$ProgramPath = "C:\Program Files\SeleniumBasic"

# 설치 경로 선택
$DriverPath = $null

if (Test-Path $LocalPath) {
    $DriverPath = $LocalPath
    Write-Host "SeleniumBasic이 사용자 로컬 경로에서 발견되었습니다." -ForegroundColor Green
    Write-Host "설치 경로: $DriverPath" -ForegroundColor Cyan
}
elseif (Test-Path $ProgramPath) {
    $DriverPath = $ProgramPath
    Write-Host "SeleniumBasic이 Program Files 경로에서 발견되었습니다." -ForegroundColor Green
    Write-Host "설치 경로: $DriverPath" -ForegroundColor Cyan
}
else {
    Write-Host "SeleniumBasic이 설치되어 있지 않습니다!" -ForegroundColor Red
    Write-Host "SeleniumBasic을 먼저 설치하세요." -ForegroundColor Yellow
    Write-Host "다음 경로 중 하나에 SeleniumBasic을 설치해주세요:" -ForegroundColor Yellow
    Write-Host "1. $LocalPath" -ForegroundColor Yellow
    Write-Host "2. $ProgramPath" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Enter 키를 눌러 종료하세요..."
    Read-Host
    exit
}

$DriverExePath = Join-Path $DriverPath "chromedriver.exe"

# 크롬 브라우저 버전 확인
$ChromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
if (-not (Test-Path $ChromePath)) {
    Write-Host "Chrome이 설치되어 있지 않습니다." -ForegroundColor Red
    Write-Host "Enter 키를 눌러 종료하세요..."
    Read-Host
    exit
}

$ChromeVersion = (Get-Item $ChromePath).VersionInfo.FileVersion
$MajorVersion = $ChromeVersion.Split(".")[0]

Write-Host "현재 Chrome 버전: $ChromeVersion" -ForegroundColor Green
Write-Host ""

# 기존 ChromeDriver 프로세스 종료
$chromeDriverProcesses = Get-Process -Name "chromedriver" -ErrorAction SilentlyContinue
if ($chromeDriverProcesses) {
    Write-Host "기존 ChromeDriver 프로세스 종료 중..." -ForegroundColor Yellow
    Stop-Process -Name "chromedriver" -Force
}

# 기존 ChromeDriver 버전 확인
$CurrentDriverVersion = ""
if (Test-Path $DriverExePath) {
    try {
        $CurrentDriverVersion = & $DriverExePath --version
        $CurrentDriverVersion = $CurrentDriverVersion -replace "ChromeDriver ", ""
        Write-Host "현재 설치된 ChromeDriver 버전: $CurrentDriverVersion" -ForegroundColor Cyan
    }
    catch {
        Write-Host "기존 ChromeDriver 버전 확인 실패" -ForegroundColor Red
    }
}

# ChromeDriver 버전 정보 가져오기
$LatestVersion = $null
$url = "https://googlechromelabs.github.io/chrome-for-testing/LATEST_RELEASE_" + $MajorVersion

try {
    Write-Host "ChromeDriver 버전 정보 확인 중..." -ForegroundColor Yellow
    
    # 웹 요청 및 응답 처리
    $response = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 10
    if ($response.StatusCode -eq 200) {
        # 응답 내용을 문자열로 변환
        $responseString = [System.Text.Encoding]::UTF8.GetString($response.Content)
        $LatestVersion = $responseString.Trim()
        Write-Host "다운로드할 ChromeDriver 버전: $LatestVersion" -ForegroundColor Green
    }
}
catch {
    Write-Host "ChromeDriver 버전 정보를 가져오는데 실패했습니다." -ForegroundColor Red
    Write-Host "에러 상세 정보: $($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host "URL: $url" -ForegroundColor Yellow
    Write-Host "Enter 키를 눌러 종료하세요..."
    Read-Host
    exit
}

# 버전 비교 (소수점 세 자리까지만)
if (-not $CurrentDriverVersion -or -not (Compare-Versions -Version1 $CurrentDriverVersion -Version2 $LatestVersion)) {
    Write-Host ""
    Write-Host "새로운 버전의 ChromeDriver 다운로드 중..." -ForegroundColor Yellow
    
    # 임시 zip 파일 경로
    $TempZip = Join-Path $DriverPath "chromedriver_win32.zip"
    
    # 이전 파일들 삭제
    Remove-Item -Path $DriverExePath -Force -ErrorAction SilentlyContinue
    Remove-Item -Path $TempZip -Force -ErrorAction SilentlyContinue
    
    # 새 버전 다운로드
    $downloadUrl = "https://edgedl.me.gvt1.com/edgedl/chrome/chrome-for-testing/$LatestVersion/win64/chromedriver-win64.zip"
    
    try {
        Write-Host "다운로드 중... ($downloadUrl)" -ForegroundColor Yellow
        Invoke-WebRequest -Uri $downloadUrl -OutFile $TempZip -TimeoutSec 30
    }
    catch {
        Write-Host "ChromeDriver 다운로드에 실패했습니다." -ForegroundColor Red
        Write-Host "에러 상세 정보: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Host "다운로드 URL: $downloadUrl" -ForegroundColor Yellow
        Write-Host "Enter 키를 눌러 종료하세요..."
        Read-Host
        exit
    }
    
    # 압축 해제
    Write-Host "ChromeDriver 압축 해제 중..." -ForegroundColor Yellow
    
    try {
        Expand-Archive -Path $TempZip -DestinationPath $DriverPath -Force
        
        # chromedriver-win64 폴더 처리 (새로운 구조)
        $chromedriverWin64Path = Join-Path $DriverPath "chromedriver-win64"
        if (Test-Path $chromedriverWin64Path) {
            Get-ChildItem -Path $chromedriverWin64Path | Move-Item -Destination $DriverPath -Force
            Remove-Item -Path $chromedriverWin64Path -Recurse -Force
        }
    }
    catch {
        Write-Host "압축 해제 중 오류가 발생했습니다." -ForegroundColor Red
        Write-Host $_.Exception.Message
        Write-Host "Enter 키를 눌러 종료하세요..."
        Read-Host
        exit
    }
    
    # 임시 파일 삭제
    Remove-Item -Path $TempZip -Force
    
    Write-Host "ChromeDriver 업데이트 완료" -ForegroundColor Green
}
else {
    Write-Host "ChromeDriver가 이미 호환되는 버전입니다." -ForegroundColor Green
}

# 환경 변수 PATH에 ChromeDriver 경로 추가
$CurrentPath = [Environment]::GetEnvironmentVariable("PATH", "Machine")
if ($CurrentPath -notlike "*$DriverPath*") {
    Write-Host ""
    Write-Host "환경 변수 PATH에 ChromeDriver 경로 추가 중..." -ForegroundColor Yellow
    [Environment]::SetEnvironmentVariable("PATH", "$CurrentPath;$DriverPath", "Machine")
    Write-Host "환경 변수 PATH에 ChromeDriver 경로 추가 완료" -ForegroundColor Green
}

Write-Host ""
Write-Host "=============================================="
Write-Host "설정이 모두 완료되었습니다!" -ForegroundColor Green
Write-Host "Chrome 버전: $ChromeVersion" -ForegroundColor Cyan
Write-Host "ChromeDriver 버전: $LatestVersion" -ForegroundColor Cyan
Write-Host "ChromeDriver 경로: $DriverPath" -ForegroundColor Cyan
Write-Host "=============================================="
Write-Host ""
Write-Host "Enter 키를 눌러 종료하세요..."
Read-Host