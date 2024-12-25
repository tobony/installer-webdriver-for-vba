# installer-webdriver-for-selenium-basic

web driver를 설치합니다.
- chrome driver와
- MS edge driver를 설치합니다.

실행 후, 선택지 1번과 2번을 선택함에 따라 해당드라이버를 Selenium Basic 설치폴더에 다운받아 복사합니다.

<br><br>

<img width="744" alt="image" src="https://github.com/user-attachments/assets/d0fc7bcc-4604-4efe-b5bc-b754bddddfa9">


<br><br>

<img width="744" alt="image" src="https://github.com/user-attachments/assets/a41ca14f-a880-43ae-b1ee-fcfde6ff8cc5">


<br><br>

## make .exe file
# 1. PS2EXE 모듈 설치
Install-Module -Name ps2exe -Scope CurrentUser -Force

# 2. 스크립트를 EXE로 변환
Invoke-ps2exe -InputFile "./install-webdriver-seleniumBasic.ps1" -OutputFile "./install-webdriver.exe" -NoConsole -RequireAdmin
