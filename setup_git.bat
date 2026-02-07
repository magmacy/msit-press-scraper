@echo off
setlocal enabledelayedexpansion

echo Git 초기화 및 설정을 시작합니다...

:: Git 설치 확인
git --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [오류] Git이 설치되어 있지 않거나 PATH에 없습니다.
    echo Git을 설치해주세요: https://git-scm.com/download/win
    pause
    exit /b
)

:: 사용자 설정 확인 및 입력
git config user.name >nul 2>&1
if %errorlevel% neq 0 (
    echo Git 사용자 이름이 설정되어 있지 않습니다.
    set /p GIT_USER="사용자 이름 입력 (예: Hong Gildong): "
    git config --global user.name "!GIT_USER!"
)

git config user.email >nul 2>&1
if %errorlevel% neq 0 (
    echo Git 이메일이 설정되어 있지 않습니다.
    set /p GIT_EMAIL="이메일 입력 (예: hong@example.com): "
    git config --global user.email "!GIT_EMAIL!"
)

:: 초기화 및 커밋
if not exist .git (
    git init
    echo Git 저장소가 초기화되었습니다.
)

git add .
git commit -m "Initial commit: MSIT Press Scraper with optimization and error handling"
if %errorlevel% neq 0 (
    echo 커밋할 변경사항이 없거나 오류가 발생했습니다.
)

:: 브랜치 이름 변경 및 원격 저장소 연결
git branch -M main
git remote remove origin >nul 2>&1
git remote add origin https://github.com/magmacy/msit-press-scraper.git

echo 원격 저장소로 푸시를 시도합니다...
git push -u origin main

if %errorlevel% neq 0 (
    echo [주의] 푸시 중 오류가 발생했습니다.
    echo 올바른 권한이 있거나 원격 저장소가 비어있는지 확인해주세요.
)

echo 완료되었습니다.
pause
