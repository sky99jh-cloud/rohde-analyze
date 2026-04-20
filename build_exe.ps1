# ROHDE 송신기 로그 분석기 — Windows 단일 exe 빌드 (PyInstaller)
# 사용: 프로젝트 루트에서  .\build_exe.ps1

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

Write-Host "의존성 설치 (requirements.txt + requirements-build.txt)..." -ForegroundColor Cyan
python -m pip install -r requirements.txt
python -m pip install -r requirements-build.txt

Write-Host "PyInstaller 빌드 (ROHDE_Analyzer.spec)..." -ForegroundColor Cyan
python -m PyInstaller --noconfirm --clean ROHDE_Analyzer.spec

$exe = Join-Path $PSScriptRoot "dist\ROHDE_Analyzer.exe"
if (Test-Path $exe) {
    Write-Host "완료: $exe" -ForegroundColor Green
} else {
    Write-Host "실패: dist\ROHDE_Analyzer.exe 를 찾을 수 없습니다." -ForegroundColor Red
    exit 1
}
