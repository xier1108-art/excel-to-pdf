# 실행 파일 빌드 스크립트
# --onedir 로 빌드합니다 (--onefile은 실행할 때마다 임시폴더에 압축해제하느라 느려짐).
# 빌드 결과: dist\ExcelToPDF\ExcelToPDF.exe  (이 폴더 전체를 복사/배포하세요)

$ErrorActionPreference = "Stop"

python -m pip install -r requirements.txt

pyinstaller `
  --name ExcelToPDF `
  --onedir `
  --windowed `
  --noconfirm `
  --clean `
  app.py

Write-Host ""
Write-Host "빌드 완료: dist\ExcelToPDF\ExcelToPDF.exe" -ForegroundColor Green
Write-Host "배포 시 dist\ExcelToPDF 폴더 전체를 복사하세요 (exe만 단독으로 옮기면 실행되지 않습니다)." -ForegroundColor Yellow
