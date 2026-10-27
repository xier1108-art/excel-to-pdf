# 엑셀 일괄 PDF 변환기

엑셀(.xlsx/.xlsm/.xls/.xlsb) 및 CSV 파일을 드래그앤드롭으로 추가해 한 번에 PDF로 변환하는 Windows 데스크톱 프로그램입니다.

## 주요 기능
- 파일/폴더 드래그앤드롭 (폴더를 넣으면 하위 엑셀 파일을 재귀적으로 모두 추가)
- Microsoft Excel COM 자동화로 변환 (서식·인쇄영역 그대로 반영). Excel이 없는 PC에서는 LibreOffice로 자동 대체
- 배치 하나당 Excel 프로세스 1개만 재사용 → 다수 파일도 빠르게 처리
- 진행 상황 표시(성공 초록/실패 빨강), 개별 파일 실패 사유 확인, 변환 중 취소 가능
- 변환이 끝난 줄을 더블클릭하면 PDF가 바로 열림
- 출력 폴더 선택(원본 위치 또는 지정 폴더), 덮어쓰기 여부, 한 페이지 너비 맞춤, 완료 후 폴더 자동 열기
- 다른 폴더의 같은 이름 파일을 한 폴더로 모아도 `보고서.pdf`, `보고서(1).pdf`… 로 각각 저장 (서로 덮어쓰지 않음)
- 변환 도중 Excel이 죽으면 자동으로 다시 띄워서 남은 파일을 계속 변환
- 변환 중에 창을 닫으면 지금 파일까지만 마치고 Excel을 정리한 뒤 종료 (숨은 EXCEL.EXE가 남지 않음)
- LibreOffice 엔진은 전용 임시 프로필로 실행 — LibreOffice를 열어둔 상태에서도 변환됨
- 마지막 사용 옵션 자동 저장 (`%LOCALAPPDATA%\ExcelToPDF\settings.json`)

## 실행 (개발 환경)
```powershell
pip install -r requirements.txt
python app.py
```

## exe로 빌드하기
```powershell
./build.ps1
```
`dist\ExcelToPDF\` 폴더가 생성됩니다. **이 폴더 전체**를 배포하세요.

> ⚠️ 반드시 `--onedir`(폴더형)로 빌드합니다. `--onefile` 단일 exe는 실행할 때마다
> 임시 폴더에 압축을 풀어야 해서 실행마다 몇 초씩 느려집니다. 폴더형은 처음 실행 후
> 캐시가 남아 있어 훨씬 빠르게 켜집니다. 필요하면 바탕화면에 `ExcelToPDF.exe` 바로가기만
> 만들어 배포하면 됩니다.

## 요구 사항
- Windows + (Microsoft Excel 또는 LibreOffice 중 하나)
- Excel 매크로/보안 경고가 뜨는 파일은 COM에서 `ReadOnly`로 열기 때문에 대부분 무시되지만,
  특수한 보안 정책이 걸린 파일은 변환이 실패할 수 있습니다(실패 사유가 목록에 표시됩니다).

## 파일 구성
- `app.py` — GUI (tkinter, 드래그앤드롭은 tkinterdnd2 사용)
- `converter.py` — 변환 엔진 (Excel COM / LibreOffice)
- `build.ps1` — PyInstaller 빌드 스크립트
