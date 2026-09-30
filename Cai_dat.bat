@echo off
rem Cai thu vien Python + Chromium cho Bo cong cu hoa don. Bam dup de chay, chay lai nhieu lan cung duoc.
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================================
echo  CAI DAT BO CONG CU HOA DON
echo ============================================================
echo.

rem --- 1. Tim Python (python hoac py) ---
set "PY="
python --version >nul 2>nul && set "PY=python"
if not defined PY py -3 --version >nul 2>nul && set "PY=py -3"
if not defined PY goto :nopython
echo [1/5] Python:
%PY% --version
echo.

rem --- 2. Thu vien chinh ---
echo [2/5] Cai thu vien Python (mat vai phut)...
%PY% -m pip install --upgrade pip
if exist "requirements.txt" (
  rem Phien ban co dinh trong requirements.txt -> may nao cai cung giong nhau
  %PY% -m pip install -r requirements.txt
) else (
  echo [CANH BAO] Khong thay requirements.txt, cai ban moi nhat cua cac thu vien.
  %PY% -m pip install google-api-python-client google-auth-oauthlib google-auth playwright beautifulsoup4 openpyxl pandas pdfplumber pytesseract pdf2image pillow
)
if errorlevel 1 goto :piperr
echo.

rem --- 3. ddddocr (tu giai CAPTCHA Petrolimex) - cai rieng de loi (neu co) khong chan cac buoc khac ---
echo [3/5] Cai ddddocr (tu giai CAPTCHA Petrolimex)...
%PY% -m pip install ddddocr==1.6.1
if errorlevel 1 (
  echo [CANH BAO] Chua cai duoc ddddocr. Van chay duoc, nhung CAPTCHA Petrolimex phai nhap tay.
)
echo.

rem --- 4. Trinh duyet Chromium cho Playwright ---
echo [4/5] Cai trinh duyet Chromium cho Playwright...
%PY% -m playwright install chromium
if errorlevel 1 goto :pwerr
echo.

rem --- 5. Kiem tra ---
echo [5/5] Kiem tra...
%PY% -c "import googleapiclient, google_auth_oauthlib, playwright, bs4, openpyxl, pandas, pdfplumber; print('   Thu vien chinh: OK')"
%PY% -c "import ddddocr; print('   ddddocr: OK')" 2>nul || echo    ddddocr: CHUA CO (CAPTCHA Petrolimex phai nhap tay)
where tesseract >nul 2>nul && (echo    Tesseract OCR: OK) || (echo    Tesseract OCR: CHUA CO - chi can neu v25 quet PDF dang anh)
where pdftoppm >nul 2>nul && (echo    Poppler: OK) || (echo    Poppler: CHUA CO - chi can neu v25 quet PDF dang anh)
if exist "credentials.json" (echo    credentials.json: OK) else (echo    credentials.json: CHUA CO - tao tren Google Cloud, dat canh main.py)
if exist "main.py" (echo    main.py: OK) else (echo    main.py: CHUA THAY - hay dat Cai_dat.bat cung thu muc voi main.py)
echo.
echo XONG. Buoc tiep theo: bam dup Chay_giao_dien.bat
echo.
pause
exit /b 0

:nopython
echo [LOI] Chua cai Python.
echo   Tai Python tai https://www.python.org/downloads/ va khi cai nho tick "Add Python to PATH".
echo   Cai xong, dong cua so nay roi bam dup lai Cai_dat.bat.
echo.
pause
exit /b 1

:piperr
echo.
echo [LOI] Cai thu vien khong thanh cong. Kiem tra ket noi Internet roi chay lai Cai_dat.bat.
echo   Neu van loi, chup man hinh cua so nay gui lai de xem.
pause
exit /b 1

:pwerr
echo.
echo [LOI] Chua cai duoc Chromium cho Playwright. Kiem tra Internet roi chay lai Cai_dat.bat.
pause
exit /b 1
