# Gom file Bộ công cụ hóa đơn sang 1 thư mục mới để chép qua PC khác.
# Chạy bằng cách bấm đúp Gom_file_chuyen_may.bat (cùng thư mục). Chỉ ĐỌC và CHÉP, không sửa/xóa gì ở thư mục gốc.
#
# Tham số (không truyền thì script hỏi từng câu):
#   -Dich "D:\ChuyenMay\Bo_cong_cu_hoa_don"   thư mục đích (phải chưa có hoặc đang trống)
#   -KemDuLieu                                 chép thêm state\, downloads\, outputs\ (khi PC mới vẫn đọc các Gmail cũ)
#   -KemFileTongPetro                          chép thêm file tổng Petro ở ổ D:
#   -KhongHoi                                  không hỏi gì, dùng mặc định (không kèm dữ liệu, không kèm file tổng)
param(
    [string]$Dich = "",
    [switch]$KemDuLieu,
    [switch]$KemFileTongPetro,
    [switch]$KhongHoi
)
$ErrorActionPreference = "Stop"

function Hoi-CoKhong([string]$cauHoi, [bool]$macDinh) {
    if ($KhongHoi) { return $macDinh }
    $goiY = if ($macDinh) { "[C/k]" } else { "[c/K]" }
    $tl = Read-Host "$cauHoi $goiY"
    if ([string]::IsNullOrWhiteSpace($tl)) { return $macDinh }
    return $tl.Trim().ToLower().StartsWith("c")
}

# ---------- 1. Thư mục dự án (thư mục chứa bo_chuyen_may) ----------
$Goc = Split-Path $PSScriptRoot -Parent
if (-not (Test-Path (Join-Path $Goc "app_config.json"))) {
    Write-Host "[LỖI] Không thấy app_config.json trong $Goc" -ForegroundColor Red
    Write-Host "      Hãy để thư mục bo_chuyen_may nằm ngay trong thư mục dự án (cạnh app_config.json)."
    exit 1
}
Write-Host "Thư mục dự án : $Goc"

# ---------- 2. Chọn bản code mới nhất (APP_VERSION cao nhất có đủ 4 file) ----------
$can = @("main.py", "quet_email.py", "app.py", "Bo_cong_cu_hoa_don.html")
$ungVien = @($Goc) + @(Get-ChildItem $Goc -Directory -Recurse -Depth 2 | Where-Object { $_.FullName -notmatch "\\(downloads|outputs|state|tokens|logs|__pycache__)(\\|$)" } | ForEach-Object { $_.FullName })
$banCode = foreach ($d in $ungVien) {
    if (@($can | Where-Object { -not (Test-Path (Join-Path $d $_)) }).Count -gt 0) { continue }
    $m = Select-String -Path (Join-Path $d "app.py") -Pattern '^APP_VERSION = "([0-9.]+)"' | Select-Object -First 1
    if (-not $m) { continue }
    [pscustomobject]@{ Dir = $d; Ver = [version]$m.Matches[0].Groups[1].Value; Time = (Get-Item (Join-Path $d "app.py")).LastWriteTime }
}
$banCode = @($banCode | Sort-Object Ver, Time -Descending)
if (-not $banCode.Count) { Write-Host "[LỖI] Không tìm thấy bộ code (main.py, quet_email.py, app.py, Bo_cong_cu_hoa_don.html)." -ForegroundColor Red; exit 1 }
$chon = $banCode[0]
Write-Host ""
Write-Host "Các bản code tìm thấy:"
$i = 0
foreach ($b in $banCode) {
    $i++
    $ten = $b.Dir.Substring($Goc.Length).TrimStart('\')
    if (-not $ten) { $ten = "(thư mục gốc của dự án)" }
    Write-Host ("  {0}. bản {1,-6} {2}   (sửa lúc {3:dd/MM HH:mm})" -f $i, $b.Ver, $ten, $b.Time)
}
if (-not $KhongHoi) {
    $tl = Read-Host "Dùng bản số mấy? (Enter = 1: bản $($chon.Ver) mới nhất)"
    if ($tl -match '^\d+$' -and [int]$tl -ge 1 -and [int]$tl -le $banCode.Count) { $chon = $banCode[[int]$tl - 1] }
}
Write-Host "=> Dùng bản $($chon.Ver): $($chon.Dir)" -ForegroundColor Green

# ---------- 3. Thư mục đích ----------
if (-not $Dich) {
    $macDinhDich = Join-Path ([Environment]::GetFolderPath("Desktop")) "Bo_cong_cu_hoa_don_chuyen_may"
    $Dich = if ($KhongHoi) { $macDinhDich } else { Read-Host "Thư mục đích (Enter = $macDinhDich)" }
    if ([string]::IsNullOrWhiteSpace($Dich)) { $Dich = $macDinhDich }
}
$Dich = $Dich.Trim().Trim('"')
if ((Test-Path $Dich) -and @(Get-ChildItem $Dich -Force).Count -gt 0) {
    Write-Host "[LỖI] Thư mục đích đã có file: $Dich" -ForegroundColor Red
    Write-Host "      Chọn thư mục khác hoặc tự xóa thư mục đó trước (script không ghi đè để tránh mất dữ liệu)."
    exit 1
}
if (-not $PSBoundParameters.ContainsKey("KemDuLieu")) {
    $KemDuLieu = Hoi-CoKhong "PC mới vẫn đọc CÁC GMAIL CŨ (chép kèm sổ theo dõi state\, downloads\, outputs\)?" $true
}
$fileTong = "D:\Làm việc\hđ đang xử lý\hoa_don_petrolimex_tong_hop_down.xlsx"
if (-not $PSBoundParameters.ContainsKey("KemFileTongPetro")) {
    $KemFileTongPetro = (Test-Path $fileTong) -and (Hoi-CoKhong "Chép kèm file tổng Petro ($fileTong)?" $KemDuLieu)
}

New-Item -ItemType Directory -Force $Dich | Out-Null
$ghiChu = New-Object System.Collections.Generic.List[string]
function Chep([string]$nguon, [string]$ten, [string]$vao = "") {
    $dichDen = if ($vao) { Join-Path $Dich $vao } else { $Dich }
    New-Item -ItemType Directory -Force $dichDen | Out-Null
    Copy-Item -LiteralPath $nguon -Destination (Join-Path $dichDen $ten) -Recurse -Force
    $nhan = if ($vao) { "$vao\$ten" } else { $ten }
    $script:ghiChu.Add(("{0,-28} <- {1}" -f $nhan, $nguon))
}

# ---------- 4. Chép code + cấu hình ----------
Write-Host ""
Write-Host "Đang chép..."
foreach ($f in $can) { Chep (Join-Path $chon.Dir $f) $f }
foreach ($f in @("email_state.py", "v25.py", "Chay_giao_dien.bat", "app_config.json")) {
    $p = if (Test-Path (Join-Path $chon.Dir $f)) { Join-Path $chon.Dir $f } else { Join-Path $Goc $f }
    if (Test-Path $p) { Chep $p $f } else { Write-Host "  [CẢNH BÁO] Không thấy $f" -ForegroundColor Yellow }
}
foreach ($f in @("Cai_dat.bat", "HUONG_DAN_CHUYEN_MAY.txt")) {
    $p = Join-Path $PSScriptRoot $f
    if (Test-Path $p) { Chep $p $f }
}
New-Item -ItemType Directory -Force (Join-Path $Dich "tokens") | Out-Null     # token mới tạo khi đăng nhập trên PC mới

# ---------- 5. Dữ liệu (tùy chọn) ----------
if ($KemDuLieu) {
    foreach ($d in @("state", "downloads", "outputs")) {
        $p = Join-Path $Goc $d
        if (Test-Path $p) {
            Chep $p $d
            # state\*.csv và file tạm SQLite (-wal/-shm) không cần: chương trình tự tạo lại
        }
    }
}
if ($KemFileTongPetro) {
    if (Test-Path $fileTong) { Chep $fileTong (Split-Path $fileTong -Leaf) "_file_tong_Petro_chep_vao_o_D" }
    else { Write-Host "  [CẢNH BÁO] Không thấy file tổng Petro: $fileTong" -ForegroundColor Yellow }
}

# ---------- 6. Ghi danh sách + nhắc việc ----------
$ds = @(
    "DANH SÁCH FILE ĐÃ GOM - $(Get-Date -Format 'dd/MM/yyyy HH:mm')",
    "Bản code: $($chon.Ver)  ($($chon.Dir))",
    "Kèm dữ liệu (state, downloads, outputs): $(if ($KemDuLieu) {'CÓ'} else {'KHÔNG'})",
    "",
    "KHÔNG chép (cố ý): credentials.json, tokens\*.json (tạo mới trên PC mới), app_ui.json (tự tạo lại),",
    "                   logs\, code gốc\, ban_sua_*, __pycache__",
    ""
) + $ghiChu.ToArray()
if ($KemFileTongPetro) {
    $ds += ""
    $ds += "File tổng Petro: chép file trong thư mục _file_tong_Petro_chep_vao_o_D vào đúng"
    $ds += "  $(Split-Path $fileTong -Parent)  trên PC mới (hoặc khai báo 'petro_master' trong app_config.json)."
}
$ds | Set-Content -Encoding UTF8 (Join-Path $Dich "DANH_SACH_FILE_DA_GOM.txt")

$tong = (Get-ChildItem $Dich -Recurse -File | Measure-Object Length -Sum)
Write-Host ""
Write-Host ("XONG: {0} file, {1:N1} MB -> {2}" -f $tong.Count, ($tong.Sum / 1MB), $Dich) -ForegroundColor Green
Write-Host ""
Write-Host "Việc tiếp theo:"
Write-Host "  1. Chép thư mục trên sang PC mới (USB / Google Drive / mạng nội bộ)."
Write-Host "  2. Tạo credentials.json mới trên Google Cloud, đặt cạnh main.py (xem HUONG_DAN_CHUYEN_MAY.txt, bước B)."
Write-Host "  3. Trên PC mới: bấm đúp Cai_dat.bat, rồi Chay_giao_dien.bat."
if ($KemFileTongPetro) { Write-Host "  4. Chép file tổng Petro vào ổ D: như ghi trong DANH_SACH_FILE_DA_GOM.txt." }
