# Gom file Bộ công cụ hóa đơn sang 1 thư mục mới để chép qua PC khác.
# Chạy bằng cách bấm đúp Gom_file_chuyen_may.bat (cùng thư mục). Chỉ ĐỌC và CHÉP, không sửa/xóa gì ở thư mục gốc.
#
# Tham số (không truyền thì script hỏi từng câu):
#   -Dich "D:\ChuyenMay\Bo_cong_cu_hoa_don"   thư mục đích (phải chưa có hoặc đang trống)
#   -KemDuLieu                                 chép thêm state\, downloads\, outputs\ (khi PC mới vẫn đọc các Gmail cũ)
#   -KemFileTongPetro                          chép thêm file tổng Petro ở ổ D:
#   -KhongHoi                                  không hỏi gì, dùng mặc định (không kèm dữ liệu, không kèm file tổng)
#   -DuAn "C:\...\gmail_api_project"           thư mục dự án (chứa app_config.json). Không truyền thì tự tìm:
#                                              thư mục chứa bo_chuyen_may, rồi Desktop\gmail_api_project
param(
    [string]$Dich = "",
    [string]$DuAn = "",
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

# ---------- 1. Thư mục dự án (thư mục có app_config.json) ----------
# Tìm lần lượt: -DuAn / thư mục chứa bo_chuyen_may / chính thư mục script / thư mục trên nữa /
# Desktop\gmail_api_project (Desktop nằm trong OneDrive cũng được). Không thấy thì hỏi đường dẫn.
function La-DuAn([string]$d) { return [bool]($d -and (Test-Path -LiteralPath (Join-Path $d "app_config.json"))) }
$Goc = $null
if ($DuAn) {
    $DuAn = $DuAn.Trim().Trim('"')
    if (-not (La-DuAn $DuAn)) { Write-Host "[LỖI] Không thấy app_config.json trong $DuAn (tham số -DuAn)" -ForegroundColor Red; exit 1 }
    $Goc = $DuAn
} else {
    $cha = Split-Path $PSScriptRoot -Parent
    $timO = @($cha, $PSScriptRoot, $(if ($cha) { Split-Path $cha -Parent }),
              $(if ([Environment]::GetFolderPath("Desktop")) { Join-Path ([Environment]::GetFolderPath("Desktop")) "gmail_api_project" }),
              $(if ($env:USERPROFILE) { Join-Path $env:USERPROFILE "Desktop\gmail_api_project" })) | Where-Object { $_ }
    $Goc = @($timO | Where-Object { La-DuAn $_ }) | Select-Object -First 1
    if (-not $Goc -and -not $KhongHoi) {
        Write-Host "Không tự tìm thấy thư mục dự án (đã tìm ở: $($timO -join '; '))." -ForegroundColor Yellow
        $tl = (Read-Host "Nhập đường dẫn thư mục dự án có app_config.json, vd C:\Users\...\Desktop\gmail_api_project").Trim().Trim('"')
        if (La-DuAn $tl) { $Goc = $tl }
    }
    if (-not $Goc) {
        Write-Host "[LỖI] Không tìm thấy thư mục dự án (thư mục có app_config.json)." -ForegroundColor Red
        Write-Host "      Chạy lại và nhập đường dẫn, hoặc: Gom_file_chuyen_may.bat -DuAn ""C:\...\gmail_api_project"""
        exit 1
    }
}
$Goc = (Resolve-Path -LiteralPath $Goc).Path.TrimEnd('\')
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

# ---------- 2b. Chương trình còn đang chạy? ----------
# Sổ theo dõi (state\*.db) ghi theo kiểu WAL: dữ liệu mới có thể còn nằm trong file *.db-wal cho tới khi
# chương trình tắt. Chép lúc đang chạy dễ được bản sổ thiếu / lệch -> nên tắt giao diện trước khi gom.
$dangChay = @()
$congMo = @()                                     # app.py dùng cổng 8765, bận thì lấy cổng kế tiếp (tới 8784)
try {                                             # chỉ hỏi cổng đang mở (hỏi cổng đóng trên Windows mất ~1-2 giây/cổng)
    $congMo = @(Get-NetTCPConnection -State Listen -LocalPort (8765..8784) -ErrorAction Stop |
                Select-Object -ExpandProperty LocalPort -Unique)
} catch { }
foreach ($port in $congMo) {
    try {
        $r = Invoke-RestMethod -Uri "http://127.0.0.1:$port/api/ping" -TimeoutSec 1
        if ($r.app -eq "hoadon") { $dangChay += "giao diện (http://127.0.0.1:$port)" }
    } catch { }
}
foreach ($d in @($chon.Dir, $Goc) | Select-Object -Unique) {
    $wal = @(Get-ChildItem (Join-Path $d "state") -Filter "*.db-wal" -ErrorAction SilentlyContinue | Where-Object { $_.Length -gt 0 })
    if ($wal.Count) { $dangChay += "sổ đang ghi dở: " + (($wal | ForEach-Object { $_.Name }) -join ", ") }
}
if ($dangChay.Count) {
    Write-Host ""
    Write-Host "[CẢNH BÁO] Bộ công cụ có vẻ đang chạy: $($dangChay -join '; ')" -ForegroundColor Yellow
    Write-Host "           Hãy tắt cửa sổ Chay_giao_dien.bat (và cửa sổ main.py nếu có) rồi gom lại, để sổ theo dõi chép đủ."
    if (-not (Hoi-CoKhong "Vẫn tiếp tục gom?" $false)) { exit 1 }
}

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
    # -KhongHoi: KHÔNG kèm dữ liệu (đúng như mô tả tham số ở đầu file; trước đây lại lấy mặc định "Có")
    $KemDuLieu = if ($KhongHoi) { $false } else { Hoi-CoKhong "PC mới vẫn đọc CÁC GMAIL CŨ (chép kèm sổ theo dõi state\, downloads\, outputs\)?" $true }
}
$fileTong = "D:\Làm việc\hđ đang xử lý\hoa_don_petrolimex_tong_hop_down.xlsx"
if (-not $PSBoundParameters.ContainsKey("KemFileTongPetro")) {
    $KemFileTongPetro = (Test-Path $fileTong) -and (Hoi-CoKhong "Chép kèm file tổng Petro ($fileTong)?" $KemDuLieu)
}

New-Item -ItemType Directory -Force $Dich | Out-Null
$ghiChu = New-Object System.Collections.Generic.List[string]
$canhBao = New-Object System.Collections.Generic.List[string]
function Moi-Nhat([string[]]$duongDan) {
    # File sửa gần nhất trong các đường dẫn (bỏ đường dẫn không tồn tại). Không có -> $null
    $co = @($duongDan | Where-Object { $_ -and (Test-Path -LiteralPath $_ -PathType Leaf) } |
            ForEach-Object { Get-Item -LiteralPath $_ } | Sort-Object FullName -Unique | Sort-Object LastWriteTime -Descending)
    if ($co.Count) { return $co[0].FullName } else { return $null }
}
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
foreach ($f in @("email_state.py", "Chay_giao_dien.bat")) {
    $p = if (Test-Path (Join-Path $chon.Dir $f)) { Join-Path $chon.Dir $f } else { Join-Path $Goc $f }
    if (Test-Path $p) { Chep $p $f } else { Write-Host "  [CẢNH BÁO] Không thấy $f" -ForegroundColor Yellow }
}
# app_config.json: lấy bản trong THƯ MỤC DỰ ÁN (vd gmail_api_project\app_config.json); không có mới lấy bản cạnh code.
$cfgDuAn, $cfgCode = (Join-Path $Goc "app_config.json"), (Join-Path $chon.Dir "app_config.json")
$cfg = if (Test-Path -LiteralPath $cfgDuAn) { $cfgDuAn } else { $cfgCode }
Chep $cfg "app_config.json"
if ($cfg -ne $cfgCode -and (Test-Path -LiteralPath $cfgCode) -and
    ((Get-FileHash -LiteralPath $cfg).Hash -ne (Get-FileHash -LiteralPath $cfgCode).Hash)) {
    $canhBao.Add("Có 2 bản app_config.json khác nhau: $cfg (đã chép) và $cfgCode. Kiểm tra danh sách profile trên PC mới.")
}

# v25: giao diện chạy file v25 đã lưu trong app_ui.json (v25_path) - có thể nằm NGOÀI thư mục code.
# Lấy bản sửa gần nhất trong: đường dẫn đã lưu / cạnh main.py / thư mục dự án.
$uiCfg = $null                                    # app_ui.json của PC cũ (không chép sang, chỉ đọc vài mục)
foreach ($ui in @((Join-Path $chon.Dir "app_ui.json"), (Join-Path $Goc "app_ui.json"))) {
    if (-not $uiCfg -and (Test-Path $ui)) {
        try { $uiCfg = Get-Content $ui -Raw -Encoding UTF8 | ConvertFrom-Json } catch { }
    }
}
$v25DaLuu = if ($uiCfg) { [string]$uiCfg.v25_path } else { "" }
$petroThuCong = if ($uiCfg) { [string]$uiCfg.petro_manual_dir } else { "" }   # thư mục mã Petro nhập thủ công đã đổi
$v25 = Moi-Nhat @($v25DaLuu, (Join-Path $chon.Dir "v25.py"), (Join-Path $Goc "v25.py"))
if ($v25) {
    Chep $v25 "v25.py"                            # luôn đặt tên v25.py cạnh main.py để PC mới tự tìm thấy
    if ($v25DaLuu -and (Test-Path -LiteralPath $v25DaLuu) -and ((Get-Item -LiteralPath $v25DaLuu).FullName -ne $v25)) {
        $canhBao.Add("Giao diện PC cũ đang chạy v25 ở $v25DaLuu, nhưng bản MỚI HƠN là $v25 -> đã chép bản mới hơn. " +
                     "Nếu còn dùng PC cũ: sửa đường dẫn File v25 ở trang Quét hóa đơn cho khớp.")
    }
} else { Write-Host "  [CẢNH BÁO] Không thấy v25.py" -ForegroundColor Yellow }

# Cai_dat.bat, requirements.txt, hướng dẫn: lấy bản MỚI NHẤT - khi cập nhật code, các file này thường được
# chép vào cạnh main.py, còn bản cũ vẫn nằm trong thư mục bo_chuyen_may.
foreach ($f in @("Cai_dat.bat", "requirements.txt", "HUONG_DAN_CHUYEN_MAY.txt")) {
    $p = Moi-Nhat @((Join-Path $chon.Dir $f), (Join-Path $Goc $f), (Join-Path $PSScriptRoot $f))
    if ($p) { Chep $p $f }
    elseif ($f -eq "requirements.txt") {
        $canhBao.Add("Không thấy requirements.txt -> Cai_dat.bat trên PC mới sẽ cài bản MỚI NHẤT của thư viện, có thể khác PC cũ.")
    } else { Write-Host "  [CẢNH BÁO] Không thấy $f" -ForegroundColor Yellow }
}
New-Item -ItemType Directory -Force (Join-Path $Dich "tokens") | Out-Null     # token mới tạo khi đăng nhập trên PC mới

# ---------- 5. Dữ liệu (tùy chọn) ----------
if ($KemDuLieu) {
    foreach ($d in @("state", "downloads", "outputs")) {
        # Chương trình ghi dữ liệu cạnh main.py; bản code nằm ở thư mục con (vd "New folder") thì dữ liệu có thể
        # ở đó thay vì ở thư mục dự án -> xét cả 2, lấy chỗ có file sửa gần nhất.
        $co = @(@((Join-Path $Goc $d), (Join-Path $chon.Dir $d)) | Select-Object -Unique | Where-Object { Test-Path -LiteralPath $_ })
        $p = $null
        if ($co.Count -eq 1) { $p = $co[0] }
        elseif ($co.Count -gt 1) {
            $moiNhat = { param($x) $f = Get-ChildItem -LiteralPath $x -Recurse -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1; if ($f) { $f.LastWriteTime } else { [datetime]::MinValue } }
            $p = @($co | Sort-Object { & $moiNhat $_ } -Descending)[0]
            $khac = @($co | Where-Object { $_ -ne $p })[0]
            $canhBao.Add("Có 2 thư mục $d\: đã chép $p (có file mới hơn), KHÔNG chép $khac. Nếu cần bản kia thì tự chép thêm.")
        }
        if ($p) {
            Chep $p $d
            # Chép nguyên thư mục, kể cả *.db-wal nếu có: file này có thể còn dữ liệu CHƯA ghi vào *.db
            # (vì vậy nên tắt chương trình trước khi gom, xem bước 2b).
        }
    }
}
# ---------- 5b. Thư mục mã Petrolimex nhập thủ công (app.py 2.0.4+) ----------
# Mặc định nằm trong downloads\Petro_thu_cong (đã chép ở trên nếu kèm dữ liệu). Nếu đã đổi sang thư mục khác
# thì đường dẫn đó nằm trong app_ui.json - file KHÔNG được chép, nên PC mới quay về downloads\Petro_thu_cong
# -> chép thư mục đã đổi vào đúng chỗ đó.
if ($petroThuCong) {
    $pt = if ([IO.Path]::IsPathRooted($petroThuCong)) { $petroThuCong } else { Join-Path $chon.Dir $petroThuCong }
    if (-not (Test-Path -LiteralPath $pt)) {
        $canhBao.Add("Không thấy thư mục mã Petrolimex nhập thủ công: $pt -> không chép.")
    } elseif ($KemDuLieu) {
        $macDinh = Join-Path $Dich "downloads\Petro_thu_cong"
        if (Test-Path -LiteralPath $macDinh) {       # thư mục mặc định cũ (dùng trước khi đổi) -> giữ lại, đổi tên
            Rename-Item -LiteralPath $macDinh ("Petro_thu_cong_cu_" + (Get-Date -Format "yyyyMMdd_HHmmss"))
        }
        Chep $pt "Petro_thu_cong" "downloads"
        $canhBao.Add("Thư mục mã Petrolimex nhập thủ công ($pt) đã chép vào downloads\Petro_thu_cong - PC mới dùng thư mục này.")
    } else {
        $canhBao.Add("CHƯA chép thư mục mã Petrolimex nhập thủ công ($pt) vì không kèm dữ liệu. " +
                     "Tự chép thư mục đó vào downloads\Petro_thu_cong trên PC mới (và chép state\processed_Petro_thu_cong.db vào state\).")
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
try {
    $root = [string]((Get-Content (Join-Path $Dich "app_config.json") -Raw -Encoding UTF8 | ConvertFrom-Json).tools.root)
    if ($root) {
        $canhBao.Add("app_config.json -> tools.root = $root  (chỉ cần khi dùng Tool_Convert: sửa thành thư mục project trên PC mới).")
    }
} catch { }
if ($canhBao.Count) { $ds += ""; $ds += "LƯU Ý:"; $ds += @($canhBao | ForEach-Object { "  - $_" }) }
$ds | Set-Content -Encoding UTF8 (Join-Path $Dich "DANH_SACH_FILE_DA_GOM.txt")

$tong = (Get-ChildItem $Dich -Recurse -File | Measure-Object Length -Sum)
Write-Host ""
Write-Host ("XONG: {0} file, {1:N1} MB -> {2}" -f $tong.Count, ($tong.Sum / 1MB), $Dich) -ForegroundColor Green
foreach ($c in $canhBao) { Write-Host "[LƯU Ý] $c" -ForegroundColor Yellow }
Write-Host ""
Write-Host "Việc tiếp theo:"
Write-Host "  1. Chép thư mục trên sang PC mới (USB / Google Drive / mạng nội bộ)."
Write-Host "  2. Tạo credentials.json mới trên Google Cloud, đặt cạnh main.py (xem HUONG_DAN_CHUYEN_MAY.txt, bước B)."
Write-Host "  3. Trên PC mới: bấm đúp Cai_dat.bat, rồi Chay_giao_dien.bat."
if ($KemFileTongPetro) { Write-Host "  4. Chép file tổng Petro vào ổ D: như ghi trong DANH_SACH_FILE_DA_GOM.txt." }
