<#=============================================================================================
 Windows 11 Edition Converter
 Converts the running Windows 11 edition in-place using only tools that ship with Windows:
   - DISM /Online /Get-TargetEditions : lists the editions this installation can move to
   - changepk.exe /ProductKey         : changes the edition
 Editions that DISM does not list are shown in red and go through Invoke-SkuConversion
 (the SKU file method). That function is a stub: write your own logic in it.
 This tool does not activate Windows. After converting, activate it with your own license.

 ENCODING: save this file as UTF-8 WITHOUT BOM, and keep every non-ASCII character inside the
 $uiText here-string below. See section 0 for why.
=============================================================================================#>

# 0. Vietnamese UI text
# Why the text lives in one here-string:
#  - "irm ... | iex" breaks when the file starts with a BOM, so the file has no BOM.
#  - Windows PowerShell 5.1 reads a BOM-less file with the ANSI code page. The Vietnamese text is
#    garbled then, but a single-quoted here-string still parses, and the block is re-read as UTF-8.
# Format: Key = text. {0}, {1}... are filled in by the script, \n is a line break.
$uiText = @'
AppTitle          = Chuyển đổi phiên bản Windows 11
NeedWin11         = Script này chỉ hỗ trợ Windows 11 (máy đang chạy: {0}, build {1}).
Relaunching       = Cần quyền quản trị. Đang mở lại script với quyền quản trị...
ElevationCancel   = Bạn đã hủy yêu cầu quyền quản trị.

SkuNotWritten     = [!] Chưa có logic chép file SKU. Hãy viết logic trong hàm Invoke-SkuConversion.

OfficialBox       = Cảnh báo: theo chính thống, Microsoft không hỗ trợ {0}, việc chuyển bắt buộc cài lại Windows. Hiện tại nếu đổi sẽ áp dụng việc chép file SKU, việc chép SKU sẽ gây 1 số vấn đề trong tương lai khi update window.
OffEval           = chuyển bản Evaluation (dùng thử) sang bản chính thức
OffSE             = chuyển Windows 11 SE sang phiên bản khác
OffChina          = chuyển Home China sang phiên bản khác
OffLtsc           = chuyển LTSC sang phiên bản thường (General Availability) tại chỗ
OffN              = đổi tại chỗ giữa bản N và bản thường
OffOther          = chuyển sang phiên bản này từ bản cài hiện tại (DISM không liệt kê)

ResultNotRun      = [!] Chưa chuyển đổi.
ResultOk1         = [+] Windows đã chấp nhận đổi phiên bản.
ResultOk2         = [+] Hãy khởi động lại Windows để hoàn tất. Hoặc máy có thể tự khởi động lại.
ResultOk3         = [i] Nếu Windows chưa được kích hoạt, hãy nhập key bản quyền tại Cài đặt > Hệ thống > Kích hoạt.
ActTrying         = [+] Đang kích hoạt Windows bằng key bạn nhập...
ActOk             = [+] Kích hoạt thành công.
ActLater          = [i] Chưa kích hoạt được lúc này ({0}), Windows sẽ tự kích hoạt sau khi khởi động lại (có thể cần Internet). Bấm Kiểm tra kích hoạt để xem.
LogUserKey        = [i] Dùng key bạn nhập: chuyển phiên bản rồi kích hoạt luôn.
LogGenericKey     = [i] Dùng key chung đã chọn để chuyển phiên bản (không kích hoạt).
KeyNeeded         = Hãy nhập Active key vào ô bên dưới.
InvalidKeyHint    = [i] Windows từ chối key này. Hãy chọn loại key khác trong ô Active key, hoặc nhập key bản quyền.
KeyOptGeneric     = (Generic Key No Activation)
KeyOptGvlk        = (GVLK Key No Activation)
KeyOptUser        = Nhập key của bạn...
KeyOptUserNote    = (Kích hoạt luôn sau khi chuyển)
ResultFail1       = [!] Mã lỗi 0x{0:X8}. Phiên bản vẫn chưa được thay đổi.
ResultFail2       = [!] Hãy kiểm tra key có đúng với phiên bản đã chọn không.
DismFailed        = DISM bị lỗi, mã thoát 0x{0:X8}.
KeyFormat         = Active key phải có dạng XXXXX-XXXXX-XXXXX-XXXXX-XXXXX.

StQuery           = > Đang đọc SoftwareLicensingProduct (WMI)...
StNoKey           = Chưa cài product key nào cho Windows.
St0               = Chưa có bản quyền
St1               = Đã kích hoạt
St2               = Đang trong thời gian dùng thử ban đầu
St3               = Đang trong thời gian dùng thử bổ sung
St4               = Đang trong thời gian gia hạn (bản không chính hãng)
St5               = Chế độ thông báo (chưa kích hoạt)
St6               = Đang trong thời gian dùng thử mở rộng
StUnknown         = Không xác định ({0})
StPermanent       = Vĩnh viễn
StDaysLeft        = {0} (còn {1} ngày)
LblName           = Tên         : {0}
LblDesc           = Mô tả       : {0}
LblChannel        = Kênh        : {0}
LblKey            = Key         : *****-*****-*****-*****-{0}
LblStatus         = Trạng thái  : {0}
LblExpiry         = Hết hạn     : {0}
StatusError       = Không đọc được trạng thái kích hoạt: {0}

ConTitle          = Chuyển đổi phiên bản Windows 11 (chế độ dòng lệnh)
ConCurrent        = Hiện tại: {0} [{1}] - Build {2}
ConKeyOptional    = Active key (Enter = bỏ qua)
ConCheck          = [A] Kiểm tra trạng thái kích hoạt
ConPrompt         = Nhập số, A, hoặc nhấn Enter để thoát
ConInvalid        = Lựa chọn không hợp lệ.
ConKeyPick        = Chọn key (Enter = 1)
ConKey            = Active key
ConConfirm        = Gõ YES để chuyển sang {0}
ConActivate       = Hãy kích hoạt trong Cài đặt (Settings) > Hệ thống > Kích hoạt.
NoneAvailable     = DISM không liệt kê phiên bản nâng cấp nào cho bản cài này.

BtnStatus         = Kiểm tra kích hoạt
BtnConvert        = Chuyển đổi
LblKeyUi          = Active key (nếu có):
KeyTooltip        = Generic Key / GVLK Key: chỉ chọn phiên bản, không kích hoạt Windows. Nhập key của bạn: chuyển phiên bản rồi kích hoạt luôn.
Current           = Hiện tại: {0}  [{1}]  -  Build {2}
GroupUpgrade      = Nâng cấp phiên bản (changepk)
GroupSku          = Chuyển đổi phiên bản (Phải chép file SKU)
TipOk             = Nâng cấp phiên bản bằng changepk
HintOk            = Chuyển sang {0} bằng changepk
ConfirmTitle      = Xác nhận
ConfirmBody       = Chuyển sang {0} bằng key chung?\n\nWindows sẽ cần khởi động lại và có thể tự khởi động lại.\nKey chung KHÔNG kích hoạt Windows, sau đó hãy nhập key bản quyền của bạn để tự kích hoạt.
ConfirmBodyKey    = Chuyển sang {0} bằng key bạn nhập, rồi kích hoạt Windows?\n\nWindows sẽ cần khởi động lại và có thể tự khởi động lại.
ConfirmSkuBody    = Chuyển sang {0}?\n\nViệc chuyển này có thể không thực hiện được và có thể khiến Windows không kích hoạt hoặc không cập nhật được.\nCông cụ này KHÔNG kích hoạt Windows.

LogBanner         = Công cụ chuyển đổi phiên bản Windows 11 - chỉ kích hoạt khi bạn nhập key bản quyền.
LogReading        = [+] Đang đọc danh sách phiên bản mà microsfot có thể hỗ trợ chuyển đổi chính thống...
LogListFailed     = [!] Không đọc được danh sách phiên bản từ DISM.
LogLegend         = [i] Màu xanh lá: nâng cấp phiên bản bằng changepk (Được support). Màu đỏ: Script tự Download và chép file SKU.
LogPick           = [i] Chọn phiên bản, điền ActiveKey nếu có, rồi bấm Chuyển đổi.
LogOpening        = [+] Đang mở Cài đặt > Hệ thống > Kích hoạt...
LogOpenFailed     = [!] Không mở được trang Kích hoạt. Hãy tự mở Cài đặt > Hệ thống > Kích hoạt.
LogChecking       = [+] Đang kiểm tra trạng thái kích hoạt...
LogConverting     = [+] Đang chuyển sang {0}...
StatusLoading     = Đang tải...
StatusFailed      = Thất bại
StatusChecking    = Đang kiểm tra kích hoạt...
StatusConverting  = Đang chuyển đổi...
StatusReady       = Sẵn sàng
SecTarget         = Chọn phiên bản window cần chuyển đổi
SecLog            = Logs
'@

if ($PSCommandPath) {
    # Run from a file: re-read the block as UTF-8, whatever code page the parser used
    try {
        $self  = [IO.File]::ReadAllText($PSCommandPath, [Text.Encoding]::UTF8)
        $block = [regex]::Match($self, "(?s)\`$uiText = @'\r?\n(.*?)\r?\n'@")
        if ($block.Success) { $uiText = $block.Groups[1].Value }
    } catch {}
} elseif ($uiText -match '\u00E1[\u00BA\u00BB]|\u00C6\u00B0|\u00C4[\u0090\u0091]') {
    # Downloaded without a charset, so irm decoded it as ISO-8859-1: restore the UTF-8 text
    $uiText = [Text.Encoding]::UTF8.GetString([Text.Encoding]::GetEncoding(28591).GetBytes($uiText))
}

$Lang = @{}
foreach ($line in $uiText -split "`r?`n") {
    if ($line -match '^\s*(\w+)\s*=\s*(.*?)\s*$') { $Lang[$Matches[1]] = $Matches[2].Replace('\n', "`n") }
}

# 1. Require Windows 11 (client, build 22000 or later)
$cv             = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
$build          = [int]$cv.CurrentBuildNumber
$currentEdition = $cv.EditionID
if ($cv.InstallationType -ne 'Client' -or $build -lt 22000) {
    Write-Warning ($Lang.NeedWin11 -f $cv.InstallationType, $build)
    return
}

# 2. Require Administrator
$principal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning $Lang.Relaunching

    if ($PSCommandPath) {
        $argList = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    } else {
        $scriptUrl = 'https://raw.githubusercontent.com/M1ke193/covert-ms/main/convertMS.ps1'
        $command   = "irm '$scriptUrl' | iex"
        $encoded   = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($command))
        $argList   = "-NoProfile -ExecutionPolicy Bypass -EncodedCommand $encoded"
    }

    try {
        Start-Process powershell.exe -Verb RunAs -ArgumentList $argList -ErrorAction Stop
    } catch {
        Write-Warning $Lang.ElevationCancel
    }
    return
}

# 3. System information
# Windows 11 is 64-bit only, but a 32-bit PowerShell must use Sysnative to reach the real 64-bit DISM / changepk
$systemDir = if (-not [Environment]::Is64BitProcess) { "$env:windir\Sysnative" } else { "$env:windir\System32" }

# The registry ProductName still says "Windows 10" on Windows 11, so prefer the WMI caption
$osName = try { (Get-CimInstance Win32_OperatingSystem -ErrorAction Stop).Caption -replace '^Microsoft\s+' } catch { "Windows 11 $currentEdition" }

# 4. Microsoft generic keys (GVLK)
# Published by Microsoft: https://learn.microsoft.com/windows-server/get-started/kms-client-activation-keys
# They only select the edition. They do not activate Windows. You can type your own key instead.
$genericKeys = @{
    Professional             = 'W269N-WFGWX-YVC9B-4J6C9-T83GX'
    ProfessionalN            = 'MH37W-N47XK-V7XM9-C7227-GCQG9'
    ProfessionalWorkstation  = 'NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J'
    ProfessionalWorkstationN = '9FNHH-K3HBT-3W4TD-6383H-6XYWF'
    ProfessionalEducation    = '6TP4R-GNPTD-KYYHQ-7B7DP-J447Y'
    ProfessionalEducationN   = 'YVWGF-BXNMC-HTQYQ-CPQ99-66QFC'
    Education                = 'NW6C2-QMPVW-D7KKK-3GKT6-VCFB2'
    EducationN               = '2WH4N-8QGBV-H22JP-CT43Q-MDWWJ'
    Enterprise               = 'NPPR9-FWDCX-D2C8J-H872K-2YT43'
    EnterpriseN              = 'DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4'
    EnterpriseG              = 'YYVX9-NTFWV-6MDM3-9PT4T-4M68B'
    EnterpriseGN             = '44RPN-FTY23-9VTTB-MP9BX-T84FV'
    EnterpriseS              = 'M7XTQ-FN8P6-TTKYV-9D4CC-J462D'   # Enterprise LTSC 2024
    EnterpriseSN             = '92NFX-8DJQP-P6BBQ-THF9C-7CG2H'   # Enterprise N LTSC 2024
    IoTEnterpriseS           = 'KBN8V-HFGQ4-MGXVD-347P6-PDQGT'   # IoT Enterprise LTSC 2024
}

# Generic installation keys (default setup keys). They also only select the edition and never activate.
# Windows rejects the Volume (GVLK) keys above from some editions (for example Home) with 0xC004F050,
# so this "Generic Key" is the default choice in the Active key selector. Community list: not on the
# Microsoft KMS page.
$setupKeys = @{
    Core                     = 'YTMG3-N6DKC-DKB77-7M9GH-8HVX7'
    CoreN                    = '4CPRK-NM3K3-X6XXQ-RXX86-WXCHW'
    CoreSingleLanguage       = 'BT79Q-G7N6G-PGBYW-4YWX6-6F4BT'
    CoreCountrySpecific      = 'N2434-X9D7W-8PF6X-8DV9T-8TYMD'
    Professional             = 'VK7JG-NPHTM-C97JM-9MPGT-3V66T'
    ProfessionalN            = '2B87N-8KFHP-DKV6R-Y2C8J-PKCKT'
    ProfessionalWorkstation  = 'DXG7C-N36C4-C4HTG-X4T3X-2YV77'
    ProfessionalWorkstationN = 'WYPNQ-8C467-V2W6J-TX4WX-WT2RQ'
    ProfessionalEducation    = '8PTT6-RNW4C-6V7J2-C2D3X-MHBPB'
    ProfessionalEducationN   = 'GJTYN-HDMQY-FRR76-HVGC7-QPF8P'
    Education                = 'YNMGQ-8RYV3-4PGQ3-C8XTP-7CFBY'
    EducationN               = '84NGF-MHBT6-FXBX8-QWJK7-DRR8H'
    Enterprise               = 'XGVPP-NMH47-7TTHJ-W3FW7-8HV2C'
    EnterpriseN              = 'WGGHN-J84D6-QYCPR-T7PJ7-X766F'
    EnterpriseGN             = 'FW7NV-4T673-HF4VX-9X4MM-B4H4T'
}

# Choices for the Active key selector of one edition, as Tuple(Kind, Text, Note):
#   Generic = setup key, Gvlk = GVLK (both only select the edition), User = the user types a key.
# A kind is left out when the edition has no such key. The first choice is the default.
function Get-KeyOptions ($Edition) {
    if ($setupKeys[$Edition]) {
        [System.Tuple[string, string, string]]::new('Generic', $setupKeys[$Edition], $Lang.KeyOptGeneric)
    }
    if ($genericKeys[$Edition]) {
        [System.Tuple[string, string, string]]::new('Gvlk', $genericKeys[$Edition], $Lang.KeyOptGvlk)
    }
    [System.Tuple[string, string, string]]::new('User', $Lang.KeyOptUser, $Lang.KeyOptUserNote)
}

$keyPattern = '^([A-Z0-9]{5}-){4}[A-Z0-9]{5}$'

# Official edition names stay in English
$editionNames = @{
    Core = 'Home'; CoreN = 'Home N'; CoreSingleLanguage = 'Home Single Language'; CoreCountrySpecific = 'Home China'
    Professional = 'Pro'; ProfessionalN = 'Pro N'
    ProfessionalSingleLanguage = 'Pro Single Language'; ProfessionalCountrySpecific = 'Pro China'
    ProfessionalWorkstation = 'Pro for Workstations'; ProfessionalWorkstationN = 'Pro for Workstations N'
    ProfessionalEducation = 'Pro Education'; ProfessionalEducationN = 'Pro Education N'
    Education = 'Education'; EducationN = 'Education N'
    Enterprise = 'Enterprise'; EnterpriseN = 'Enterprise N'
    EnterpriseG = 'Enterprise G'; EnterpriseGN = 'Enterprise G N'
    EnterpriseS = 'Enterprise LTSC'; EnterpriseSN = 'Enterprise N LTSC'
    EnterpriseEval = 'Enterprise (Evaluation)'; EnterpriseNEval = 'Enterprise N (Evaluation)'
    EnterpriseSEval = 'Enterprise LTSC (Evaluation)'; EnterpriseSNEval = 'Enterprise N LTSC (Evaluation)'
    IoTEnterprise = 'IoT Enterprise'; IoTEnterpriseS = 'IoT Enterprise LTSC'; IoTEnterpriseK = 'IoT Enterprise K'
    ServerRdsh = 'Enterprise multi-session'
    CloudEdition = 'SE'; CloudEditionN = 'SE N'
}

# Common editions that are always shown; the ones DISM does not list use the SKU file method
$catalog = 'Core', 'CoreN', 'CoreSingleLanguage',
    'Professional', 'ProfessionalN', 'ProfessionalWorkstation', 'ProfessionalWorkstationN',
    'ProfessionalEducation', 'ProfessionalEducationN', 'Education', 'EducationN',
    'Enterprise', 'EnterpriseN'

# 5. Helpers
function Get-EditionName ($Id) {
    $name = $editionNames[$Id]
    if (-not $name) { $name = $Id }
    "Windows 11 $name"
}

# Official reason shown in the red box for an edition DISM did NOT list (changepk cannot do it).
# A few known cases get their own text; everything else is simply "not listed by DISM".
function Get-SkuReason ($From, $To) {
    $fromIsN = $From -cmatch 'N(Eval)?$'
    $toIsN   = $To -cmatch 'N$'

    $case = if ($From -like '*Eval') {
        'Eval'
    } elseif ($From -like 'CloudEdition*') {
        'SE'
    } elseif ($From -eq 'CoreCountrySpecific') {
        'China'
    } elseif ($From -like '*EnterpriseS*' -and $To -notlike '*EnterpriseS*') {
        'Ltsc'
    } elseif ($fromIsN -ne $toIsN) {
        'N'
    } else {
        'Other'
    }
    $Lang.OfficialBox -f $Lang["Off$case"]
}

# Editions from the catalog that DISM did not list (red), in catalog order
function Get-SkuEditions ($Targets) {
    foreach ($e in $catalog) {
        if ($e -eq $currentEdition -or $Targets -contains $e) { continue }
        [pscustomobject]@{
            Edition  = $e
            Name     = Get-EditionName $e
            Official = Get-SkuReason $currentEdition $e
        }
    }
}

function Get-ResultMessage ($ExitCode) {
    if ($null -eq $ExitCode) {
        $Lang.ResultNotRun
    } elseif ($ExitCode -in 0, 3010) {
        $Lang.ResultOk1
        $Lang.ResultOk2
        $Lang.ResultOk3
    } else {
        $Lang.ResultFail1 -f [int]$ExitCode
        $Lang.ResultFail2
    }
}

# The functions below also run in the background runspace, so they must not touch the UI
# and receive the UI text table ($Text) as a parameter.
# Run a built-in tool hidden and stream its output line by line to $Log
function Invoke-Tool ($Exe, $ArgList, $Log, $Secret) {
    $shown = if ($Secret) { $ArgList.Replace($Secret, '*****-*****-*****-*****-' + $Secret.Substring(24)) } else { $ArgList }
    & $Log ("> {0} {1}" -f [IO.Path]::GetFileName($Exe), $shown)

    $psi = [System.Diagnostics.ProcessStartInfo]::new($Exe, $ArgList)
    $psi.UseShellExecute        = $false
    $psi.CreateNoWindow         = $true
    $psi.RedirectStandardInput  = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true

    $proc = [System.Diagnostics.Process]::Start($psi)
    $proc.StandardInput.Close()
    $stderr = $proc.StandardError.ReadToEndAsync()
    $lines  = [System.Collections.Generic.List[string]]::new()
    while ($null -ne ($line = $proc.StandardOutput.ReadLine())) {
        $text = $line.Trim()
        if ($text) { $lines.Add($text); & $Log "  $text" }
    }
    $proc.WaitForExit()
    foreach ($line in ($stderr.Result -split "`r?`n")) {
        $text = $line.Trim()
        if ($text) { & $Log "  $text" }
    }
    [pscustomobject]@{ ExitCode = $proc.ExitCode; Lines = $lines }
}

# Ask DISM which editions this installation can be converted to (English output, so it can be parsed)
function Get-TargetEditions ($SystemDir, $Log, $Text) {
    $result = Invoke-Tool ([IO.Path]::Combine($SystemDir, 'dism.exe')) '/English /Online /Get-TargetEditions' $Log
    if ($result.ExitCode -ne 0) { throw ($Text.DismFailed -f [int]$result.ExitCode) }
    foreach ($line in $result.Lines) {
        if ($line -match '^Target Edition\s*:\s*(\S+)') { $Matches[1] }
    }
}

# Change the edition with changepk.exe (green editions: listed by DISM)
function Invoke-EditionChange ($Key, $SystemDir, $Log) {
    (Invoke-Tool ([IO.Path]::Combine($SystemDir, 'changepk.exe')) "/ProductKey $Key" $Log $Key).ExitCode
}

# Activate the installed Windows key online with Microsoft (same as slmgr /ato).
# Only called when the user typed their own key. Returns $true when activated.
function Invoke-Activation ($Log, $Text) {
    & $Log $Text.ActTrying
    try {
        $filter   = "ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' AND PartialProductKey IS NOT NULL"
        $products = @(Get-CimInstance -ClassName SoftwareLicensingProduct -Filter $filter -ErrorAction Stop)
        if (-not $products.Count) { throw $Text.StNoKey }
        foreach ($p in $products) { $null = Invoke-CimMethod -InputObject $p -MethodName Activate -ErrorAction Stop }
        & $Log $Text.ActOk
        $true
    } catch {
        & $Log ($Text.ActLater -f $_.Exception.Message.Trim())
        $false
    }
}

# Change the edition with the SKU file method (red editions: not listed by DISM).
# ---- WRITE YOUR OWN LOGIC HERE ----
# Parameters:
#   $Edition   target EditionID, for example 'Core' or 'Professional'
#   $Key       product key typed by the user, or '' when the box was left empty
#   $SystemDir System32 (or Sysnative from a 32-bit PowerShell)
#   $Log       call  & $Log "text"  to write a line to the log
#   $Text      the UI text table ($Lang)
# Return 0 or 3010 when the change succeeded (3010 = restart required), another exit code
# when it failed, or $null when nothing was done.
# This function also runs in the background runspace, so it must not touch the UI.
function Invoke-SkuConversion ($Edition, $Key, $SystemDir, $Log, $Text) {
    & $Log $Text.SkuNotWritten
    $null
}

# Read-only activation check through the built-in licensing WMI class
# (same data as slmgr /dli + /xpr, but it does not need VBScript, which newer Windows builds can remove)
function Get-ActivationStatus ($Log, $Text) {
    # 55c92734-... is the Windows application ID; other IDs belong to Office and similar products
    $filter = "ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' AND PartialProductKey IS NOT NULL"
    & $Log $Text.StQuery
    $products = @(Get-CimInstance -ClassName SoftwareLicensingProduct -Filter $filter -ErrorAction Stop)
    if (-not $products.Count) {
        & $Log "  $($Text.StNoKey)"
        return
    }
    foreach ($p in $products) {
        $code   = [int]$p.LicenseStatus
        $grace  = [long]$p.GracePeriodRemaining
        $status = if ($code -ge 0 -and $code -le 6) { $Text["St$code"] } else { $Text.StUnknown -f $code }
        $expiry = if ($code -eq 1 -and $grace -eq 0) {
            $Text.StPermanent
        } elseif ($grace -gt 0) {
            $Text.StDaysLeft -f [DateTime]::Now.AddMinutes($grace).ToString('dd/MM/yyyy HH:mm'), [Math]::Floor($grace / 1440)
        } else {
            '-'
        }
        & $Log ("  " + ($Text.LblName -f $p.Name))
        & $Log ("  " + ($Text.LblDesc -f $p.Description))
        if ($p.ProductKeyChannel) { & $Log ("  " + ($Text.LblChannel -f $p.ProductKeyChannel)) }
        & $Log ("  " + ($Text.LblKey -f $p.PartialProductKey))
        & $Log ("  " + ($Text.LblStatus -f $status))
        & $Log ("  " + ($Text.LblExpiry -f $expiry))
    }
}

# 6. Console mode (when WPF cannot be used, for example in a PowerShell remoting session)
function Start-ConsoleMode {
    $log = { param($m) Write-Host $m -ForegroundColor DarkGray }
    Write-Host $Lang.ConTitle -ForegroundColor Cyan
    Write-Host ($Lang.ConCurrent -f $osName, $currentEdition, $build)
    try {
        $targets = @(Get-TargetEditions $systemDir $log $Lang)
    } catch {
        Write-Warning $_.Exception.Message
        return
    }

    $skuEditions = @(Get-SkuEditions $targets)
    # One numbered list: green editions (changepk) first, then red ones (SKU file method)
    $choices = @(foreach ($t in $targets) { [pscustomobject]@{ Edition = $t; Name = (Get-EditionName $t); Sku = $false } }) +
               @(foreach ($s in $skuEditions) { [pscustomobject]@{ Edition = $s.Edition; Name = $s.Name; Sku = $true; Official = $s.Official } })

    while ($true) {
        Write-Host ""
        if ($targets) {
            Write-Host "$($Lang.GroupUpgrade):" -ForegroundColor Green
        } else {
            Write-Warning $Lang.NoneAvailable
        }
        $skuHeader = $false
        for ($i = 0; $i -lt $choices.Count; $i++) {
            $c = $choices[$i]
            if ($c.Sku) {
                # All SKU-method editions share one heading
                if (-not $skuHeader) { Write-Host "$($Lang.GroupSku):" -ForegroundColor Red; $skuHeader = $true }
                Write-Host ("  [{0}] {1}" -f ($i + 1), $c.Name) -ForegroundColor Red
            } else {
                Write-Host ("  [{0}] {1}" -f ($i + 1), $c.Name) -ForegroundColor Green
            }
        }
        Write-Host "  $($Lang.ConCheck)"

        $answer = (Read-Host $Lang.ConPrompt).Trim()
        if (-not $answer) { return }
        if ($answer -eq 'A') {
            try { Get-ActivationStatus $log $Lang } catch { Write-Warning ($Lang.StatusError -f $_.Exception.Message) }
            continue
        }

        $n = $answer -as [int]
        if (-not $n -or $n -lt 1 -or $n -gt $choices.Count) {
            Write-Warning $Lang.ConInvalid
            continue
        }

        $choice  = $choices[$n - 1]
        $edition = $choice.Edition
        if ($choice.Sku) {
            Write-Host ""
            Write-Host "  $($choice.Official)" -ForegroundColor Red
        }

        # Same three key choices as the window: generic key, GVLK key, or the user's own key
        $options = @(Get-KeyOptions $edition)
        Write-Host ""
        for ($i = 0; $i -lt $options.Count; $i++) {
            Write-Host ("  [{0}] {1}  {2}" -f ($i + 1), $options[$i].Item2, $options[$i].Item3)
        }
        $pick = (Read-Host $Lang.ConKeyPick).Trim()
        $p    = if ($pick) { $pick -as [int] } else { 1 }
        if (-not $p -or $p -lt 1 -or $p -gt $options.Count) {
            Write-Warning $Lang.ConInvalid
            continue
        }
        $option  = $options[$p - 1]
        $userKey = $option.Item1 -eq 'User'
        $key     = if ($userKey) {
            (Read-Host $(if ($choice.Sku) { $Lang.ConKeyOptional } else { $Lang.ConKey })).Trim().ToUpper()
        } else {
            $option.Item2
        }
        # The SKU method may run without a key; changepk always needs one
        if (($key -or -not $choice.Sku) -and $key -notmatch $keyPattern) {
            Write-Warning $Lang.KeyFormat
            continue
        }

        if ((Read-Host ($Lang.ConConfirm -f $choice.Name)) -ne 'YES') { continue }
        $code = if ($choice.Sku) {
            Invoke-SkuConversion $edition $key $systemDir $log $Lang
        } else {
            Invoke-EditionChange $key $systemDir $log
        }
        Get-ResultMessage $code | ForEach-Object { Write-Host $_ }
        if ($code -eq -1073418160) { Write-Host $Lang.InvalidKeyHint }   # 0xC004F050
        if ($code -in 0, 3010) {
            if ($userKey -and -not $choice.Sku) { $null = Invoke-Activation $log $Lang }
            # Always show the Activation page after a successful change
            Write-Host "    $($Lang.ConActivate)"
            try { Start-Process 'ms-settings:activation' -ErrorAction Stop } catch {}
        }
        return
    }
}

$guiAvailable = $false
try {
    Add-Type -AssemblyName PresentationFramework, PresentationCore -ErrorAction Stop
    $guiAvailable = [Threading.Thread]::CurrentThread.ApartmentState -eq 'STA'
} catch {}

if (-not $guiAvailable) {
    Start-ConsoleMode
    return
}

# 7. GUI mode: dark editor-style window (header, target row, reason box, colored log, status bar)
# Visible text is set from $Lang after loading, so the XAML stays ASCII
$xamlText = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Width="960" Height="620" MinWidth="800" MinHeight="460"
        WindowStartupLocation="CenterScreen" Background="#0D1117" Foreground="#C9D1D9"
        FontFamily="Cascadia Mono, Cascadia Code, Consolas" FontSize="12.5"
        UseLayoutRounding="True" TextOptions.TextFormattingMode="Display">
    <Window.Resources>
        <!-- Palette: bg #0D1117, panel #161B22, line #30363D, text #C9D1D9, muted #8B949E, accent #58A6FF -->

        <Style TargetType="ToolTip">
            <Setter Property="Background" Value="#161B22"/>
            <Setter Property="Foreground" Value="#C9D1D9"/>
            <Setter Property="BorderBrush" Value="#30363D"/>
            <Setter Property="Padding" Value="8,5"/>
            <Setter Property="MaxWidth" Value="460"/>
            <Setter Property="ContentTemplate">
                <Setter.Value>
                    <DataTemplate><TextBlock Text="{Binding}" TextWrapping="Wrap"/></DataTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="Button">
            <Setter Property="Foreground" Value="#C9D1D9"/>
            <Setter Property="Background" Value="#21262D"/>
            <Setter Property="BorderBrush" Value="#30363D"/>
            <Setter Property="Padding" Value="16,0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="1" CornerRadius="6" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="bd" Property="BorderBrush" Value="#8B949E"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="bd" Property="Opacity" Value="0.8"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="bd" Property="Opacity" Value="0.4"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="PrimaryButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
            <Setter Property="Background" Value="#238636"/>
            <Setter Property="BorderBrush" Value="#2EA043"/>
            <Setter Property="Foreground" Value="#FFFFFF"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
        </Style>

        <Style TargetType="TextBox">
            <Setter Property="Background" Value="#0D1117"/>
            <Setter Property="Foreground" Value="#C9D1D9"/>
            <Setter Property="BorderBrush" Value="#30363D"/>
            <Setter Property="CaretBrush" Value="#58A6FF"/>
            <Setter Property="SelectionBrush" Value="#1F6FEB"/>
            <Setter Property="Padding" Value="8,0"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TextBox">
                        <Border x:Name="bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="1" CornerRadius="6">
                            <ScrollViewer x:Name="PART_ContentHost" Margin="{TemplateBinding Padding}"
                                          VerticalAlignment="{TemplateBinding VerticalContentAlignment}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="bd" Property="BorderBrush" Value="#8B949E"/>
                            </Trigger>
                            <Trigger Property="IsKeyboardFocused" Value="True">
                                <Setter TargetName="bd" Property="BorderBrush" Value="#58A6FF"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="bd" Property="Opacity" Value="0.4"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="ComboToggle" TargetType="ToggleButton">
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ToggleButton">
                        <Border x:Name="bd" Background="#0D1117" BorderBrush="#30363D" BorderThickness="1" CornerRadius="6">
                            <Path HorizontalAlignment="Right" VerticalAlignment="Center" Margin="0,0,12,0"
                                  Data="M0,0 L4,4 L8,0" Stroke="#8B949E" StrokeThickness="1.5"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="bd" Property="BorderBrush" Value="#8B949E"/>
                            </Trigger>
                            <Trigger Property="IsChecked" Value="True">
                                <Setter TargetName="bd" Property="BorderBrush" Value="#58A6FF"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="ComboBox">
            <Setter Property="Foreground" Value="#C9D1D9"/>
            <Setter Property="MaxDropDownHeight" Value="360"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBox">
                        <Grid>
                            <ToggleButton Style="{StaticResource ComboToggle}" Focusable="False" ClickMode="Press"
                                          IsChecked="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}"/>
                            <ContentPresenter Margin="12,0,30,0" VerticalAlignment="Center" IsHitTestVisible="False"
                                              Content="{TemplateBinding SelectionBoxItem}"
                                              ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"/>
                            <Popup x:Name="PART_Popup" Placement="Bottom" AllowsTransparency="True" Focusable="False"
                                   PopupAnimation="Fade"
                                   IsOpen="{Binding IsDropDownOpen, RelativeSource={RelativeSource TemplatedParent}}">
                                <Border Background="#161B22" BorderBrush="#30363D" BorderThickness="1" CornerRadius="6"
                                        Margin="0,4,0,0" Padding="0,4"
                                        MinWidth="{Binding ActualWidth, RelativeSource={RelativeSource TemplatedParent}}"
                                        MaxHeight="{TemplateBinding MaxDropDownHeight}">
                                    <ScrollViewer>
                                        <ItemsPresenter KeyboardNavigation.DirectionalNavigation="Contained"/>
                                    </ScrollViewer>
                                </Border>
                            </Popup>
                        </Grid>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.4"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="ComboBoxItem">
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBoxItem">
                        <Border x:Name="bd" Padding="12,6" Background="Transparent">
                            <ContentPresenter/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="bd" Property="Background" Value="#21262D"/>
                            </Trigger>
                            <Trigger Property="IsHighlighted" Value="True">
                                <Setter TargetName="bd" Property="Background" Value="#1C2D41"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Slim scrollbar -->
        <Style TargetType="ScrollBar">
            <Setter Property="Width" Value="10"/>
            <Setter Property="MinWidth" Value="10"/>
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ScrollBar">
                        <Track x:Name="PART_Track" IsDirectionReversed="True">
                            <Track.Thumb>
                                <Thumb>
                                    <Thumb.Template>
                                        <ControlTemplate TargetType="Thumb">
                                            <Border Background="#30363D" CornerRadius="4" Margin="2"/>
                                        </ControlTemplate>
                                    </Thumb.Template>
                                </Thumb>
                            </Track.Thumb>
                        </Track>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="Orientation" Value="Horizontal">
                    <Setter Property="Width" Value="NaN"/>
                    <Setter Property="MinWidth" Value="0"/>
                    <Setter Property="Height" Value="10"/>
                </Trigger>
            </Style.Triggers>
        </Style>
    </Window.Resources>

    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <Border Grid.Row="0" Background="#161B22" BorderBrush="#30363D" BorderThickness="0,0,0,1" Padding="18,12">
            <DockPanel>
                <Button x:Name="btnStatus" DockPanel.Dock="Right" Height="32" VerticalAlignment="Center" IsEnabled="False"/>                <StackPanel VerticalAlignment="Center">
                    <TextBlock FontSize="15" FontWeight="SemiBold">
                        <Run Text="PS&gt; " Foreground="#3FB950"/><Run Text="edition-converter" Foreground="#E6EDF3"/><Run Text=" -Target Win11" Foreground="#8B949E"/>
                    </TextBlock>
                    <TextBlock x:Name="txtCurrent" Margin="0,4,0,0" Foreground="#8B949E" TextTrimming="CharacterEllipsis"/>
                </StackPanel>
            </DockPanel>
        </Border>

        <!-- Body -->
        <DockPanel Grid.Row="1" Margin="18,14,18,14">
            <TextBlock x:Name="lblTarget" DockPanel.Dock="Top" Foreground="#8B949E" Margin="0,0,0,6"/>
            <Grid DockPanel.Dock="Top">
                <Grid.RowDefinitions>
                    <RowDefinition Height="46"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="300"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <ComboBox x:Name="cmbTarget" Grid.Column="0" IsEnabled="False" FontWeight="SemiBold"/>
                <TextBlock x:Name="lblKey" Grid.Column="1" Margin="16,0,8,0" VerticalAlignment="Center" Foreground="#8B949E"/>
                <!-- Active key selector: each choice shows the key and, below it, a note in brackets -->
                <ComboBox x:Name="cmbKey" Grid.Column="2" IsEnabled="False">
                    <ComboBox.ItemTemplate>
                        <DataTemplate>
                            <StackPanel Margin="0,1">
                                <TextBlock Text="{Binding Item2}" FontWeight="SemiBold"/>
                                <TextBlock Text="{Binding Item3}" Foreground="#8B949E" FontSize="11"/>
                            </StackPanel>
                        </DataTemplate>
                    </ComboBox.ItemTemplate>
                </ComboBox>
                <Button x:Name="btnConvert" Grid.Column="3" Style="{StaticResource PrimaryButton}" MinWidth="120"
                        Margin="10,0,0,0" IsEnabled="False"/>
                <!-- Shown only for "type your own key" -->
                <TextBox x:Name="txtKey" Grid.Row="1" Grid.Column="2" Height="34" Margin="0,8,0,0"
                         CharacterCasing="Upper" MaxLength="29" Visibility="Collapsed"/>
            </Grid>
            <Border x:Name="hintBox" DockPanel.Dock="Top" Margin="0,10,0,0" Padding="12,8" BorderThickness="1"
                    CornerRadius="6" Visibility="Collapsed">
                <TextBlock x:Name="txtHint" TextWrapping="Wrap" LineHeight="19"/>
            </Border>
            <TextBlock x:Name="lblLog" DockPanel.Dock="Top" Foreground="#8B949E" Margin="0,14,0,6"/>
            <Border Background="#010409" BorderBrush="#30363D" BorderThickness="1" CornerRadius="6">
                <RichTextBox x:Name="txtLog" Background="Transparent" Foreground="#C9D1D9" BorderThickness="0"
                             IsReadOnly="True" Padding="10,8" FontSize="12" SelectionBrush="#1F6FEB"
                             VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
                    <FlowDocument PagePadding="0" LineHeight="17" FontFamily="Cascadia Mono, Cascadia Code, Consolas">
                        <Paragraph x:Name="logPara" Margin="0"/>
                    </FlowDocument>
                </RichTextBox>
            </Border>
        </DockPanel>

        <!-- Status bar -->
        <Border Grid.Row="2" Background="#161B22" BorderBrush="#30363D" BorderThickness="0,1,0,0" Padding="18,6">
            <DockPanel>
                <TextBlock x:Name="txtEnv" DockPanel.Dock="Right" Foreground="#8B949E" FontSize="11.5"/>
                <StackPanel Orientation="Horizontal">
                    <Ellipse x:Name="dotStatus" Width="8" Height="8" Fill="#D29922" VerticalAlignment="Center" Margin="0,0,8,0"/>
                    <TextBlock x:Name="txtStatus" FontSize="11.5" VerticalAlignment="Center"/>
                </StackPanel>
            </DockPanel>
        </Border>
    </Grid>
</Window>
'@

[xml]$xaml = $xamlText
$window = [Windows.Markup.XamlReader]::Load([System.Xml.XmlNodeReader]::new($xaml))

$ui = @{}
foreach ($name in 'txtCurrent', 'btnStatus', 'lblTarget', 'cmbTarget', 'lblKey', 'cmbKey', 'txtKey', 'btnConvert',
                  'hintBox', 'txtHint', 'lblLog', 'txtLog', 'logPara', 'dotStatus', 'txtStatus', 'txtEnv') {
    $ui[$name] = $window.FindName($name)
}
$window.Title          = $Lang.AppTitle
$ui.txtCurrent.Text    = '# ' + ($Lang.Current -f $osName, $currentEdition, $build)
$ui.btnStatus.Content  = $Lang.BtnStatus
$ui.btnConvert.Content = $Lang.BtnConvert
$ui.lblTarget.Text     = "// $($Lang.SecTarget)"
$ui.lblKey.Text        = $Lang.LblKeyUi
$ui.cmbKey.ToolTip     = $Lang.KeyTooltip
$ui.txtKey.ToolTip     = $Lang.KeyTooltip
$ui.lblLog.Text        = "// $($Lang.SecLog)"
$ui.txtEnv.Text        = 'PowerShell {0}.{1} | build {2} | {3}' -f $PSVersionTable.PSVersion.Major, $PSVersionTable.PSVersion.Minor,
                         $build, $(if ([Environment]::Is64BitProcess) { 'x64' } else { 'x86' })

# Dark palette. Green = listed by DISM (changepk), red = SKU file method (Invoke-SkuConversion)
$brushConverter = [System.Windows.Media.BrushConverter]::new()
function New-Brush ($Hex) { $b = $brushConverter.ConvertFromString($Hex); $b.Freeze(); $b }
$brushOk       = New-Brush '#3FB950'
$brushOkBg     = New-Brush '#0F2A19'
$brushOkLine   = New-Brush '#238636'
$brushSku      = New-Brush '#FF7B72'
$brushSkuBg    = New-Brush '#2D1316'
$brushSkuLine  = New-Brush '#DA3633'
$brushInfo     = New-Brush '#58A6FF'
$brushCmd      = New-Brush '#D2A8FF'
$brushMuted    = New-Brush '#8B949E'
$brushBusy     = New-Brush '#D29922'
$brushText     = New-Brush '#C9D1D9'

$script:isBusy         = $true
$script:quietRefresh   = $false   # true while the silent refresh after a change is running

# Colored log: [+] green, [!] red, [i] blue, "> command" purple, indented tool output muted
function Get-LogBrush ($Line) {
    if ($Line -match '^\s*\[\+\]') { return $brushOk }
    if ($Line -match '^\s*\[!\]')  { return $brushSku }
    if ($Line -match '^\s*\[i\]')  { return $brushInfo }
    if ($Line -match '^>')         { return $brushCmd }
    if ($Line -match '^\s{2}')     { return $brushMuted }
    $null
}

function Add-LogLines ([string[]]$Lines) {
    foreach ($line in $Lines) {
        $run   = [System.Windows.Documents.Run]::new($line)
        $brush = Get-LogBrush $line
        if ($brush) { $run.Foreground = $brush }
        $ui.logPara.Inlines.Add($run)
        $ui.logPara.Inlines.Add([System.Windows.Documents.LineBreak]::new())
    }
    $ui.txtLog.ScrollToEnd()
}

function Write-Log ($Message) { Add-LogLines @([string]$Message) }

# Windows 11: dark title bar to match the window (ignored where not supported)
try {
    if (-not ('ConvertMS.Dwm' -as [type])) {
        Add-Type -Namespace ConvertMS -Name Dwm -ErrorAction Stop -MemberDefinition @'
[DllImport("dwmapi.dll")] public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attribute, ref int value, int size);
'@
    }
    $window.Add_SourceInitialized({
        try {
            $hwnd = [System.Windows.Interop.WindowInteropHelper]::new($window).Handle
            $on = 1;              [void][ConvertMS.Dwm]::DwmSetWindowAttribute($hwnd, 20, [ref]$on, 4)       # immersive dark mode
            $caption = 0x00221B16; [void][ConvertMS.Dwm]::DwmSetWindowAttribute($hwnd, 35, [ref]$caption, 4)  # caption #161B22 (BGR)
        } catch {}
    })
} catch {}

# Each list item carries @{ Edition; Name; Sku; Official } in its Tag
function Get-SelectedInfo {
    $item = $ui.cmbTarget.SelectedItem
    if ($item -and $item.Tag) { $item.Tag }
}

function Set-Busy ($Busy, $Status) {
    $script:isBusy = [bool]$Busy
    $info = Get-SelectedInfo
    $window.Title = if ($Status) { "$($Lang.AppTitle) - $Status" } else { $Lang.AppTitle }
    $ui.txtStatus.Text  = if ($Status) { $Status } else { $Lang.StatusReady }
    $ui.dotStatus.Fill  = if ($Status -eq $Lang.StatusFailed) { $brushSku } elseif ($Busy) { $brushBusy } else { $brushOk }
    $ui.cmbTarget.IsEnabled  = -not $Busy -and $ui.cmbTarget.Items.Count -gt 0
    $ui.cmbKey.IsEnabled     = -not $Busy -and $null -ne $info
    $ui.txtKey.IsEnabled     = -not $Busy -and $null -ne $info
    $ui.btnConvert.IsEnabled = -not $Busy -and $null -ne $info
    $ui.btnStatus.IsEnabled  = -not $Busy
}

# Background work runs in one reusable runspace and only talks to the UI through $sync.
# A timer drains the output queue in batches, so the window never freezes.
$sync = [hashtable]::Synchronized(@{
    Queue  = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
    Done   = $false
    Result = $null
})

$iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault2()
foreach ($fn in 'Invoke-Tool', 'Get-TargetEditions', 'Invoke-EditionChange', 'Invoke-Activation', 'Invoke-SkuConversion', 'Get-ActivationStatus') {
    $iss.Commands.Add([System.Management.Automation.Runspaces.SessionStateFunctionEntry]::new($fn, (Get-Item "function:$fn").ScriptBlock.ToString()))
}
$runspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace($iss)
$runspace.Open()

$timer = [System.Windows.Threading.DispatcherTimer]::new()
$timer.Interval = [TimeSpan]::FromMilliseconds(150)
$timer.Add_Tick({
    $finished = $sync.Done
    $batch = [System.Collections.Generic.List[string]]::new()
    $line = $null
    while ($sync.Queue.TryDequeue([ref]$line)) { $batch.Add($line) }
    if ($batch.Count) { Add-LogLines $batch.ToArray() }
    if (-not $finished) { return }

    $timer.Stop()
    $script:job.Dispose()
    & $script:onDone $sync.Result
})

function Start-Background ([scriptblock]$Work, [object[]]$Arguments, [scriptblock]$OnDone) {
    $sync.Done   = $false
    $sync.Result = $null
    $script:onDone = $OnDone
    $script:job = [PowerShell]::Create()
    $script:job.Runspace = $runspace
    [void]$script:job.AddScript($Work).AddArgument($sync)
    foreach ($a in $Arguments) { [void]$script:job.AddArgument($a) }
    [void]$script:job.BeginInvoke()
    $timer.Start()
}

$listWork = {
    param($sync, $systemDir, $text, $quiet)
    # Quiet (refresh after a change): DISM output is not written to the log; errors still are
    $log = if ($quiet) { { param($m) } } else { { param($m) $sync.Queue.Enqueue($m) } }
    try { $sync.Result = [pscustomobject]@{ Targets = @(Get-TargetEditions $systemDir $log $text) } }
    catch { $sync.Queue.Enqueue("[!] $($_.Exception.Message)") }
    finally { $sync.Done = $true }
}

$convertWork = {
    param($sync, $key, $systemDir, $activate, $text)
    $log = { param($m) $sync.Queue.Enqueue($m) }
    try {
        $code      = Invoke-EditionChange $key $systemDir $log
        $activated = $false
        if ($activate -and $code -in 0, 3010) { $activated = Invoke-Activation $log $text }
        $sync.Result = [pscustomobject]@{ Code = $code; Activated = $activated }
    }
    catch { $sync.Queue.Enqueue("[!] $($_.Exception.Message)") }
    finally { $sync.Done = $true }
}

$skuWork = {
    param($sync, $edition, $key, $systemDir, $text)
    $log = { param($m) $sync.Queue.Enqueue($m) }
    try { $sync.Result = [pscustomobject]@{ Code = (Invoke-SkuConversion $edition $key $systemDir $log $text); Activated = $false } }
    catch { $sync.Queue.Enqueue("[!] $($_.Exception.Message)") }
    finally { $sync.Done = $true }
}

$statusWork = {
    param($sync, $text)
    $log = { param($m) $sync.Queue.Enqueue($m) }
    try { Get-ActivationStatus $log $text }
    catch { $sync.Queue.Enqueue("[!] " + ($text.StatusError -f $_.Exception.Message)) }
    finally { $sync.Done = $true }
}

$onListed = {
    param($result)
    # After a refresh the lists are rebuilt silently; only the first load writes them to the log
    $quiet = $script:quietRefresh
    $script:quietRefresh = $false
    if (-not $result) {
        Write-Log ""
        Write-Log $Lang.LogListFailed
        Set-Busy $false $Lang.StatusFailed
        return
    }
    if (-not $quiet) { Write-Log "" }

    # Editions Windows offers: green
    $targets = @($result.Targets | Where-Object { $_ })
    foreach ($target in $targets) {
        $name = Get-EditionName $target
        $item = [System.Windows.Controls.ComboBoxItem]::new()
        $item.Content    = $name
        $item.Foreground = $brushOk
        $item.ToolTip    = $Lang.TipOk
        $item.Tag        = [pscustomobject]@{ Edition = $target; Name = $name; Sku = $false; Official = $null }
        [void]$ui.cmbTarget.Items.Add($item)
    }

    # Editions DISM does not list: red, converted with the SKU file method
    $skuEditions = @(Get-SkuEditions $targets)
    foreach ($s in $skuEditions) {
        $item = [System.Windows.Controls.ComboBoxItem]::new()
        $item.Content    = $s.Name
        $item.Foreground = $brushSku
        $item.ToolTip    = $s.Official
        $item.Tag        = [pscustomobject]@{ Edition = $s.Edition; Name = $s.Name; Sku = $true; Official = $s.Official }
        [void]$ui.cmbTarget.Items.Add($item)
    }

    if (-not $quiet) {
        if ($targets) {
            Write-Log "[+] $($Lang.GroupUpgrade):"
            foreach ($target in $targets) { Write-Log "    - $(Get-EditionName $target)" }
        } else {
            Write-Log "[!] $($Lang.NoneAvailable)"
        }
        # All SKU-method editions in one group
        if ($skuEditions) {
            Write-Log ""
            Write-Log "[i] $($Lang.GroupSku):"
            foreach ($s in $skuEditions) { Write-Log "    - $($s.Name)" }
        }
        Write-Log ""
    }

    if ($ui.cmbTarget.Items.Count) {
        $ui.cmbTarget.SelectedIndex = 0
        if (-not $quiet) {
            Write-Log $Lang.LogLegend
            Write-Log $Lang.LogPick
        }
    }
    Set-Busy $false
}

$onConverted = {
    param($result)
    # $result = @{ Code; Activated }, or $null when the work threw before finishing
    $code = if ($result) { $result.Code } else { $null }
    Get-ResultMessage $code | ForEach-Object { Write-Log $_ }
    if ($code -eq -1073418160) { Write-Log $Lang.InvalidKeyHint }   # 0xC004F050: try another key choice
    if ($code -in 0, 3010) {
        # Always show the Activation page after a successful change, activated or not
        Write-Log $Lang.LogOpening
        try {
            Start-Process 'ms-settings:activation' -ErrorAction Stop
        } catch {
            Write-Log $Lang.LogOpenFailed
        }
        # Read everything again so the window shows the new state (edition, DISM list, key choices)
        Invoke-Refresh
    } else {
        # Keep the current choices on failure so another key can be tried
        Set-Busy $false $Lang.StatusFailed
    }
}

# Refresh: re-read the running edition, then run DISM again and rebuild the lists.
# Called only after a successful change. Silent: nothing new is written to the log.
function Invoke-Refresh {
    try {
        $cv = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction Stop
        $script:build          = [int]$cv.CurrentBuildNumber
        $script:currentEdition = $cv.EditionID
        $script:osName = try { (Get-CimInstance Win32_OperatingSystem -ErrorAction Stop).Caption -replace '^Microsoft\s+' } catch { "Windows 11 $($cv.EditionID)" }
    } catch {}
    $ui.txtCurrent.Text = '# ' + ($Lang.Current -f $osName, $currentEdition, $build)
    $ui.txtEnv.Text     = 'PowerShell {0}.{1} | build {2} | {3}' -f $PSVersionTable.PSVersion.Major, $PSVersionTable.PSVersion.Minor,
                          $build, $(if ([Environment]::Is64BitProcess) { 'x64' } else { 'x86' })

    # Clear the old lists and the reason box
    $ui.cmbTarget.Items.Clear()
    $ui.cmbTarget.Foreground = $brushText
    $ui.cmbKey.Items.Clear()
    $ui.txtKey.Visibility    = 'Collapsed'
    $ui.hintBox.Visibility   = 'Collapsed'

    $script:quietRefresh = $true
    Set-Busy $true $Lang.StatusLoading
    Start-Background $listWork @($systemDir, $Lang, $true) $onListed
}

$onStatusChecked = {
    param($result)
    Set-Busy $false
}

$ui.btnStatus.Add_Click({
    Set-Busy $true $Lang.StatusChecking
    Write-Log ""
    Write-Log $Lang.LogChecking
    Start-Background $statusWork @(, $Lang) $onStatusChecked
})

$ui.cmbTarget.Add_SelectionChanged({
    $item = $ui.cmbTarget.SelectedItem
    $info = Get-SelectedInfo
    if (-not $info) { return }

    # The closed box takes the color of the selected edition
    $ui.cmbTarget.Foreground = $item.Foreground

    # Reload the key choices for this edition; the first one (generic key when it exists) is preselected
    $ui.cmbKey.Items.Clear()
    foreach ($option in Get-KeyOptions $info.Edition) { [void]$ui.cmbKey.Items.Add($option) }
    $ui.cmbKey.SelectedIndex = 0

    # Box under the list: green for changepk, red with the official reason for the SKU method
    $ui.hintBox.Visibility = 'Visible'
    if ($info.Sku) {
        $ui.hintBox.BorderBrush = $brushSkuLine
        $ui.hintBox.Background  = $brushSkuBg
        $ui.txtHint.Foreground  = $brushSku
        $ui.txtHint.Text        = $info.Official
    } else {
        $ui.hintBox.BorderBrush = $brushOkLine
        $ui.hintBox.Background  = $brushOkBg
        $ui.txtHint.Foreground  = $brushOk
        $ui.txtHint.Text        = $Lang.HintOk -f $info.Name
    }
    Set-Busy $script:isBusy
})

# "Type your own key" shows the input box under the selector
$ui.cmbKey.Add_SelectionChanged({
    $option = $ui.cmbKey.SelectedItem
    $custom = $null -ne $option -and $option.Item1 -eq 'User'
    $ui.txtKey.Visibility = if ($custom) { 'Visible' } else { 'Collapsed' }
    if ($custom -and -not $script:isBusy) { [void]$ui.txtKey.Focus() }
})

$ui.btnConvert.Add_Click({
    $info   = Get-SelectedInfo
    $option = $ui.cmbKey.SelectedItem
    if (-not $info -or -not $option) { return }
    # Generic / GVLK choice: that key, no activation. Own key: convert with it, then activate
    $userKey = $option.Item1 -eq 'User'
    $key     = if ($userKey) { $ui.txtKey.Text.Trim().ToUpper() } else { $option.Item2 }
    # The SKU method may run without a key; changepk always needs one
    if ($userKey -and -not $key -and -not $info.Sku) {
        Write-Log "[!] $($Lang.KeyNeeded)"
        return
    }
    if ($key -and $key -notmatch $keyPattern) {
        Write-Log "[!] $($Lang.KeyFormat)"
        return
    }
    $question = if ($info.Sku) { $Lang.ConfirmSkuBody } elseif ($userKey) { $Lang.ConfirmBodyKey } else { $Lang.ConfirmBody }
    if ([System.Windows.MessageBox]::Show(($question -f $info.Name), $Lang.ConfirmTitle, 'YesNo', 'Warning') -ne 'Yes') { return }

    Set-Busy $true $Lang.StatusConverting
    Write-Log ""
    Write-Log ($Lang.LogConverting -f $info.Name)
    if ($info.Sku) {
        Start-Background $skuWork @($info.Edition, [string]$key, $systemDir, $Lang) $onConverted
    } else {
        Write-Log $(if ($userKey) { $Lang.LogUserKey } else { $Lang.LogGenericKey })
        Start-Background $convertWork @($key, $systemDir, $userKey, $Lang) $onConverted
    }
})

# Show the window first, then read the target editions in the background
$window.Add_ContentRendered({
    Write-Log $Lang.LogBanner
    Write-Log $Lang.LogReading
    Set-Busy $true $Lang.StatusLoading
    Start-Background $listWork @($systemDir, $Lang) $onListed
})

# 8. Show window
$null = $window.ShowDialog()
$runspace.Dispose()
