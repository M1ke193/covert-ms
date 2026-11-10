# Windows 11 Edition Converter

Script PowerShell có giao diện WPF để **chuyển đổi phiên bản Windows 11 ngay trên máy đang chạy** (ví dụ Home → Pro, Pro → Education / Enterprise), dùng các công cụ có sẵn của Windows.

> Script **chỉ kích hoạt Windows khi bạn chọn "Nhập key của bạn" và nhập key bản quyền**. Với Generic Key hoặc GVLK Key, script chỉ chuyển phiên bản; Windows sẽ chưa được kích hoạt.

## Yêu cầu

- Windows 11 (bản cài thường, build 22000 trở lên). Windows 10 và Windows Server không được hỗ trợ.
- Windows PowerShell 5.1 (`powershell.exe`, có sẵn trên Windows 11).
- Quyền quản trị (Administrator). Khi chạy từ file, script tự mở lại với quyền quản trị.

## Cách chạy

Mở PowerShell tại thư mục chứa script:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\convertMS.ps1
```

hoặc

```powershell
irm https://raw.githubusercontent.com/M1ke193/covert-ms/main/convertMS.ps1 | iex
```

## Script làm được gì

### 1. Đọc danh sách phiên bản có thể chuyển
- Khi mở, script chạy `DISM /Online /Get-TargetEditions` để lấy danh sách phiên bản mà bản cài hiện tại được phép chuyển sang.
- Dòng đầu cửa sổ hiển thị phiên bản hiện tại, EditionID và build.

### 2. Hiển thị phiên bản theo màu
- **Màu xanh**: phiên bản DISM liệt kê, chuyển được bằng `changepk.exe`.
- **Màu đỏ**: mọi phiên bản DISM không liệt kê cho bản cài hiện tại, tức là `changepk` không chuyển được (ví dụ đổi giữa bản N và bản thường, từ LTSC hoặc bản Evaluation).
- Khi chọn một phiên bản, khung bên dưới hiển thị thông tin:
  - Xanh: phiên bản này được chuyển bằng `changepk`.
  - Đỏ: lý do chính thức của Microsoft.

### 3. Chuyển phiên bản
Chọn phiên bản, chọn loại key trong ô **Active key**, rồi bấm **Chuyển đổi** và xác nhận.

Ô **Active key** là ô chọn có tối đa 3 mục. Mỗi mục hiện key ở dòng trên và ghi chú trong ngoặc ở dòng dưới:

| Mục | Ví dụ (Pro) | Ghi chú | Kết quả |
|---|---|---|---|
| **Generic Key** (chọn sẵn) | `VK7JG-NPHTM-C97JM-9MPGT-3V66T` | `(Generic Key No Activation)` | Chỉ chuyển phiên bản, không kích hoạt |
| **GVLK Key** | `W269N-WFGWX-YVC9B-4J6C9-T83GX` | `(GVLK Key No Activation)` | Chỉ chuyển phiên bản, không kích hoạt |
| **Nhập key của bạn...** | (hiện ô nhập bên dưới) | `(Kích hoạt luôn sau khi chuyển)` | Chuyển phiên bản rồi kích hoạt luôn |

- **Generic Key** là key cài đặt chung (danh sách cộng đồng, không có trên trang KMS của Microsoft). **GVLK Key** là key KMS client do Microsoft công bố. Cả hai chỉ dùng để chọn phiên bản.
- Script dùng **đúng key bạn chọn**, không tự thử key khác. Nếu Windows từ chối với lỗi `0xC004F050` (key không hợp lệ cho bản cài này, ví dụ GVLK khi đi từ Home), log gợi ý chọn loại key khác.
- Phiên bản thiếu loại key nào thì ô chọn ẩn mục đó:
  - Generic Key có cho: Home, Home N, Home Single Language, Home Country Specific, Pro, Pro N, Pro for Workstations (N), Pro Education (N), Education (N), Enterprise (N), Enterprise G N.
  - GVLK có cho: Pro, Pro N, Pro for Workstations (N), Pro Education (N), Education (N), Enterprise (N), Enterprise G (N), Enterprise LTSC (N), IoT Enterprise LTSC.
  - Phiên bản không có cả hai (ví dụ IoT Enterprise, Enterprise multi-session, SE) chỉ còn mục "Nhập key của bạn".

**Khi chọn "Nhập key của bạn":**
1. Chạy `changepk.exe /ProductKey <key của bạn>` để chuyển phiên bản.
2. Nếu thành công, **kích hoạt luôn** với Microsoft (tương đương `slmgr /ato`).
3. Nếu Windows chưa kích hoạt được trước khi khởi động lại, log ghi rõ. Windows sẽ tự kích hoạt sau khi khởi động lại (cần Internet).

**Sau khi chuyển thành công** (dù đã kích hoạt hay chưa), script luôn **mở trang Settings > System > Activation** để bạn xem trạng thái hoặc nhập key bản quyền. Cần khởi động lại để hoàn tất; máy có thể tự khởi động lại.

### 4. Kiểm tra trạng thái kích hoạt
Nút **Kiểm tra kích hoạt** chỉ đọc thông tin, không thay đổi gì. Kết quả gồm:
- Tên và mô tả sản phẩm.
- Kênh key (Retail / OEM / Volume).
- 5 ký tự cuối của key.
- Trạng thái kích hoạt.
- Ngày hết hạn, hoặc "Vĩnh viễn".

### 5. Nhật ký (log)
Toàn bộ thao tác và output của DISM / changepk được hiển thị trong khung nhật ký, tô màu theo loại dòng:

| Dòng | Màu |
|---|---|
| `[+]` | xanh lá |
| `[!]` | đỏ |
| `[i]` | xanh dương |
| `> lệnh` | tím |
| Output của công cụ | xám |

## Công cụ sử dụng

Script chỉ dùng các thành phần **có sẵn trong Windows**:

| Công cụ | Dùng để |
|---|---|
| `dism.exe /Online /Get-TargetEditions` | Lấy danh sách phiên bản có thể chuyển |
| `changepk.exe /ProductKey <key>` | Chuyển sang phiên bản đã chọn, bằng key chọn trong ô Active key |
| WMI `SoftwareLicensingProduct` (`Get-CimInstance`) | Đọc trạng thái kích hoạt (thay cho `slmgr /dli` và `/xpr`, không cần VBScript) |
| WMI `SoftwareLicensingProduct.Activate` (`Invoke-CimMethod`) | Kích hoạt trực tuyến khi bạn nhập key bản quyền (tương đương `slmgr /ato`) |
| `ms-settings:activation` | Mở trang Settings > System > Activation sau khi chuyển |
| WPF (PresentationFramework) | Giao diện cửa sổ |
