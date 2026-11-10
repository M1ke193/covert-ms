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
- **Màu đỏ**: phiên bản DISM không liệt kê cho bản cài hiện tại, ví dụ hạ xuống Home, đổi giữa bản N và bản thường, từ LTSC hoặc bản Evaluation.
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
2. Nếu thành công, **kích hoạt luôn** với Microsoft (tương đương `slmgr /ato`, xem mục *Về kích hoạt*).
3. Nếu Windows chưa kích hoạt được trước khi khởi động lại, log ghi rõ. Windows sẽ tự kích hoạt sau khi khởi động lại (cần Internet).

**Sau khi chuyển thành công** (dù đã kích hoạt hay chưa), script luôn **mở trang Cài đặt > Hệ thống > Kích hoạt** để bạn xem trạng thái hoặc nhập key bản quyền. Cần khởi động lại để hoàn tất; máy có thể tự khởi động lại.

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
| `ms-settings:activation` | Mở trang Kích hoạt trong Cài đặt sau khi chuyển |
| WPF (PresentationFramework) | Giao diện cửa sổ |

### Về kích hoạt

Script **không dùng KMS**. Việc kích hoạt chỉ xảy ra với **key bản quyền bạn nhập**, theo các bước:

1. **Cài key và chuyển phiên bản:** `changepk.exe /ProductKey <key>`.
2. **Kích hoạt trực tuyến với Microsoft** (chỉ khi bạn nhập key, và bước 1 trả về mã 0 hoặc 3010):
   ```powershell
   Get-CimInstance -ClassName SoftwareLicensingProduct `
       -Filter "ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' AND PartialProductKey IS NOT NULL" |
       Invoke-CimMethod -MethodName Activate
   ```
   Nếu chưa kích hoạt được trước khi khởi động lại, Windows sẽ tự kích hoạt key đó sau khi khởi động lại (cần Internet).
3. **Mở trang Kích hoạt:** `Start-Process 'ms-settings:activation'`, luôn chạy sau khi chuyển thành công, để bạn xem trạng thái hoặc nhập key bản quyền.
4. **Đọc trạng thái kích hoạt** (nút Kiểm tra kích hoạt, chỉ đọc): `Get-CimInstance -ClassName SoftwareLicensingProduct …`.

Key chung (GVLK hoặc key cài đặt chung) **không bao giờ kích hoạt** Windows. Theo Microsoft, GVLK chỉ kích hoạt được khi máy kết nối với máy chủ KMS hợp lệ trong mạng của tổ chức.

Nếu muốn tự kích hoạt bằng lệnh, dùng công cụ có sẵn của Windows với **key bản quyền của bạn** (chạy PowerShell bằng quyền quản trị):

```powershell
slmgr.vbs /ipk XXXXX-XXXXX-XXXXX-XXXXX-XXXXX   # cài key bản quyền
slmgr.vbs /ato                                 # kích hoạt trực tuyến với Microsoft
slmgr.vbs /dlv                                 # xem chi tiết trạng thái
```