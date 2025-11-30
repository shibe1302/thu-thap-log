# --- CẤU HÌNH ĐƯỜNG DẪN DLL ---
# Hãy trỏ đúng đường dẫn đến file WinSCPnet.dll trên máy bạn
try {
    Add-Type -Path "C:\Program Files (x86)\WinSCP\WinSCPnet.dll"
}
catch {
    Write-Error "Không tìm thấy WinSCPnet.dll. Vui lòng cài đặt WinSCP hoặc sửa đường dẫn."
    exit
}

# --- HÀM CỦA BẠN (Đã tối ưu hóa một chút để tái sử dụng Regex) ---
$Global:MacRegex = [regex]::new("(_[^_]+_)", [System.Text.RegularExpressions.RegexOptions]::Compiled)

function Get-MacFromFileName {
    param (
        [string]$FileName
    )
    $match = $Global:MacRegex.Match($FileName)
    if ($match.Success) {
        return $match.Groups[1].Value.Trim('_')
    }
    return $null
}

# --- THÔNG TIN KẾT NỐI ---
$sessionOptions = New-Object WinSCP.SessionOptions
$sessionOptions.Protocol = [WinSCP.Protocol]::Sftp
$sessionOptions.HostName = "127.0.0.1"  # <--- Thay IP Server
$sessionOptions.UserName = "a"       # <--- Thay User
$sessionOptions.Password = "a"       # <--- Thay Pass
$sessionOptions.GiveUpSecurityAndAcceptAnySshHostKey = $true

# --- INPUT TỪ NGƯỜI DÙNG ---
$InputMac = "1C0B8B181B38"  # Địa chỉ MAC bạn muốn tìm
$RemotePath = "tess2/ucg" # Thư mục gốc trên server muốn quét

# --- BẮT ĐẦU XỬ LÝ ---
$session = New-Object WinSCP.Session

$start = Get-Date
try {
    # Kết nối
    Write-Host "Dang ket noi..."
    $session.Open($sessionOptions)

    # Cấu hình tìm kiếm:
    # 1. Đường dẫn gốc
    # 2. Mask: "*MAC*" -> Giúp lọc nhanh từ phía Server (Performance tốt hơn)
    # 3. AllDirectories: Tìm đệ quy trong subfolder
    $mask = "*$InputMac*" 
    $options = [WinSCP.EnumerationOptions]::AllDirectories

    Write-Host "Dang tim kiem file chua MAC: $InputMac trong $RemotePath (bao gom thu muc con)..."
    
    # Gọi hàm EnumerateRemoteFiles theo tài liệu
    $files = $session.EnumerateRemoteFiles($RemotePath, $mask, $options)

    $count = 0
    
    # Duyệt qua các file tìm được (IEnumerable)
    foreach ($fileInfo in $files) {
        
        # Lấy tên file
        $currentFileName = $fileInfo.Name
        
        # Dùng hàm của bạn để trích xuất MAC chính xác từ tên file để so sánh
        # (Bước này giúp đảm bảo đúng định dạng _MAC_ như regex yêu cầu)
        $extractedMac = Get-MacFromFileName -FileName $currentFileName
        
        if ($extractedMac -eq $InputMac) {
            Write-Host "[TIM THAY] File: $($fileInfo.FullName)" -ForegroundColor Green
            # Bạn có thể thêm lệnh tải file về ở đây nếu cần. Ví dụ:
            # $session.GetFiles($fileInfo.FullName, "C:\Local\Path\").Check()
            $count++
        }
    }

    if ($count -eq 0) {
        Write-Host "Khong tim thay file nao khop voi MAC: $InputMac" -ForegroundColor Yellow
    } else {
        Write-Host "Tong cong tim thay: $count file." -ForegroundColor Cyan
    }

}
catch {
    Write-Error "Co loi xay ra: $($_.Exception.Message)"
}
finally {
    # Luôn đóng session
    $session.Dispose()
}
$duration = $end - $start
Write-Host "  - Thoi gian thuc hien: $($duration.ToString('hh\:mm\:ss'))" -ForegroundColor Cyan