
$ScpHost = "127.0.0.1"
$ScpUser = "aa"
$ScpPassword = "aa"
$Protocol = "SFTP"                          
$RemoteFolder = "/ucg"                          
$LocalDestination = "E:\download_log"               
$winscpDllPath = "C:\Program Files (x86)\WinSCP\WinSCPnet.dll"
$MacFilePath = "E:\nghien_cuu_FTU\UCG_FIBER_40pcs_log\data.txt"          
$MaxScanThreads = 10  
$MaxDownloadThreads = 10  
$ConnectionTimeout = 30  
$Port = "22"


$Global:MacRegex = [regex]::new("(_[^_]+_)", [System.Text.RegularExpressions.RegexOptions]::Compiled)
function Validate-Configuration {
    Write-Host "[Validation] Checking configuration..." -ForegroundColor Cyan
    
    
    $ValidProtocols = @("SFTP", "Scp", "Ftp", "Ftps", "Webdav", "S3")
    if ($Protocol -notin $ValidProtocols) {
        [System.Windows.Forms.MessageBox]::Show(
            "Loi Protocol: '$Protocol' khong hop le!`n`nProtocol hop le: $($ValidProtocols -join ', ')",
            "Configuration Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        exit 1
    }
    
    
    if (-not (Test-Path $winscpDllPath)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Loi: khong tìm thay WinSCP DLL tai:`n$winscpDllPath`n`nVui long kiem tra duong dan cai dat WinSCP.",
            "File Not Found",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        exit 1
    }
    
    
    if (-not (Test-Path $MacFilePath)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Loi: khong tìm thay file MAC list tai:`n$MacFilePath`n`nVui long kiem tra duong dan file.",
            "File Not Found",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        exit 1
    }
    
    
    if (-not (Test-Path $LocalDestination)) {
        try {
            New-Item -ItemType Directory -Path $LocalDestination -Force | Out-Null
            Write-Host "Tạo folder đích: $LocalDestination" -ForegroundColor Green
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Loi: khong thể tạo folder đích:`n$LocalDestination`n`nLoi: $($_.Exception.Message)",
                "Directory Creation Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            exit 1
        }
    }
    else {
        Write-Host "Folder đích tồn tai: $LocalDestination" -ForegroundColor Green
    }
    
    
    try {
        $testFile = Join-Path $LocalDestination ".write_test_$([System.IO.Path]::GetRandomFileName())"
        [System.IO.File]::WriteAllText($testFile, "test")
        Remove-Item $testFile -Force
        Write-Host "Có quyen ghi vào folder đích" -ForegroundColor Green
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Loi: khong có quyen ghi vào folder đích:`n$LocalDestination`n`nLoi: $($_.Exception.Message)",
            "Permission Denied",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        exit 1
    }
    
    
    if ([string]::IsNullOrWhiteSpace($ScpHost)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Loi: Host/IP khong duoc de trống!",
            "Configuration Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        exit 1
    }
    
    
    if ([string]::IsNullOrWhiteSpace($ScpUser) -or [string]::IsNullOrWhiteSpace($ScpPassword)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Loi: Tên nguoi dung hoac mat khau khong duoc de trống!",
            "Configuration Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        exit 1
    }
    
    
    if ([string]::IsNullOrWhiteSpace($RemoteFolder)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Loi: thu muc remote khong duoc de trống!",
            "Configuration Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        exit 1
    }
    
    Write-Host "Tất cả cau hinh hop le`n" -ForegroundColor Green
}


function Load-WinSCPDll {
    Write-Host "[Loading] dang tai WinSCP DLL..." -ForegroundColor Cyan
    try {
        Add-Type -Path $winscpDllPath
        Write-Host "WinSCP DLL tai thanh cong" -ForegroundColor Green
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Loi: khong thể tai WinSCP DLL:`n`nLoi: $($_.Exception.Message)`n`ndam bao WinSCP đã được cai dat dung cách.",
            "DLL Loading Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        exit 1
    }
}


function Test-ServerConnection {
    Write-Host "[Connection] dang kiem tra ket noi server..." -ForegroundColor Cyan
    
    try {
        $sessionOptions = New-Object WinSCP.SessionOptions
        $sessionOptions.Protocol = [WinSCP.Protocol]::$Protocol
        $sessionOptions.HostName = $ScpHost
        $sessionOptions.UserName = $ScpUser
        $sessionOptions.Password = $ScpPassword
        $sessionOptions.Timeout = New-TimeSpan -Seconds $ConnectionTimeout
        $sessionOptions.GiveUpSecurityAndAcceptAnySshHostKey = $true
        $sessionOptions.PortNumber = $Port
        
        $testSession = New-Object WinSCP.Session
        
        try {
            $testSession.Open($sessionOptions)
            Write-Host "ket noi server thanh cong" -ForegroundColor Green
            return $true
        }
        catch {
            $errorMsg = $_.Exception.Message
            
            if ($errorMsg -match "timed out|Timeout|Time out") {
                [System.Windows.Forms.MessageBox]::Show(
                    "Loi Timeout: Request vuot quá $ConnectionTimeout giay.`n`nkiem tra:`n- Host/IP: $ScpHost`n- ket noi mạng`n- Firewall settings",
                    "Connection Timeout",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
            }
            elseif ($errorMsg -match "refused|Refused|refused to connect") {
                [System.Windows.Forms.MessageBox]::Show(
                    "Loi: Server tu choi ket noi.`n`nkiem tra:`n- Host/IP có dung khong: $ScpHost`n- Port có mở khong`n- Server dang chạy?",
                    "Connection Refused",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
            }
            elseif ($errorMsg -match "Authentication failed|auth|denied|password|username") {
                [System.Windows.Forms.MessageBox]::Show(
                    "Loi xac thuc: Tên nguoi dung hoac mat khau khong dung!`n`nVui long kiem tra:`n- Username: $ScpUser`n- Password`n- Protocol: $Protocol",
                    "Authentication Failed",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
            }
            else {
                [System.Windows.Forms.MessageBox]::Show(
                    "Loi ket noi server:`n`n$errorMsg`n`nHost: $ScpHost`nProtocol: $Protocol",
                    "Connection Error",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
            }
            return $false
        }
        finally {
            $testSession.Dispose()
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Loi khong mong muốn khi kiem tra ket noi:`n`n$($_.Exception.Message)",
            "Unexpected Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return $false
    }
}


function Test-RemoteFolder {
    param($Session)
    
    Write-Host "[RemoteFolder] kiem tra thu muc remote: $RemoteFolder" -ForegroundColor Cyan
    
    try {
        $fileInfos = $Session.EnumerateRemoteFiles($RemoteFolder, $null, [WinSCP.EnumerationOptions]::None)
        $fileInfos | Select-Object -First 1 | Out-Null
        Write-Host "thu muc remote OK" -ForegroundColor Green
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        
        [System.Windows.Forms.MessageBox]::Show(
            "Loi: khong thể truy cap thu muc remote!`n`nthu muc: $RemoteFolder`n`nLoi: $errorMsg`n`nkiem tra:`n- duong dan có dung khong`n- thu muc có tồn tai khong`n- Có quyen truy cap khong",
            "Remote Folder Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return $false
    }
}


function Get-MacFromFileName {
    param ($FileName)
    $match = $Global:MacRegex.Match($FileName)
    if ($match.Success) {
        return $match.Groups[1].Value.Trim('_')
    }
    return $null
}



Write-Host "`n========== SCP/SFTP PARALLEL FILE SCANNER - OPTIMIZED VERSION ==========" -ForegroundColor Magenta
Validate-Configuration
Load-WinSCPDll

if (-not (Test-ServerConnection)) {
    Write-Host "Validation failed. Exiting." -ForegroundColor Red
    exit 1
}
Write-Host ""

if (-not (Test-Path $LocalDestination)) { 
    New-Item -ItemType Directory -Path $LocalDestination | Out-Null 
}



$start = Get-Date
Write-Host "[1/5] Dang doc danh sach MAC..." -ForegroundColor Cyan
$MacDb = @{}

try {
    if (Test-Path $MacFilePath) {
        $RawMacs = Get-Content $MacFilePath -ErrorAction Stop
        foreach ($mac in $RawMacs) {
            $cleanMac = $mac.Trim().ToUpper()
            if (-not [string]::IsNullOrWhiteSpace($cleanMac)) {
                $MacDb[$cleanMac] = $true 
            }
        }
        if ($MacDb.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                "canh bao: File MAC list khong chứa du lieu hop le!`n`nFile: $MacFilePath",
                "Empty MAC List",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            exit 1
        }
        Write-Host "-> Da nap $($MacDb.Count) MAC vao bo nho." -ForegroundColor Green
    } 
    else {
        throw "khong tìm thay file MAC list tai: $MacFilePath"
    }
}
catch {
    [System.Windows.Forms.MessageBox]::Show(
        "Loi khi đọc file MAC list:`n`n$($_.Exception.Message)",
        "MAC File Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    exit 1
}



Write-Host "[2/5] Dang lay danh sach folder con trong $RemoteFolder..." -ForegroundColor Cyan

$RootFolders = @()
$sessionOptions = New-Object WinSCP.SessionOptions
$sessionOptions.Protocol = [WinSCP.Protocol]::$Protocol
$sessionOptions.HostName = $ScpHost
$sessionOptions.UserName = $ScpUser
$sessionOptions.Password = $ScpPassword
$sessionOptions.GiveUpSecurityAndAcceptAnySshHostKey = $true
$sessionOptions.PortNumber = $Port

$session = New-Object WinSCP.Session

try {
    $session.Open($sessionOptions)
    
    $directoryInfo = $session.ListDirectory($RemoteFolder)
    
    foreach ($item in $directoryInfo.Files) {
        if ($item.Name -eq "." -or $item.Name -eq "..") { continue }
        
        if ($item.IsDirectory) {
            $RootFolders += $item.FullName
        }
    }
    
    Write-Host "-> Tim thay $($RootFolders.Count) folder con" -ForegroundColor Green
    
    if ($RootFolders.Count -eq 0) {
        Write-Host "-> Khong co folder con, se scan truc tiep trong $RemoteFolder" -ForegroundColor Yellow
        $RootFolders = @($RemoteFolder)
    }
}
catch {
    Write-Error "Loi khi lay danh sach folder: $($_.Exception.Message)"
    exit 1
}
finally {
    $session.Dispose()
}



Write-Host "[3/5] Dang khoi tao scan song song voi toi da $MaxScanThreads luong..." -ForegroundColor Cyan


$FolderBatches = @()
$BatchSize = [Math]::Ceiling($RootFolders.Count / $MaxScanThreads)

for ($i = 0; $i -lt $RootFolders.Count; $i += $BatchSize) {
    $count = [Math]::Min($BatchSize, ($RootFolders.Count - $i))
    $batch = $RootFolders[$i..($i + $count - 1)]
    $FolderBatches += , $batch
}

Write-Host "-> Chia thanh $($FolderBatches.Count) batch de xu ly" -ForegroundColor Cyan



$ScanJobBlock = {
    param($FolderList, $MacDbKeys, $SessionOptsHash, $DllPath)
    
    
    Add-Type -Path $DllPath
    
    
    $LocalResults = @{}
    $ScannedCount = 0
    
    
    $jobOptions = New-Object WinSCP.SessionOptions
    $jobOptions.Protocol = [WinSCP.Protocol]::Sftp
    $jobOptions.HostName = $SessionOptsHash.HostName
    $jobOptions.UserName = $SessionOptsHash.UserName
    $jobOptions.Password = $SessionOptsHash.Password
    $jobOptions.GiveUpSecurityAndAcceptAnySshHostKey = $true
    $jobOptions.PortNumber = $Port
    
    $jobSession = New-Object WinSCP.Session
    
    
    function Scan-FolderRecursive {
        param($Path, $Session, $MacKeys, $Results, [ref]$Counter)
        
        try {
            $dirInfo = $Session.ListDirectory($Path)
            
            foreach ($fileInfo in $dirInfo.Files) {
                if ($fileInfo.Name -eq "." -or $fileInfo.Name -eq "..") { continue }
                
                if ($fileInfo.IsDirectory) {
                    Scan-FolderRecursive -Path $fileInfo.FullName -Session $Session -MacKeys $MacKeys -Results $Results -Counter $Counter
                }
                else {
                    $Counter.Value++
                    
                    
                    if ($fileInfo.Name -match "(_[^_]+_)") {
                        $extractedMac = $matches[1].Trim('_').ToUpper()
                        
                        if ($MacKeys -contains $extractedMac) {
                            $Results[$fileInfo.FullName] = $fileInfo.Name
                        }
                    }
                }
            }
        }
        catch {
            Write-Warning "Job: Khong the truy cap folder: $Path"
        }
    }
    
    try {
        $jobSession.Open($jobOptions)
        
        
        foreach ($folder in $FolderList) {
            $counter = 0
            Scan-FolderRecursive -Path $folder -Session $jobSession -MacKeys $MacDbKeys -Results $LocalResults -Counter ([ref]$counter)
            $ScannedCount += $counter
        }
        
        
        return @{
            Files        = $LocalResults
            ScannedCount = $ScannedCount
        }
    }
    catch {
        Write-Error "Job Session Error: $($_.Exception.Message)"
        return @{
            Files        = @{}
            ScannedCount = 0
        }
    }
    finally {
        $jobSession.Dispose()
    }
}


$SessionOptsHash = @{
    HostName = $ScpHost
    UserName = $ScpUser
    Password = $ScpPassword
}
$MacDbKeys = $MacDb.Keys


$ScanJobs = @()
$jobIndex = 0
foreach ($batch in $FolderBatches) {
    $jobIndex++
    Write-Host "   -> Khoi tao Job #$jobIndex voi $($batch.Count) folder(s)" -ForegroundColor Gray
    $ScanJobs += Start-Job -ScriptBlock $ScanJobBlock -ArgumentList $batch, $MacDbKeys, $SessionOptsHash, $winscpDllPath
}


Write-Host "-> Dang cho cac job hoan thanh..." -ForegroundColor Cyan
$ScanJobs | Wait-Job | Out-Null


$FilesToDownload = [System.Collections.Generic.List[object]]::new()
$TotalScanned = 0

foreach ($job in $ScanJobs) {
    $result = Receive-Job -Job $job
    
    if ($result -and $result.Files) {
        $TotalScanned += $result.ScannedCount
        
        foreach ($key in $result.Files.Keys) {
            $FilesToDownload.Add(@{
                    RemotePath = $key
                    FileName   = $result.Files[$key]
                })
        }
    }
}


$ScanJobs | Remove-Job

$TotalFiles = $FilesToDownload.Count
Write-Host "-> Scan hoan tat!" -ForegroundColor Green
Write-Host "   - Tong so file da quet: $TotalScanned" -ForegroundColor Yellow
Write-Host "   - File khop MAC: $TotalFiles" -ForegroundColor Yellow

if ($TotalFiles -eq 0) { 
    Write-Host "Khong co file nao de tai xuong. Ket thuc." -ForegroundColor Yellow
    exit 
}



Write-Host "[4/5] Dang khoi tao $MaxDownloadThreads luong tai xuong..." -ForegroundColor Cyan


$DownloadBatches = @()
$DownloadBatchSize = [Math]::Ceiling($TotalFiles / $MaxDownloadThreads)
for ($i = 0; $i -lt $TotalFiles; $i += $DownloadBatchSize) {
    $count = [Math]::Min($DownloadBatchSize, ($TotalFiles - $i))
    $DownloadBatches += , $FilesToDownload.GetRange($i, $count)
}


$DownloadJobBlock = {
    param($FileBatch, $SessionOptsHash, $DllPath, $DestDir)
    
    Add-Type -Path $DllPath
    
    $jobOptions = New-Object WinSCP.SessionOptions
    $jobOptions.Protocol = [WinSCP.Protocol]::Sftp
    $jobOptions.HostName = $SessionOptsHash.HostName
    $jobOptions.UserName = $SessionOptsHash.UserName
    $jobOptions.Password = $SessionOptsHash.Password
    $jobOptions.GiveUpSecurityAndAcceptAnySshHostKey = $true
    $jobOptions.PortNumber = $Port

    $jobSession = New-Object WinSCP.Session
    
    $MaxRetries = 3
    $DelaySeconds = 5
    
    try {
        $jobSession.Open($jobOptions)
        
        foreach ($f in $FileBatch) {
            $localFilePath = Join-Path $DestDir $f.FileName
            
            for ($i = 0; $i -lt $MaxRetries; $i++) {
                try {
                    if (Test-Path $localFilePath) { 
                        Write-Host "File exist: $($f.FileName)"
                        break 
                    }
                    
                    $transferResult = $jobSession.GetFiles($f.RemotePath, $localFilePath, $False)
                    $transferResult.Check()
                    break 
                }
                catch {
                    if ($_.Exception.Message -match "Code: 32") {
                        Write-Warning "Code 32 cho file $($f.FileName). Cho $($DelaySeconds)s... ($($i + 1)/$MaxRetries)"
                        Start-Sleep -Seconds $DelaySeconds
                    }
                    else {
                        Write-Error "Failed download $($f.FileName): $($_.Exception.Message)"
                        break 
                    }
                }
            }
        }
    }
    catch {
        Write-Error "Download Job Error: $($_.Exception.Message)"
    }
    finally {
        $jobSession.Dispose()
    }
}


$DownloadJobs = @()
foreach ($batch in $DownloadBatches) {
    $DownloadJobs += Start-Job -ScriptBlock $DownloadJobBlock -ArgumentList $batch, $SessionOptsHash, $winscpDllPath, $LocalDestination
}

Write-Host "[5/5] Dang tai xuong..." -ForegroundColor Cyan
$DownloadJobs | Wait-Job | Out-Null
$DownloadJobs | Receive-Job
$DownloadJobs | Remove-Job



Write-Host "`n========================================" -ForegroundColor Green
Write-Host "HOAN TAT! Kiem tra folder: $LocalDestination" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green

$end = Get-Date
$duration = $end - $start

Write-Host "`nThong ke:" -ForegroundColor Cyan
Write-Host "  - Tong so file da quet: $TotalScanned" -ForegroundColor Cyan 
Write-Host "  - File khop MAC: $TotalFiles" -ForegroundColor Cyan
Write-Host "  - Scan threads: $($FolderBatches.Count)" -ForegroundColor Cyan
Write-Host "  - Download threads: $($DownloadBatches.Count)" -ForegroundColor Cyan
Write-Host "  - Thoi gian thuc hien: $($duration.ToString('hh\:mm\:ss'))" -ForegroundColor Cyan