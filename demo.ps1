


$ScpHost          = "127.0.0.1"
$ScpUser          = "shibe"
$ScpPassword      = "shibe1302"
$Protocol         = "Sftp"                          
$RemoteFolder     = "/ucg"                          
$LocalDestination = "E:\download_log"               
$winscpDllPath    = "C:\Program Files (x86)\WinSCP\WinSCPnet.dll"
$MacFilePath      = "E:\nghien_cuu_FTU\UCG_FIBER_40pcs_log\data.txt"          
$MaxThreads       = 10                             


if (-not (Test-Path $winscpDllPath)) { Write-Error "Khong tim thay WinSCPnet.dll tai: $winscpDllPath"; exit }
try { Add-Type -Path $winscpDllPath } catch { Write-Error "Loi load DLL WinSCP: $_"; exit }


if (-not (Test-Path $LocalDestination)) { New-Item -ItemType Directory -Path $LocalDestination | Out-Null }




$start = Get-Date
Write-Host "[1/4] Dang doc danh sach MAC..." -ForegroundColor Cyan
$MacDb = @{}

if (Test-Path $MacFilePath) {
    $RawMacs = Get-Content $MacFilePath
    foreach ($mac in $RawMacs) {
        $cleanMac = $mac.Trim().ToUpper()
        if (-not [string]::IsNullOrWhiteSpace($cleanMac)) {
            $MacDb[$cleanMac] = $true 
        }
    }
    Write-Host "-> Da nap $($MacDb.Count) MAC vao bo nho." -ForegroundColor Green
} else {
    Write-Error "Khong tim thay file danh sach MAC tai: $MacFilePath"; exit
}


function Get-MacFromFileName {
    param ($FileName)
    if ($FileName -match "(_[^_]+_)") {
        return $matches[1].Trim('_')
    }
    return $null
}




Write-Host "[2/4] Dang ket noi WinSCP va quet file..." -ForegroundColor Cyan

$FilesToDownload = [System.Collections.Generic.List[object]]::new()


$sessionOptions = New-Object WinSCP.SessionOptions
$sessionOptions.ParseUrl("$($Protocol.ToLower())://$($ScpUser):$($ScpPassword)@$($ScpHost)/")

$sessionOptions.GiveUpSecurityAndAcceptAnySshHostKey = $true 

$session = New-Object WinSCP.Session
$global:ScannedFiles

try {
    $session.Open($sessionOptions)

    
    function Get-WinSCPFilesRecursive ($Path) {
        try {
            
            $directoryInfo = $session.ListDirectory($Path)
            
            foreach ($fileInfo in $directoryInfo.Files) {
                
                if ($fileInfo.Name -eq "." -or $fileInfo.Name -eq "..") { continue }

                if ($fileInfo.IsDirectory) {
                    
                    Get-WinSCPFilesRecursive $fileInfo.FullName
                }
                else {
                    $global:ScannedFiles++
                    
                    $extractedMac = Get-MacFromFileName -FileName $fileInfo.Name
                    
                    if ($extractedMac -and $MacDb.ContainsKey($extractedMac)) {
                        
                        $FilesToDownload.Add(@{
                            RemotePath = $fileInfo.FullName
                            FileName   = $fileInfo.Name
                        })
                    }
                }
            }
        }
        catch {
            Write-Warning "Khong the truy cap folder: $Path. Loi: $_"
        }
    }

    
    Get-WinSCPFilesRecursive $RemoteFolder

}
finally {
    $session.Dispose()
}

$TotalFiles = $FilesToDownload.Count
Write-Host "-> Quet xong. Tim thay: $TotalFiles file." -ForegroundColor Yellow
if ($TotalFiles -eq 0) { exit }




Write-Host "[3/4] Dang khoi tao $MaxThreads luong tai xuong..." -ForegroundColor Cyan


$Batches = @()
$BatchSize = [Math]::Ceiling($TotalFiles / $MaxThreads)
for ($i = 0; $i -lt $TotalFiles; $i += $BatchSize) {
    $count = [Math]::Min($BatchSize, ($TotalFiles - $i))
    $Batches += ,$FilesToDownload.GetRange($i, $count)
}





$JobBlock = {
    param($FileBatch, $SessionOptsHash, $DllPath, $DestDir)
    
    Add-Type -Path $DllPath
    
    
    $jobOptions = New-Object WinSCP.SessionOptions
    $jobOptions.Protocol = [WinSCP.Protocol]::Sftp
    $jobOptions.HostName = $SessionOptsHash.HostName
    $jobOptions.UserName = $SessionOptsHash.UserName
    $jobOptions.Password = $SessionOptsHash.Password
    $jobOptions.GiveUpSecurityAndAcceptAnySshHostKey = $true

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
                        Write-Warning "Code 32 cho file $($f.FileName). $($DelaySeconds) giây... ($($i + 1)/$MaxRetries)"
                        Start-Sleep -Seconds $DelaySeconds
                    }
                    else {
                        
                        Write-Error "failed to download $($f.FileName): $($_.Exception.Message)"
                        break 
                    }
                }
            }
        }
    }
    catch {
        Write-Error "Job Session Error: $($_.Exception.Message)"
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


$Jobs = @()
foreach ($batch in $Batches) {
    $Jobs += Start-Job -ScriptBlock $JobBlock -ArgumentList $batch, $SessionOptsHash, $winscpDllPath, $LocalDestination
}


$TotalFiles = $FilesToDownload.Count

Write-Host "[4/4] Dang tai xuong..." -ForegroundColor Cyan
$Jobs | Wait-Job | Out-Null
$Jobs | Receive-Job
$Jobs | Remove-Job

Write-Host "HOAN TAT! Kiem tra folder: $LocalDestination" -ForegroundColor Green
$end = Get-Date
$duration = $end - $start
Write-Output "time: $duration" -ForegroundColor Green

Write-Host "-> Quet xong. Tim thay: $TotalFiles file hop le." -ForegroundColor Yellow
Write-Host "-> Tong so file da duyet qua: $ScannedFiles" -ForegroundColor Cyan