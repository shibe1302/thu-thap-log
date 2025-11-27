Add-Type -Path "C:\Program Files (x86)\WinSCP\WinSCPnet.dll"

$sessionOptions = New-Object WinSCP.SessionOptions
$sessionOptions.Protocol = [WinSCP.Protocol]::Sftp
$sessionOptions.HostName = "127.0.0.1"
$sessionOptions.PortNumber = 22
$sessionOptions.UserName = ""

$session = New-Object WinSCP.Session

try {
    $fingerprint = $session.ScanFingerprint($sessionOptions, "SHA-256")
    Write-Host "Fingerprint tim duoc:"
    Write-Host $fingerprint
}
finally {
    $session.Dispose()
}
