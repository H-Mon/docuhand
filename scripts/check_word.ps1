$procs = Get-Process WINWORD -ErrorAction SilentlyContinue
if ($procs) {
    $procs | ForEach-Object { Write-Output ("ORPHAN {0}" -f $_.Id) }
    $procs | Stop-Process -Force
    Write-Output "KILLED"
} else {
    Write-Output "no WINWORD running"
}
