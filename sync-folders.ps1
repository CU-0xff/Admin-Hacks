#sync-folders.ps1 - Syncs folders from one location to another using an CSV file as input
#usage: sync-folders.ps1 -csvfile <csvfile> -logfile <logfile> [-useHash]
#csvfile: CSV file with the following columns: source, destination / comma delimited UTF8 / with header

param(
    [Parameter(Mandatory=$true)]
    [string]$csvfile,
    [Parameter(Mandatory=$true)]
    [string]$logfile,
    [Parameter(Mandatory=$false)]
    [switch]$useHash
)

Write-Host "sync-folders.ps1 - Syncs folders from one location to another using an CSV file as input"

function WriteLog
{
    Param ([string]$LogString)
    $Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
    $LogMessage = "$Stamp $LogString"
    Add-content $LogFile -value $LogMessage
    Write-Output $LogMessage
}

Write-Host -NoNewLine "Testing logfile $logfile"

# Test if the log file is writable
try{
    WriteLog "Starting sync-folders.ps1"
}
catch {
    Write-Error -Message "Unable to write to log file $logfile"
    exit 1
}

Write-Host "OK"

Write-Host -NoNewLine "Opening CSV file $csvfile"
$csv = Import-Csv $csvfile -Delimiter "," -Encoding UTF8
Write-Host " OK"

foreach ($line in $csv) {
    Write-Host "Copy from $($line.source) to $($line.destination)"
    WriteLog "Copy from $($line.source) to $($line.destination)"
    $sourceFiles = Get-ChildItem $line.source -File -Recurse
    foreach ($sourceFileObject in $sourceFiles) {
        $sourceFile = $sourceFileObject.FullName.ToLower()
        $destinationFile = $sourceFile.Replace($line.source.ToLower(), $line.destination)
        $destinationFolder = $sourceFileObject.Directory.FullName.ToLower()
        $destinationFolder = ($destinationFolder).Replace($line.source.ToLower(), $line.destination)
        if (!(Test-Path $destinationFile)) {
            Write-Host "File $($sourceFile) does not exist in $($destinationFolder), copying"
            WriteLog "File $($sourceFile) does not exist in $($destinationFolder), copying"
            if(-not(Test-Path $destinationFolder)) {
                Write-Host "Folder $($destinationFolder) does not exist, creating"
                WriteLog "Folder $($destinationFolder) does not exist, creating"
                New-Item -ItemType Directory -Force -Path $destinationFolder
            }
            Copy-Item -Path $sourceFile -Destination $destinationFile -Recurse -Force -Verbose -ErrorAction SilentlyContinue -ErrorVariable errors
        }
        else {
            Write-Host "File $($sourceFile) exists in $($destinationFolder), skipping"
            WriteLog "File $($sourceFile) exists in $($destinationFolder), skipping"
        }
        if ($useHash) {
            $hash = Get-FileHash -Path $sourceFile -Algorithm MD5
            $hash2 = Get-FileHash -Path $destinationFile -Algorithm MD5
            if ($hash.hash -ne $hash2.hash) {
                Write-Host "File $($sourceFile) to $($destinationFile) - Hashes are different, copying"
                WriteLog "File $($sourceFile) to $($destinationFile) - Hashes are different, copying"
                Copy-Item -Path $sourceFile -Destination $destinationFolder -Recurse -Force -Verbose -ErrorAction SilentlyContinue -ErrorVariable errors
            }
            else {
                Write-Host "File $($sourceFile) to $($destinationFile) - Hashes are the same, skipping"
                WriteLog "File $($sourceFile) to $($destinationFile) - Hashes are the same, skipping"
            }
        }
    }
}
