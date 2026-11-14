# auto_outlook_import.ps1
# Runs every 3 hours (via Windows Task Scheduler):
#   1) Saves CWRS-*.xlsx attachments from the Outlook folder to outlook_import\
#   2) Runs extractor.py which imports them into the SQLite archive
#      (already-imported reports are skipped automatically, so re-running is safe)
# Requires: Windows + Desktop Outlook installed & logged in.

$ErrorActionPreference = 'Stop'

# --- Paths (edit here if needed) -------------------------------------------
$scriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$outlookFolderName = 'outlook_import'                 # subfolder used by the app/extractor
$importRoot = Join-Path $scriptDir $outlookFolderName  # where attachments are saved
$logFile    = Join-Path $scriptDir 'logs\auto_outlook_import.log'

# --- Outlook mailbox / folder path (same as export_outlook_ctep.ps1) -------
$mailboxName = 'f.meshkipour@petroyada.com'
$folderPath  = @('Inbox', '1. Report', 'C. Reports Archive')

New-Item -ItemType Directory -Force -Path $importRoot | Out-Null
New-Item -ItemType Directory -Force -Path (Split-Path $logFile) | Out-Null

function Log($msg) {
    $line = ('{0} | {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $msg)
    Add-Content -Path $logFile -Value $line -Encoding UTF8
    Write-Output $line
}

try {
    Log 'Run started'

    # 1) Export new attachments from Outlook -------------------------------
    $ol  = New-Object -ComObject Outlook.Application
    $ns  = $ol.GetNamespace('MAPI')
    $fld = $ns.Folders.Item($mailboxName)
    foreach ($name in $folderPath) { $fld = $fld.Folders.Item($name) }

    $saved = 0
    foreach ($m in $fld.Items) {
        if ($m.Class -ne 43) { continue }                       # only mail items
        foreach ($a in $m.Attachments) {
            if ($a.FileName -match '^CWRS-.*\.(xlsx|xlsm)$') {
                $dest = Join-Path $importRoot $a.FileName
                if (Test-Path $dest) {
                    $stem = [IO.Path]::GetFileNameWithoutExtension($a.FileName)
                    $ext  = [IO.Path]::GetExtension($a.FileName)
                    $dest = Join-Path $importRoot ($stem + '_' + $m.ReceivedTime.ToString('yyyyMMdd_HHmmss') + $ext)
                    if (Test-Path $dest) { continue }           # already saved before -> skip
                }
                $a.SaveAsFile($dest)
                $saved++
                Log ('SAVED | ' + $m.ReceivedTime.ToString('yyyy-MM-dd HH:mm') + ' | ' + $m.Subject + ' | ' + $a.FileName)
            }
        }
    }
    Log ("Attachments saved this run: " + $saved)

    # 2) Import into the archive DB ----------------------------------------
    $py = Get-Command py -ErrorAction SilentlyContinue
    if ($py) { $pythonExe = 'py' ; $pythonArgs = @('-3', (Join-Path $scriptDir 'extractor.py')) }
    else     { $pythonExe = (Get-Command python).FullName ; $pythonArgs = @((Join-Path $scriptDir 'extractor.py')) }

    $output = & $pythonExe @pythonArgs 2>&1 | Out-String
    Add-Content -Path $logFile -Value $output -Encoding UTF8
    Log ("Extractor exit code: " + $LASTEXITCODE)
    Log 'Run finished OK'
}
catch {
    Log ('ERROR: ' + $_.Exception.Message)
    exit 1
}
