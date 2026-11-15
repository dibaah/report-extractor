$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$out = Join-Path $scriptDir 'outlook_import'
$logDir = Join-Path $scriptDir 'logs'
$logFile = Join-Path $logDir 'outlook_attachment_export.log'
$mailboxName = 'f.meshkipour@petroyada.com'
$folderPath = @('Inbox', '1. Report', 'C. Reports Archive')

New-Item -ItemType Directory -Force -Path $out | Out-Null
New-Item -ItemType Directory -Force -Path $logDir | Out-Null

function Log($message) {
    $line = '{0} | {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $message
    Add-Content -LiteralPath $logFile -Value $line -Encoding UTF8
    Write-Host $line
}

try {
    Log 'Outlook attachment export started'
    Log ("Destination: " + $out)

    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace('MAPI')
    $folder = $ns.Folders.Item($mailboxName)
    foreach ($name in $folderPath) {
        $folder = $folder.Folders.Item($name)
    }

    $saved = 0
    $skipped = 0
    foreach ($mail in $folder.Items) {
        if ($mail.Class -ne 43) { continue }
        foreach ($attachment in $mail.Attachments) {
            if ($attachment.FileName -notmatch '^(CWRS|CSR)-.*\.(xlsx|xlsm)$') { continue }

            $destination = Join-Path $out $attachment.FileName
            if (Test-Path -LiteralPath $destination) {
                $stem = [IO.Path]::GetFileNameWithoutExtension($attachment.FileName)
                $extension = [IO.Path]::GetExtension($attachment.FileName)
                $received = $mail.ReceivedTime.ToString('yyyyMMdd_HHmmss')
                $destination = Join-Path $out ($stem + '_' + $received + $extension)
                if (Test-Path -LiteralPath $destination) {
                    $skipped++
                    continue
                }
            }

            $attachment.SaveAsFile($destination)
            $saved++
            Log ("SAVED | " + $mail.ReceivedTime.ToString('yyyy-MM-dd HH:mm') + " | " + $attachment.FileName)
        }
    }

    Log ("Saved this run: " + $saved)
    Log ("Already present: " + $skipped)
    Log 'Outlook attachment export completed. No database import was performed.'
    exit 0
}
catch {
    Log ("ERROR: " + $_.Exception.Message)
    exit 1
}
