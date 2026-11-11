$ErrorActionPreference = 'Stop'
$out = 'C:\Users\SIPC\Documents\CTEP Shift Report Archive\outlook_import'
$ol = New-Object -ComObject Outlook.Application
$ns = $ol.GetNamespace('MAPI')
$f = $ns.Folders.Item('f.meshkipour@petroyada.com').Folders.Item('Inbox').Folders.Item('1. Report').Folders.Item('C. Reports Archive')
$items = $f.Items
Write-Output ("MAIL_ITEMS=" + $items.Count)
$saved = 0
foreach ($m in $items) {
  if ($m.Class -ne 43) { continue }
  foreach ($a in $m.Attachments) {
    if ($a.FileName -match '^CWRS-.*\.(xlsx|xlsm|xls)$') {
      $dest = Join-Path $out $a.FileName
      if (Test-Path $dest) {
        $stem = [IO.Path]::GetFileNameWithoutExtension($a.FileName)
        $ext = [IO.Path]::GetExtension($a.FileName)
        $dest = Join-Path $out ($stem + '_' + $m.ReceivedTime.ToString('yyyyMMdd_HHmmss') + $ext)
      }
      $a.SaveAsFile($dest)
      $saved++
      Write-Output ("SAVED | " + $m.ReceivedTime.ToString('yyyy-MM-dd HH:mm') + " | " + $m.Subject + " | " + $a.FileName + " | " + $dest)
    }
  }
}
Write-Output ("SAVED_TOTAL=" + $saved)
