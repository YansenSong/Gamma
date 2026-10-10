param([Parameter(Mandatory=$true)][string]$WorkbookPath,[Parameter(Mandatory=$true)][string]$InputJson)
$ErrorActionPreference = 'Stop'
$excel=$null; $book=$null
try {
  $values=Get-Content -LiteralPath $InputJson -Raw -Encoding UTF8 | ConvertFrom-Json
  $excel=New-Object -ComObject Excel.Application
  $excel.Visible=$false; $excel.DisplayAlerts=$false; $excel.AskToUpdateLinks=$false
  $excel.AutomationSecurity=3
  $book=$excel.Workbooks.Open((Resolve-Path -LiteralPath $WorkbookPath).Path,0,$false)
  foreach ($entry in $values.PSObject.Properties) {
    $parts=$entry.Name.Split('!'); if ($parts.Count -ne 2) { throw "Invalid target $($entry.Name)" }
    $sheet=$book.Worksheets.Item($parts[0]); $cell=$sheet.Range($parts[1])
    if ($cell.HasFormula) { throw "Refusing to overwrite formula $($entry.Name)" }
    $value=$entry.Value; if ($value -is [ValueType] -and $value -isnot [bool]) { $value=[System.Convert]::ToDouble($value,[Globalization.CultureInfo]::InvariantCulture) }
    if ($value -is [string]) { $cell.Value2=[string]$value } else { $cell.Value2=[double]$value }
  }
  $excel.CalculateFullRebuild()
  $book.Save()
  $version=$excel.Version
  $book.Close($true); $book=$null
  $excel.Quit(); $excel=$null
  [pscustomobject]@{engine='Microsoft Excel';version=[string]$version}|ConvertTo-Json -Compress
} catch {
  if ($book) { try { $book.Close($false) } catch {} }
  if ($excel) { try { $excel.Quit() } catch {} }
  [Console]::Error.WriteLine("$($_.InvocationInfo.ScriptLineNumber): $($_.Exception.ToString())"); exit 2
}
