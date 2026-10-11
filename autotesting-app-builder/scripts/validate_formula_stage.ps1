param(
  [Parameter(Mandatory=$true)][string]$PlanPath,
  [Parameter(Mandatory=$true)][string]$StageManifestPath,
  [Parameter(Mandatory=$true)][string]$EvidencePath
)
$ErrorActionPreference = 'Stop'
function Get-Sha256([string]$Path) {
  $stream = [System.IO.File]::OpenRead($Path)
  $sha = [System.Security.Cryptography.SHA256]::Create()
  try { return ([System.BitConverter]::ToString($sha.ComputeHash($stream))).Replace('-', '').ToLowerInvariant() }
  finally { $sha.Dispose(); $stream.Dispose() }
}
$plan = Get-Content -LiteralPath $PlanPath -Raw -Encoding UTF8 | ConvertFrom-Json
$manifest = Get-Content -LiteralPath $StageManifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
$planDirectory = Split-Path -Parent (Resolve-Path -LiteralPath $PlanPath).Path
$workbookPath = if ([System.IO.Path]::IsPathRooted([string]$plan.workbook)) { [string]$plan.workbook } else { Join-Path $planDirectory $plan.workbook }
$source = (Resolve-Path -LiteralPath $workbookPath).Path
$staged = (Resolve-Path -LiteralPath $manifest.staged_path).Path
if ((Get-Sha256 $source) -ne $manifest.source_sha256) { throw 'Source hash changed' }
if ((Get-Sha256 $staged) -ne $manifest.staged_sha256) { throw 'Staged workbook hash changed' }
$excel = $null
try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $excel.AskToUpdateLinks = $false
  $excel.AutomationSecurity = 3
  $results = @()
  $index = 0
  foreach ($case in $plan.boundary_cases) {
    $trial = Join-Path ([System.IO.Path]::GetTempPath()) ("gamma-formula-trial-" + [guid]::NewGuid().ToString('N') + '.xlsx')
    Copy-Item -LiteralPath $staged -Destination $trial -ErrorAction Stop
    $book = $null
    try {
      $book = $excel.Workbooks.Open($trial, 0, $false)
      foreach ($input in $case.inputs) {
        $cell = $book.Worksheets.Item($input.sheet).Range($input.cell)
        if ($cell.HasFormula) { throw "Boundary input overwrites a formula: $($input.sheet)!$($input.cell)" }
        # Keep one numeric COM setter type across cases. Windows PowerShell's
        # binder can cache Int32 on early cases and fail on a later Decimal.
        $inputValue = if ($input.value -is [ValueType]) { [double]$input.value } else { [string]$input.value }
        try { $cell.Value2 = $inputValue }
        catch { throw "Boundary input $index $($input.sheet)!$($input.cell) ($($inputValue.GetType().FullName)): $($_.Exception.Message)" }
      }
      $excel.CalculateFullRebuild()
      $checks = @()
      foreach ($expectation in $case.expected) {
        $resultCell = $book.Worksheets.Item($expectation.sheet).Range($expectation.cell)
        if ($resultCell.Text -match '^#(N/A|VALUE!|DIV/0!|REF!|NUM!|NAME\?|NULL!)$') {
          throw "Boundary case $index produced Excel error at $($expectation.sheet)!$($expectation.cell): $($resultCell.Text)"
        }
        $actual = $resultCell.Value2
        $passed = $false
        if ($actual -is [ValueType] -and $expectation.value -is [ValueType]) {
          $passed = [math]::Abs([double]$actual - [double]$expectation.value) -le 0.000000001
        } else { $passed = [string]$actual -ceq [string]$expectation.value }
        $checks += [pscustomobject]@{sheet=$expectation.sheet;cell=$expectation.cell;expected=$expectation.value;actual=$actual;passed=$passed}
      }
      $passedCase = @($checks | Where-Object { -not $_.passed }).Count -eq 0
      $results += [pscustomobject]@{index=$index;status=$(if ($passedCase) {'passed'} else {'failed'});checks=$checks}
      $book.Close($false); $book = $null
    } finally {
      if ($book) { $book.Close($false) }
      Remove-Item -LiteralPath $trial -Force -ErrorAction SilentlyContinue
    }
    $index++
  }
  $status = if (@($results | Where-Object {$_.status -ne 'passed'}).Count -eq 0) {'passed'} else {'failed'}
  $evidence = [ordered]@{status=$status;engine='Microsoft Excel';excel_version=[string]$excel.Version;
    plan_sha256=$manifest.plan_sha256;source_sha256=$manifest.source_sha256;staged_sha256=$manifest.staged_sha256;
    timestamp_utc=[DateTime]::UtcNow.ToString('o');cases=$results}
  $evidence | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $EvidencePath -Encoding UTF8
  if ($status -ne 'passed') { throw 'One or more native Excel boundary cases failed' }
} finally {
  if ($excel) { $excel.Quit() }
}
