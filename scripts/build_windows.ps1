 #Requires -Version 5.1

# Do NOT "Stop" on stderr from native apps (PyInstaller writes INFO/WARN there).
$ErrorActionPreference = "Continue"
$ProgressPreference = "SilentlyContinue"

$AppName    = "Rex Excel Filter"
$Entrypoint = "excelfilter.py"
$Version    = "1.0.1"

$IconIco = "Assets\rexexcelfilter.ico"

# Windows add-data uses ; as the separator
$TkdndPath = (py -c "import tkinterdnd2, pathlib; print(pathlib.Path(tkinterdnd2.__file__).parent/'tkdnd')").Trim()

$AddData = @(
  "Assets\Rexie.png;Assets",
  "$TkdndPath;tkinterdnd2\\tkdnd"
)

$HiddenImports = @(
  "tkinterdnd2"
)

$OutDir  = "$env:USERPROFILE\dev\builds\Windows\RexExcelFilter"
$ZipBase = "RexExcelFilter"

Remove-Item -Recurse -Force build, dist -ErrorAction SilentlyContinue
Remove-Item -Force *.spec -ErrorAction SilentlyContinue

$PyArgs = @(
  "--noconfirm",
  "--clean",
  "--windowed",
  "--name", $AppName,
  "--icon", $IconIco
)

foreach ($d in $AddData)       { $PyArgs += @("--add-data", $d) }
foreach ($h in $HiddenImports) { $PyArgs += @("--hidden-import", $h) }

# Run PyInstaller and capture output, then explicitly check exit code.
$logDir = Join-Path (Get-Location) "logs"
New-Item -ItemType Directory -Force -Path $logDir | Out-Null
$logPath = Join-Path $logDir "pyinstaller_windows.log"

Write-Host "Building with PyInstaller..."
py -m PyInstaller @PyArgs $Entrypoint *>&1 | Tee-Object -FilePath $logPath
if ($LASTEXITCODE -ne 0) {
  Write-Host "❌ PyInstaller failed. See log: $logPath"
  exit $LASTEXITCODE
}

New-Item -ItemType Directory -Force -Path $OutDir | Out-Null

$ZipName  = "$ZipBase-$Version-Windows.zip"
$DistPath = Join-Path (Get-Location) "dist"
$ZipPath  = Join-Path $OutDir $ZipName

if (Test-Path $ZipPath) { Remove-Item -Force $ZipPath }

Compress-Archive -Path (Join-Path $DistPath "*") -DestinationPath $ZipPath -Force
Write-Host "✅ Built: $ZipPath"
 
