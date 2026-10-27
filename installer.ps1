$ErrorActionPreference = 'Stop'
$source = $PSScriptRoot
$installRoot = Join-Path $env:LOCALAPPDATA 'Programs\MemoireCorrecteurProGloss'
$venv = Join-Path $installRoot '.venv'

Write-Host '=== Installation MEMOIRE CORRECTEUR PRO gloss ===' -ForegroundColor Cyan
if (-not (Test-Path (Join-Path $source 'app\app.py'))) {
    throw 'Tsy hita ny app\app.py. Ataovy Extract All aloha ny ZIP, dia alefaso ao amin''ilay dossier nalaina ny installer.'
}
$pythonLauncher = Get-Command py.exe -ErrorAction SilentlyContinue
$pythonExe = Get-Command python.exe -ErrorAction SilentlyContinue
if (-not $pythonLauncher -and -not $pythonExe) {
    Write-Host 'Python 3 tsy hita. Apetraho Python 3.11 na vaovao kokoa, ary safidio Add Python to PATH.' -ForegroundColor Yellow
    Write-Host 'Rohy: https://www.python.org/downloads/windows/'
    throw 'Apetraho aloha ny Python, dia avereno ny installer.'
}

New-Item -ItemType Directory -Force -Path $installRoot | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $installRoot 'app') | Out-Null
Copy-Item -Path (Join-Path $source 'app\*') -Destination (Join-Path $installRoot 'app') -Recurse -Force
Copy-Item -Path (Join-Path $source 'requirements.txt') -Destination $installRoot -Force
Copy-Item -Path (Join-Path $source 'LANCER_APP.vbs') -Destination $installRoot -Force
Copy-Item -Path (Join-Path $source 'LANCER_WINDOWS.bat') -Destination $installRoot -Force
Copy-Item -Path (Join-Path $source 'LANCER_TELEPHONE_WIFI.bat') -Destination $installRoot -Force
if (Test-Path (Join-Path $source 'assets')) {
    Copy-Item -Path (Join-Path $source 'assets') -Destination $installRoot -Recurse -Force
}
if (Test-Path (Join-Path $source '.streamlit')) {
    Copy-Item -Path (Join-Path $source '.streamlit') -Destination $installRoot -Recurse -Force
}

if ((Test-Path $venv) -and ((-not (Test-Path (Join-Path $venv 'pyvenv.cfg'))) -or (-not (Test-Path (Join-Path $venv 'Scripts\python.exe'))))) {
    Write-Host 'Manadio .venv taloha tsy feno...' -ForegroundColor Yellow
    Remove-Item -Path $venv -Recurse -Force
}
if (-not (Test-Path (Join-Path $venv 'Scripts\python.exe'))) {
    Write-Host 'Mamorona tontolo Python...' -ForegroundColor Cyan
    if ($pythonLauncher) { & py -3 -m venv $venv } else { & python -m venv $venv }
    if ($LASTEXITCODE -ne 0) { throw 'Tsy nahomby ny famoronana .venv.' }
}
$venvPython = Join-Path $venv 'Scripts\python.exe'
$readyFile = Join-Path $venv 'dependencies_ready.txt'
if (-not (Test-Path $readyFile)) {
    Write-Host 'Mametraka bibliotheques. Mila Internet amin''ity dingana ity...' -ForegroundColor Cyan
    & $venvPython -m pip install --upgrade pip
    if ($LASTEXITCODE -ne 0) { throw 'Tsy nahomby ny fanavaozana pip.' }
    & $venvPython -m pip install -r (Join-Path $installRoot 'requirements.txt')
    if ($LASTEXITCODE -ne 0) { throw 'Tsy nahomby ny fametrahana requirements.txt.' }
    Set-Content -Path $readyFile -Value 'ready' -Encoding ASCII
}

$desktop = [Environment]::GetFolderPath('Desktop')
$shortcutPath = Join-Path $desktop 'MEMOIRE CORRECTEUR PRO gloss.lnk'
$wsh = New-Object -ComObject WScript.Shell
$shortcut = $wsh.CreateShortcut($shortcutPath)
$shortcut.TargetPath = Join-Path $env:WINDIR 'System32\wscript.exe'
$shortcut.Arguments = '"' + (Join-Path $installRoot 'LANCER_APP.vbs') + '"'
$shortcut.WorkingDirectory = $installRoot
$shortcut.Description = 'Lancer MEMOIRE CORRECTEUR PRO gloss'
$iconPath = Join-Path $installRoot 'assets\memoire_correcteur_pro.ico'
if (Test-Path $iconPath) { $shortcut.IconLocation = $iconPath } else { $shortcut.IconLocation = (Join-Path $env:WINDIR 'System32\shell32.dll') + ',23' }
$shortcut.Save()

Write-Host ''
Write-Host 'Vita ny installation!' -ForegroundColor Green
Write-Host "Dossier: $installRoot"
Write-Host "Raccourci: $shortcutPath"
Write-Host 'Ny fametrahana voalohany ihany no mila Internet mba hampidirana les dépendances.'
