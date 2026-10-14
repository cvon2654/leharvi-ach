# Solver installer for Excel (Windows). Run it through Install-Solver.bat.
#
#  1. Copies Solver to %LOCALAPPDATA%\Solver (no admin rights needed)
#  2. Puts the Solver macros into your Personal Macro Workbook (PERSONAL.XLSB)
#  3. Adds an "Accounting Solver" tab to the Excel ribbon + Quick Access Toolbar buttons
#  4. Adds a "Solver" shortcut to your desktop
#
# Everything is offline. To remove it, run Uninstall-Solver.bat.

$ErrorActionPreference = "Stop"
$src = Split-Path -Parent $PSScriptRoot
$dest = Join-Path $env:LOCALAPPDATA "Solver"
$officeUI = Join-Path $env:LOCALAPPDATA "Microsoft\Office\Excel.officeUI"
$marker = "AccountingSolverTab"

function Say($text, $color = "White") { Write-Host $text -ForegroundColor $color }
function Fail($text) {
    Say ""
    Say "PROBLEM: $text" "Red"
    Say ""
    Read-Host "Press Enter to close"
    exit 1
}

Say ""
Say "=== Installing Solver for Excel ===" "Cyan"
Say ""

# --- 0. Excel must be closed ---------------------------------------------
while (Get-Process -Name EXCEL -ErrorAction SilentlyContinue) {
    Say "Please close Excel (save your work first), then press Enter." "Yellow"
    Read-Host | Out-Null
}

# --- 1. Copy the program --------------------------------------------------
Say "1/4  Copying Solver to $dest"
if (-not ((Test-Path (Join-Path $src "run_solver.py")) -or (Test-Path (Join-Path $src "Solver.exe")))) {
    Fail "Run this from inside the unzipped Solver folder (Solver.exe / run_solver.py not found)."
}
New-Item -ItemType Directory -Force $dest | Out-Null
foreach ($item in @("solver", "vendor", "excel", "run_solver.py", "Solver.bat", "README.md", "Solver.exe")) {
    $p = Join-Path $src $item
    if (Test-Path $p) {
        if ($src -ne $dest) { Copy-Item $p $dest -Recurse -Force }
    }
}

# How will the macro start Solver?  Solver.exe, or Python (pythonw.exe).
$python = ""
if (-not (Test-Path (Join-Path $dest "Solver.exe"))) {
    $candidates = @()
    foreach ($name in @("pythonw", "pyw")) {
        $cmd = Get-Command $name -ErrorAction SilentlyContinue
        if ($cmd) { $candidates += $cmd.Source }
    }
    foreach ($name in @("py", "python")) {
        $cmd = Get-Command $name -ErrorAction SilentlyContinue
        if ($cmd -and $cmd.Source -notlike "*WindowsApps*") {
            try {
                $exe = & $cmd.Source -c "import sys; print(sys.executable)" 2>$null
                if ($exe) { $candidates += ($exe.Trim() -replace "python\.exe$", "pythonw.exe") }
            } catch { }
        }
    }
    foreach ($c in $candidates) {
        if ($c -and (Test-Path $c) -and $c -notlike "*WindowsApps*") { $python = $c; break }
    }
    if (-not $python) {
        Say ""
        Say "Python was not found, and there is no Solver.exe in this folder." "Yellow"
        Say "Solver needs one of them. Easiest: install Python (free) from python.org," "Yellow"
        Say "tick 'Add python.exe to PATH', then run Install-Solver.bat again." "Yellow"
        Start-Process "https://www.python.org/downloads/"
        Fail "Python or Solver.exe is required."
    }
    # Check the bundled Excel library loads with this Python.
    $check = & ($python -replace "pythonw\.exe$", "python.exe") (Join-Path $dest "run_solver.py") --help 2>&1
    if ($LASTEXITCODE -ne 0) { Fail "Python found at $python but Solver would not start:`n$check" }
    Say "     Using Python: $python"
} else {
    Say "     Using Solver.exe"
}

# --- 2. Macro into PERSONAL.XLSB ------------------------------------------
Say "2/4  Adding the Solver macros to Excel"
$bas = Join-Path $env:TEMP "AccountingSolver.bas"
$code = Get-Content (Join-Path $dest "excel\AccountingSolver.bas") -Raw
$code = $code -replace 'SOLVER_FOLDER As String = "[^"]*"', ('SOLVER_FOLDER As String = "' + $dest + '"')
if ($python) {
    $code = $code -replace 'PYTHON_EXE As String = "[^"]*"', ('PYTHON_EXE As String = "' + $python + '"')
}
[System.IO.File]::WriteAllText($bas, $code, [System.Text.Encoding]::Default)

# Excel only lets a script add code when "Trust access to the VBA project
# object model" is on. Turn it on for the install, then put it back.
$restore = @{}
$officeKeys = Get-ChildItem "HKCU:\Software\Microsoft\Office" -ErrorAction SilentlyContinue |
    Where-Object { $_.PSChildName -match '^\d+\.\d+$' }
foreach ($k in $officeKeys) {
    $sec = Join-Path $k.PSPath "Excel\Security"
    if (-not (Test-Path (Join-Path $k.PSPath "Excel"))) { continue }
    New-Item -Path $sec -Force | Out-Null
    $old = (Get-ItemProperty -Path $sec -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM
    $restore[$sec] = $old
    Set-ItemProperty -Path $sec -Name AccessVBOM -Value 1 -Type DWord
}

$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
} catch {
    Fail "Could not start Excel. Is Microsoft Excel (desktop) installed?"
}
try {
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $startup = $excel.StartupPath
    New-Item -ItemType Directory -Force $startup | Out-Null
    $personal = Join-Path $startup "PERSONAL.XLSB"
    if (Test-Path $personal) {
        $wb = $excel.Workbooks.Open($personal)
    } else {
        $wb = $excel.Workbooks.Add()
        $wb.SaveAs($personal, 50)          # 50 = .xlsb
    }
    $components = $wb.VBProject.VBComponents
    foreach ($c in @($components)) {
        if ($c.Name -eq "AccountingSolver") { $components.Remove($c) }
    }
    $components.Import($bas) | Out-Null
    $wb.Windows.Item(1).Visible = $false   # PERSONAL.XLSB is always a hidden workbook
    $wb.Save()
    $wb.Close($true)
} catch {
    $msg = $_.Exception.Message
    try { $excel.Quit() } catch { }
    Fail "Excel would not accept the macro: $msg`nYou can add it by hand instead - see README.md, 'Use it inside Excel'."
} finally {
    if ($excel) {
        try { $excel.Quit() } catch { }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    foreach ($sec in $restore.Keys) {
        if ($null -eq $restore[$sec]) {
            Remove-ItemProperty -Path $sec -Name AccessVBOM -ErrorAction SilentlyContinue
        } else {
            Set-ItemProperty -Path $sec -Name AccessVBOM -Value $restore[$sec] -Type DWord
        }
    }
    Remove-Item $bas -ErrorAction SilentlyContinue
}
# Wait until Excel has fully exited before touching its ribbon file.
for ($i = 0; $i -lt 30 -and (Get-Process -Name EXCEL -ErrorAction SilentlyContinue); $i++) {
    Start-Sleep -Milliseconds 500
}

# --- 3. Ribbon tab + Quick Access Toolbar ----------------------------------
Say "3/4  Adding the 'Accounting Solver' tab to the ribbon"
$ns = 'xmlns:x1="http://schemas.microsoft.com/office/2009/07/customui/macro" xmlns:mso="http://schemas.microsoft.com/office/2009/07/customui"'
$tab = @"
<mso:tab id="mso_c1.$marker" label="Accounting Solver" insertBeforeQ="mso:TabInsert"><mso:group id="mso_c2.AccountingSolverGroup" label="Solve this workbook" autoScale="true"><mso:button idQ="x1:PERSONAL.XLSB_SolveStepByStep_0" label="Step by Step" imageMso="CalculateNow" onAction="PERSONAL.XLSB!SolveStepByStep" visible="true" size="large" screentip="Fill the answers one box at a time (OK = next box)"/><mso:button idQ="x1:PERSONAL.XLSB_SolveAutoPlay_0" label="Auto Play" imageMso="SlideShowFromBeginning" onAction="PERSONAL.XLSB!SolveAutoPlay" visible="true" size="large" screentip="Fill the answers one box at a time automatically"/><mso:button idQ="x1:PERSONAL.XLSB_SolveAllAtOnce_0" label="Solve All" imageMso="CalculateSheet" onAction="PERSONAL.XLSB!SolveAllAtOnce" visible="true" size="large" screentip="Fill every answer at once"/></mso:group></mso:tab>
"@
$qat = '<mso:button idQ="x1:PERSONAL.XLSB_SolveStepByStep_1" visible="true" label="Solver: step by step" onAction="PERSONAL.XLSB!SolveStepByStep" imageMso="CalculateNow"/>'

$ok = $true
if (-not (Test-Path $officeUI)) {
    New-Item -ItemType Directory -Force (Split-Path $officeUI) | Out-Null
    $xml = "<mso:customUI $ns><mso:ribbon><mso:qat><mso:sharedControls>$qat</mso:sharedControls></mso:qat><mso:tabs>$tab</mso:tabs></mso:ribbon></mso:customUI>"
} else {
    $xml = Get-Content $officeUI -Raw -Encoding UTF8
    if ($xml -notmatch $marker) {
        Copy-Item $officeUI "$officeUI.before-solver" -Force
        if ($xml -notmatch 'xmlns:x1=') { $xml = $xml -replace '<mso:customUI ', '<mso:customUI xmlns:x1="http://schemas.microsoft.com/office/2009/07/customui/macro" ' }
        if ($xml -match '</mso:tabs>') {
            $xml = $xml -replace '</mso:tabs>', "$tab</mso:tabs>"
        } elseif ($xml -match '</mso:ribbon>') {
            $xml = $xml -replace '</mso:ribbon>', "<mso:tabs>$tab</mso:tabs></mso:ribbon>"
        } elseif ($xml -match '</mso:customUI>') {
            $xml = $xml -replace '</mso:customUI>', "<mso:ribbon><mso:tabs>$tab</mso:tabs></mso:ribbon></mso:customUI>"
        } else { $ok = $false }
        if ($ok -and $xml -match '</mso:sharedControls>') {
            $xml = $xml -replace '</mso:sharedControls>', "$qat</mso:sharedControls>"
        }
    }
}
if ($ok) {
    [System.IO.File]::WriteAllText($officeUI, $xml, (New-Object System.Text.UTF8Encoding($false)))
} else {
    Say "     Could not add the ribbon tab automatically. Add the macro SolveStepByStep" "Yellow"
    Say "     via File > Options > Customize Ribbon > Macros." "Yellow"
}

# --- 4. Desktop shortcut ----------------------------------------------------
Say "4/4  Adding a 'Solver' shortcut to your desktop"
try {
    $shell = New-Object -ComObject WScript.Shell
    $lnk = $shell.CreateShortcut((Join-Path ([Environment]::GetFolderPath("Desktop")) "Solver.lnk"))
    $lnk.TargetPath = Join-Path $dest "Solver.bat"
    $lnk.WorkingDirectory = $dest
    $lnk.IconLocation = "$env:SystemRoot\System32\imageres.dll,109"
    $lnk.WindowStyle = 7
    $lnk.Description = "Solver - fills in Excel accounting problems"
    $lnk.Save()
} catch { Say "     (skipped the desktop shortcut)" "Yellow" }

Say ""
Say "Done!" "Green"
Say ""
Say "Open Excel and look for the 'Accounting Solver' tab:"
Say "  Step by Step - fills the answers one box at a time (OK = next box)"
Say "  Auto Play    - fills one box after another by itself"
Say "  Solve All    - fills everything at once"
Say ""
Say "Save your problem as .xlsx first, then click a button."
Say "You can also drag an Excel file onto the 'Solver' desktop shortcut."
Say ""
Read-Host "Press Enter to close"
