# Removes what install.ps1 added. Run it through Uninstall-Solver.bat.
$ErrorActionPreference = "Continue"
$dest = Join-Path $env:LOCALAPPDATA "Solver"
$officeUI = Join-Path $env:LOCALAPPDATA "Microsoft\Office\Excel.officeUI"

while (Get-Process -Name EXCEL -ErrorAction SilentlyContinue) {
    Write-Host "Please close Excel, then press Enter." -ForegroundColor Yellow
    Read-Host | Out-Null
}

# Ribbon tab and toolbar button
if (Test-Path $officeUI) {
    $xml = Get-Content $officeUI -Raw -Encoding UTF8
    $xml = $xml -replace '<mso:tab id="mso_c1\.AccountingSolverTab".*?</mso:tab>', ''
    $xml = $xml -replace '<mso:button idQ="x1:PERSONAL\.XLSB_SolveStepByStep_1"[^>]*/>', ''
    $xml = $xml -replace '<mso:tabs></mso:tabs>', ''
    [System.IO.File]::WriteAllText($officeUI, $xml, (New-Object System.Text.UTF8Encoding($false)))
}

# Macro module
$restore = @{}
Get-ChildItem "HKCU:\Software\Microsoft\Office" -ErrorAction SilentlyContinue |
    Where-Object { $_.PSChildName -match '^\d+\.\d+$' -and (Test-Path (Join-Path $_.PSPath "Excel")) } |
    ForEach-Object {
        $sec = Join-Path $_.PSPath "Excel\Security"
        New-Item -Path $sec -Force | Out-Null
        $restore[$sec] = (Get-ItemProperty -Path $sec -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM
        Set-ItemProperty -Path $sec -Name AccessVBOM -Value 1 -Type DWord
    }
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.DisplayAlerts = $false
    $personal = Join-Path $excel.StartupPath "PERSONAL.XLSB"
    if (Test-Path $personal) {
        $wb = $excel.Workbooks.Open($personal)
        foreach ($c in @($wb.VBProject.VBComponents)) {
            if ($c.Name -eq "AccountingSolver") { $wb.VBProject.VBComponents.Remove($c) }
        }
        $wb.Save()
        $wb.Close($true)
    }
    $excel.Quit()
} catch { Write-Host "Could not remove the macro: $($_.Exception.Message)" -ForegroundColor Yellow }
foreach ($sec in $restore.Keys) {
    if ($null -eq $restore[$sec]) { Remove-ItemProperty -Path $sec -Name AccessVBOM -ErrorAction SilentlyContinue }
    else { Set-ItemProperty -Path $sec -Name AccessVBOM -Value $restore[$sec] -Type DWord }
}

Remove-Item (Join-Path ([Environment]::GetFolderPath("Desktop")) "Solver.lnk") -ErrorAction SilentlyContinue
Remove-Item $dest -Recurse -Force -ErrorAction SilentlyContinue
Write-Host "Solver was removed." -ForegroundColor Green
Read-Host "Press Enter to close"
