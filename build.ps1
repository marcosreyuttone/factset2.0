# Builds dist\FS2Tools.xlam from the source files in src\.
# Run build.cmd (or: powershell -ExecutionPolicy Bypass -File build.ps1).
# Excel must be installed. Close any Excel window that has FS2Tools loaded first.

param([string]$OutFile = (Join-Path $PSScriptRoot "dist\FS2Tools.xlam"))

$ErrorActionPreference = "Stop"
$src = Join-Path $PSScriptRoot "src"
$OutFile = [IO.Path]::GetFullPath($OutFile)
New-Item -ItemType Directory -Force -Path (Split-Path $OutFile) | Out-Null
if (Test-Path $OutFile) { Remove-Item $OutFile }

function Read-Source([string]$path) {
    # The VBA editor wants CRLF line endings.
    (Get-Content -Raw -Path $path) -replace "`r?`n", "`r`n"
}

function Set-Code($component, [string]$code) {
    $cm = $component.CodeModule
    if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
    $cm.AddFromString($code)
}

# Building needs "Trust access to the VBA project object model". Turn it on for
# the duration of the build and put the previous value back afterwards.
$version = "16.0"   # Excel 2016, 2019, 2021 and 365
$curVer = Get-ItemProperty "Registry::HKEY_CLASSES_ROOT\Excel.Application\CurVer" -ErrorAction SilentlyContinue
if ($curVer) { $version = ($curVer."(default)" -replace "Excel\.Application\.", "") + ".0" }
$secKey = "HKCU:\Software\Microsoft\Office\$version\Excel\Security"
New-Item -Force -Path $secKey -ErrorAction SilentlyContinue | Out-Null
$oldVbom = (Get-ItemProperty -Path $secKey -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM
Set-ItemProperty -Path $secKey -Name AccessVBOM -Value 1 -Type DWord

$excel = $null
try {
    Write-Host "Starting Excel $version..."
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false

    $wb = $excel.Workbooks.Add()
    $project = $wb.VBProject

    foreach ($file in Get-ChildItem (Join-Path $src "modules\*.bas")) {
        Write-Host "  module  $($file.BaseName)"
        $c = $project.VBComponents.Add(1)   # standard module
        $c.Name = $file.BaseName
        Set-Code $c (Read-Source $file.FullName)
    }
    foreach ($file in Get-ChildItem (Join-Path $src "forms\*.vba")) {
        Write-Host "  form    $($file.BaseName)"
        $c = $project.VBComponents.Add(3)   # UserForm (controls are created in code)
        $c.Name = $file.BaseName
        Set-Code $c (Read-Source $file.FullName)
    }
    Write-Host "  ThisWorkbook"
    # The workbook module's name is localised (e.g. "EsteLibro"), so look it up by code name.
    Set-Code $project.VBComponents.Item($wb.CodeName) (Read-Source (Join-Path $src "ThisWorkbook.vba"))

    $wb.SaveAs($OutFile, 55)   # xlOpenXMLAddIn
    $wb.Close($false)
}
finally {
    if ($excel) {
        $excel.Quit()
        [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    if ($null -eq $oldVbom) { Remove-ItemProperty -Path $secKey -Name AccessVBOM -ErrorAction SilentlyContinue }
    else { Set-ItemProperty -Path $secKey -Name AccessVBOM -Value $oldVbom -Type DWord }
}

# Excel cannot add a custom ribbon itself, so put it into the file's zip package.
Write-Host "Adding ribbon tab..."
Start-Sleep -Seconds 1   # let Excel release the file
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
$utf8 = New-Object System.Text.UTF8Encoding $false
$zip = [IO.Compression.ZipFile]::Open($OutFile, "Update")
try {
    $ui = $zip.CreateEntry("customUI/customUI14.xml")
    $w = New-Object IO.StreamWriter($ui.Open(), $utf8)
    $w.Write([IO.File]::ReadAllText((Join-Path $src "customUI14.xml")))
    $w.Close()

    $relsEntry = $zip.GetEntry("_rels/.rels")
    $r = New-Object IO.StreamReader($relsEntry.Open())
    $rels = $r.ReadToEnd()
    $r.Close()
    $relsEntry.Delete()
    $rel = '<Relationship Id="rIdFS2Ribbon" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml"/>'
    $rels = $rels -replace "</Relationships>", "$rel</Relationships>"
    $w = New-Object IO.StreamWriter($zip.CreateEntry("_rels/.rels").Open(), $utf8)
    $w.Write($rels)
    $w.Close()
}
finally {
    $zip.Dispose()
}

Write-Host ""
Write-Host "Built $OutFile" -ForegroundColor Green

# Copy into Excel's own add-ins folder, which Excel trusts by default.
$addinsDir = Join-Path $env:APPDATA "Microsoft\AddIns"
New-Item -ItemType Directory -Force -Path $addinsDir | Out-Null
try {
    Copy-Item $OutFile $addinsDir -Force
    Write-Host "Copied to $addinsDir" -ForegroundColor Green
    Write-Host "First time only: Excel > File > Options > Add-ins > Manage: Excel Add-ins > Go... > tick FS2Tools."
}
catch {
    Write-Host "Could not copy to $addinsDir (is Excel open with FS2Tools loaded?). Close Excel and run the build again." -ForegroundColor Yellow
}
