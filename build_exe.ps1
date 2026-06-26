$ErrorActionPreference = "Stop"

$ProjectDir = $PSScriptRoot
$PythonVersion = "3.12.7"
$BuildRequirements = Join-Path $ProjectDir "build-requirements.txt"
$ToolsDir = Join-Path $ProjectDir ".build-tools"
$VenvDir = Join-Path $ProjectDir ".build-venv"
$DistDir = Join-Path $ProjectDir "dist"
$BuildDir = Join-Path $ProjectDir "build"

function Get-PythonArchitecture {
    param([Parameter(Mandatory = $true)][string]$PythonExe)

    $result = & $PythonExe -c "import struct; print(struct.calcsize('P') * 8)"
    if ($LASTEXITCODE -ne 0) {
        throw "Gagal menjalankan Python: $PythonExe"
    }
    return [int]$result
}

function Get-X64Python {
    $launcher = Get-Command py -ErrorAction SilentlyContinue
    if ($launcher) {
        $candidate = (& $launcher.Source -3.12 -c "import sys; print(sys.executable)" 2>$null)
        if ($LASTEXITCODE -eq 0 -and $candidate) {
            $candidate = $candidate.Trim()
            if ((Get-PythonArchitecture $candidate) -eq 64) {
                return $candidate
            }
        }
    }

    $candidate = (Get-Command python -ErrorAction SilentlyContinue).Source
    if ($candidate -and (Get-PythonArchitecture $candidate) -eq 64) {
        return $candidate
    }

    throw "Python 3.12 64-bit tidak ditemukan. Instal Python 3.12 64-bit lalu jalankan ulang."
}

function Install-X86Python {
    $installDir = Join-Path $ToolsDir "Python312-x86"
    $pythonExe = Join-Path $installDir "python.exe"
    if (Test-Path $pythonExe) {
        return $pythonExe
    }

    New-Item -ItemType Directory -Force -Path $ToolsDir | Out-Null
    $installer = Join-Path $ToolsDir "python-$PythonVersion-x86.exe"
    $downloadUrl = "https://www.python.org/ftp/python/$PythonVersion/python-$PythonVersion.exe"

    Write-Host "Mengunduh Python $PythonVersion 32-bit..."
    Invoke-WebRequest -Uri $downloadUrl -OutFile $installer

    Write-Host "Memasang Python 32-bit lokal untuk proses build..."
    $arguments = @(
        "/quiet"
        "InstallAllUsers=0"
        "TargetDir=`"$installDir`""
        "Include_launcher=0"
        "Include_test=0"
        "Include_doc=0"
        "Shortcuts=0"
        "AssociateFiles=0"
        "PrependPath=0"
    )
    $process = Start-Process -FilePath $installer -ArgumentList $arguments -Wait -PassThru -WindowStyle Hidden
    if ($process.ExitCode -ne 0 -or -not (Test-Path $pythonExe)) {
        throw "Instalasi Python 32-bit gagal (exit code $($process.ExitCode))."
    }

    return $pythonExe
}

function Initialize-BuildEnvironment {
    param(
        [Parameter(Mandatory = $true)][string]$BasePython,
        [Parameter(Mandatory = $true)][string]$Architecture
    )

    $environmentDir = Join-Path $VenvDir $Architecture
    $environmentPython = Join-Path $environmentDir "Scripts\python.exe"
    if (-not (Test-Path $environmentPython)) {
        Write-Host "Membuat build environment $Architecture..."
        & $BasePython -m venv $environmentDir
        if ($LASTEXITCODE -ne 0) {
            throw "Gagal membuat build environment $Architecture."
        }
    }

    $expectedBits = if ($Architecture -eq "x64") { 64 } else { 32 }
    $actualBits = Get-PythonArchitecture $environmentPython
    if ($actualBits -ne $expectedBits) {
        throw "Build environment $Architecture salah arsitektur ($actualBits-bit). Hapus '$environmentDir' lalu coba lagi."
    }

    Write-Host "Memastikan dependensi build $Architecture..."
    & $environmentPython -m pip install --disable-pip-version-check --upgrade -r $BuildRequirements | Out-Host
    if ($LASTEXITCODE -ne 0) {
        throw "Gagal memasang dependensi build $Architecture."
    }

    return $environmentPython
}

function Build-Application {
    param(
        [Parameter(Mandatory = $true)][string]$PythonExe,
        [Parameter(Mandatory = $true)][string]$Architecture
    )

    $applicationName = "RekapExcelBooking-$Architecture"
    $architectureBuildDir = Join-Path $BuildDir $Architecture
    $specDir = Join-Path $architectureBuildDir "spec"
    New-Item -ItemType Directory -Force -Path $specDir | Out-Null

    Write-Host "Membangun $applicationName.exe..."
    & $PythonExe -m PyInstaller `
        --noconfirm `
        --clean `
        --onefile `
        --windowed `
        --name $applicationName `
        --distpath $DistDir `
        --workpath $architectureBuildDir `
        --specpath $specDir `
        (Join-Path $ProjectDir "desktop_app.py")
    if ($LASTEXITCODE -ne 0) {
        throw "Build $Architecture gagal."
    }

    $exePath = Join-Path $DistDir "$applicationName.exe"
    $file = Get-Item $exePath
    $hash = Get-FileHash $exePath -Algorithm SHA256
    Write-Host "Selesai: $exePath"
    Write-Host "  Ukuran : $($file.Length) byte"
    Write-Host "  SHA256 : $($hash.Hash)"
}

Set-Location $ProjectDir
New-Item -ItemType Directory -Force -Path $VenvDir, $DistDir, $BuildDir | Out-Null

$legacyExe = Join-Path $DistDir "RekapExcelBooking.exe"
if (Test-Path $legacyExe) {
    Write-Host "Menghapus EXE lama tanpa label arsitektur..."
    Remove-Item -LiteralPath $legacyExe -Force
}

$x64BasePython = Get-X64Python
$x86BasePython = Install-X86Python

if ((Get-PythonArchitecture $x64BasePython) -ne 64) {
    throw "Interpreter x64 bukan Python 64-bit: $x64BasePython"
}
if ((Get-PythonArchitecture $x86BasePython) -ne 32) {
    throw "Interpreter x86 bukan Python 32-bit: $x86BasePython"
}

$x64Python = Initialize-BuildEnvironment -BasePython $x64BasePython -Architecture "x64"
$x86Python = Initialize-BuildEnvironment -BasePython $x86BasePython -Architecture "x86"

Build-Application -PythonExe $x64Python -Architecture "x64"
Build-Application -PythonExe $x86Python -Architecture "x86"

Write-Host ""
Write-Host "Dua versi aplikasi selesai dibuat di folder dist."
