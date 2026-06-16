# =============================================================================
# empacotar.ps1 — Gera o .exe do extrator e o ZIP pronto para outra maquina
# Execute: powershell -ExecutionPolicy Bypass -File empacotar.ps1
# =============================================================================

$ErrorActionPreference = "Stop"

$pastaProjeto = Split-Path -Parent $MyInvocation.MyCommand.Path
$releaseDir = Join-Path $pastaProjeto "release"
$pacoteDir = Join-Path $releaseDir "Extrator_XML_Monitor"
$destinoZip = Join-Path $releaseDir "RPA_Coonagro_Extrator_XML.zip"
$specPath = Join-Path $pastaProjeto "XML_Monitor.spec"
$distExe = Join-Path $pastaProjeto "dist\XML_Monitor.exe"
$assetsOrigem = Join-Path $pastaProjeto "assets"
$assetsDestino = Join-Path $pacoteDir "assets"
$guiaOrigem = Join-Path $pastaProjeto "GUIA_USUARIO_EXTRATOR_XML.txt"

function Obter-PythonExe {
    try {
        $cmd = Get-Command python -ErrorAction SilentlyContinue
        if ($cmd) {
            $versao = & $cmd.Source --version 2>&1
            if ($versao -match "Python") {
                return $cmd.Source
            }
        }
    } catch {}

    foreach ($candidato in @(
        "$env:LOCALAPPDATA\Programs\Python\Python314\python.exe",
        "$env:LOCALAPPDATA\Programs\Python\Python313\python.exe",
        "$env:LOCALAPPDATA\Programs\Python\Python312\python.exe",
        "$env:LOCALAPPDATA\Programs\Python\Python311\python.exe"
    )) {
        if (Test-Path $candidato) {
            return $candidato
        }
    }

    throw "Python nao encontrado para gerar o extrator."
}

$pythonExe = Obter-PythonExe

Write-Host ""
Write-Host "=============================================" -ForegroundColor Green
Write-Host "  RPA Coonagro - Build do extrator XML" -ForegroundColor Green
Write-Host "=============================================" -ForegroundColor Green
Write-Host ""
Write-Host "Python: $pythonExe" -ForegroundColor Gray

Write-Host ""
Write-Host "[1/4] Garantindo PyInstaller e Pillow..." -ForegroundColor Cyan
& $pythonExe -m pip install pyinstaller pillow --quiet

Write-Host ""
Write-Host "[2/4] Limpando artefatos anteriores..." -ForegroundColor Cyan
foreach ($caminho in @(
    (Join-Path $pastaProjeto "build"),
    (Join-Path $pastaProjeto "dist"),
    $pacoteDir,
    $destinoZip
)) {
    if (Test-Path $caminho) {
        Remove-Item $caminho -Recurse -Force
    }
}
New-Item -ItemType Directory -Path $releaseDir -Force | Out-Null

Write-Host ""
Write-Host "[3/4] Gerando XML_Monitor.exe..." -ForegroundColor Cyan
Push-Location $pastaProjeto
try {
    & $pythonExe -m PyInstaller --noconfirm --clean $specPath
} finally {
    Pop-Location
}

if (-not (Test-Path $distExe)) {
    throw "Falha ao gerar dist\XML_Monitor.exe."
}

Write-Host ""
Write-Host "[4/4] Montando pacote final..." -ForegroundColor Cyan
New-Item -ItemType Directory -Path $pacoteDir -Force | Out-Null
Copy-Item $distExe (Join-Path $pacoteDir "XML_Monitor.exe") -Force
Copy-Item (Join-Path $pastaProjeto "setup.ps1") (Join-Path $pacoteDir "setup.ps1") -Force
if (Test-Path $guiaOrigem) {
    Copy-Item $guiaOrigem (Join-Path $pacoteDir "GUIA_USUARIO_EXTRATOR_XML.txt") -Force
}
if (Test-Path $assetsOrigem) {
    Copy-Item $assetsOrigem $assetsDestino -Recurse -Force
}

$leiame = @(
    "=======================================================",
    "  RPA Coonagro - Extrator XML",
    "=======================================================",
    "",
    "ARQUIVOS DO PACOTE",
    "------------------",
    "- XML_Monitor.exe",
    "- setup.ps1",
    "- GUIA_USUARIO_EXTRATOR_XML.txt",
    "- assets\\ (logo e imagens de apoio)",
    "",
    "COMO INSTALAR",
    "-------------",
    "1. Extraia este ZIP em qualquer pasta.",
    "2. Execute setup.ps1.",
    "3. Leia GUIA_USUARIO_EXTRATOR_XML.txt.",
    "4. Ao final, abra XML_Monitor.exe na pasta instalada.",
    "",
    "OBSERVACOES",
    "-----------",
    "- O setup cria o config.json automaticamente.",
    "- O extrator nao precisa de Python instalado na maquina destino.",
    "- A pasta SharePoint sincronizada precisa existir no PC do usuario.",
    "- O Outlook deve ter a pasta 'XML Coonagro' criada para o monitor funcionar.",
    "======================================================="
)
$leiame | Set-Content -Path (Join-Path $pacoteDir "LEIA-ME.txt") -Encoding UTF8

Compress-Archive -Path (Join-Path $pacoteDir "*") -DestinationPath $destinoZip

Write-Host ""
Write-Host "Build concluido." -ForegroundColor Green
Write-Host "EXE: $distExe" -ForegroundColor Gray
Write-Host "ZIP: $destinoZip" -ForegroundColor Gray
Write-Host ""
