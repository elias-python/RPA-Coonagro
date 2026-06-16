# setup.ps1 — Instalacao automatica do extrator XML
# Extraia o ZIP em qualquer lugar e execute este arquivo.
# O setup instala apenas o extrator pronto em .exe.

$ErrorActionPreference = "Stop"

Write-Host ""
Write-Host "=============================================" -ForegroundColor Green
Write-Host "  RPA Coonagro - Setup do extrator XML" -ForegroundColor Green
Write-Host "=============================================" -ForegroundColor Green
Write-Host ""

# ── Pasta onde este script esta (origem do ZIP extraido) ─────────────────────
$pastaOrigem = Split-Path -Parent $MyInvocation.MyCommand.Path
$exeOrigem = Join-Path $pastaOrigem "XML_Monitor.exe"
$assetsOrigem = Join-Path $pastaOrigem "assets"
$guiaOrigem = Join-Path $pastaOrigem "GUIA_USUARIO_EXTRATOR_XML.txt"

# ── Pasta destino padrao (OneDrive do usuario atual) ─────────────────────────
$acentuado = [char]0x00C1  # A maiusculo com acento agudo
$nomeArea  = "${acentuado}rea de Trabalho"
$pastaDestino = "$env:USERPROFILE\OneDrive - The Mosaic Company\$nomeArea\RPA Coonagro - Extrator XML"

# ── 1. Cria estrutura de pastas ───────────────────────────────────────────────
Write-Host "[1/4] Criando estrutura de pastas..." -ForegroundColor Cyan
New-Item -ItemType Directory -Path $pastaDestino -Force | Out-Null
Write-Host "      $pastaDestino" -ForegroundColor Gray
Write-Host "      OK" -ForegroundColor Green

# ── 2. Copia arquivos do ZIP para o destino ───────────────────────────────────
Write-Host ""
Write-Host "[2/4] Copiando extrator..." -ForegroundColor Cyan
if (-not (Test-Path $exeOrigem)) {
    Write-Host "      ERRO: XML_Monitor.exe nao encontrado na pasta extraida." -ForegroundColor Red
    Write-Host "      Gere o pacote novamente com empacotar.ps1." -ForegroundColor Yellow
    pause; exit 1
}
Copy-Item $exeOrigem (Join-Path $pastaDestino "XML_Monitor.exe") -Force
Write-Host "      + XML_Monitor.exe" -ForegroundColor Gray
if (Test-Path $assetsOrigem) {
    Copy-Item $assetsOrigem (Join-Path $pastaDestino "assets") -Recurse -Force
    Write-Host "      + assets" -ForegroundColor Gray
}
if (Test-Path $guiaOrigem) {
    Copy-Item $guiaOrigem (Join-Path $pastaDestino "GUIA_USUARIO_EXTRATOR_XML.txt") -Force
    Write-Host "      + GUIA_USUARIO_EXTRATOR_XML.txt" -ForegroundColor Gray
}
Write-Host "      OK" -ForegroundColor Green

# ── 3. Detecta pasta SharePoint (pasta_trabalho) ─────────────────────────────
Write-Host ""
Write-Host "[3/4] Detectando pasta SharePoint..." -ForegroundColor Cyan
$pastaTrabalho = $null
foreach ($c in @(
    "$env:USERPROFILE\The Mosaic Company\Controladoria PGA1 (Arquivos) - RPA - Coonagro",
    "$env:ONEDRIVE\Controladoria PGA1 (Arquivos) - RPA - Coonagro"
)) {
    if (Test-Path $c) { $pastaTrabalho = $c; Write-Host "      Encontrada: $c" -ForegroundColor Green; break }
}
if (-not $pastaTrabalho) {
    Write-Host "      Nao detectada automaticamente." -ForegroundColor Yellow
    Write-Host "      Informe o caminho da pasta SharePoint sincronizada:" -ForegroundColor Yellow
    Write-Host "      (Ex: C:\Users\$env:USERNAME\The Mosaic Company\Controladoria PGA1 (Arquivos) - RPA - Coonagro)" -ForegroundColor Gray
    $pastaTrabalho = Read-Host "      Caminho"
}

# ── 4. Gera config.json e arquivos auxiliares ─────────────────────────────────
Write-Host ""
Write-Host "[4/4] Gerando config.json..." -ForegroundColor Cyan
$ptEsc  = $pastaTrabalho -replace '\\', '\\'
$dataCfg = [datetime]::Now.ToString('yyyy-MM-dd HH:mm')

Set-Content -Path "$pastaDestino\config.json" -Encoding UTF8 -Value "{
  `"_gerado`": `"$dataCfg`",
    `"pasta_trabalho`": `"$ptEsc`"
}"
Write-Host "      Salvo em: $pastaDestino\config.json" -ForegroundColor Green
foreach ($arquivoAux in @("nfs_lancadas.txt", "xmls_pendentes_recarregar.txt")) {
        $destinoAux = Join-Path $pastaDestino $arquivoAux
        if (-not (Test-Path $destinoAux)) {
                Set-Content -Path $destinoAux -Encoding UTF8 -Value ""
        }
}

# ── Resumo ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "=============================================" -ForegroundColor Green
Write-Host "  Instalacao concluida!" -ForegroundColor Green
Write-Host "=============================================" -ForegroundColor Green
Write-Host ""
Write-Host "  Pasta do projeto:" -ForegroundColor Yellow
Write-Host "  $pastaDestino" -ForegroundColor White
Write-Host ""
Write-Host "  Proximos passos:" -ForegroundColor Yellow
Write-Host "  1. Leia GUIA_USUARIO_EXTRATOR_XML.txt"
Write-Host "  2. Confirme que a pasta SharePoint esta sincronizada"
Write-Host "  3. Crie no Outlook a pasta XML Coonagro"
Write-Host "  4. Abra XML_Monitor.exe"
Write-Host "  5. O XML_Monitor.exe vai rodar sem depender de Python"
Write-Host ""
pause
