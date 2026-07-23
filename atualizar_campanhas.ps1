# Atualizar Campanhas
# Rode: .\atualizar_campanhas.ps1

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$baseDir = "C:\Users\mar\OneDrive - SPADER DISTRIBUIDORA DE ALIMENTOS L\Área de Trabalho"
$camDir = "$baseDir\dashboards\campanhas"
$nodeModules = "C:\Users\mar\AppData\Local\Temp\opencode\node_modules\xlsx"

Write-Host "========================================" -ForegroundColor Yellow
Write-Host "  Atualizador de Campanhas" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

# 1. Verificar arquivos
Write-Host "[1/4] Verificando arquivos..." -ForegroundColor Cyan
$camFile = Get-ChildItem "$camDir" -Filter "*CAMPANHAS*" | Select-Object -First 1
$baseFile = "$baseDir\_bases\base_8026_2026.xlsx"

if (-not $camFile) {
    Write-Host "  ERRO: RELATORIO - CAMPANHAS.xlsx nao encontrado em $camDir" -ForegroundColor Red
    exit 1
}
Write-Host "  Planilha: $($camFile.Name)" -ForegroundColor Green

if (-not (Test-Path $baseFile)) {
    Write-Host "  ERRO: base_8026_2026.xlsx nao encontrado em $baseDir\_bases\" -ForegroundColor Red
    exit 1
}
Write-Host "  Base 8026: OK" -ForegroundColor Green

# 2. Rodar enriquecimento
Write-Host ""
Write-Host "[2/4] Enriquecendo dados..." -ForegroundColor Cyan
node "$scriptDir\enrich_campanhas.js"
if ($LASTEXITCODE -ne 0) {
    Write-Host "  ERRO ao rodar enrich_campanhas.js" -ForegroundColor Red
    exit 1
}
Write-Host "  data.json atualizado!" -ForegroundColor Green

# 3. Git commit + push
Write-Host ""
Write-Host "[3/4] Enviando para GitHub..." -ForegroundColor Cyan
Push-Location $camDir
git add data.json "RELATÓRIO - CAMPANHAS.xlsx"
$commitMsg = "feat: atualizacao manual $(Get-Date -Format 'dd/MM/yyyy HH:mm')"
git commit -m $commitMsg 2>&1 | Out-Null
$pushResult = git push 2>&1
Pop-Location

if ($LASTEXITCODE -eq 0) {
    Write-Host "  Push realizado com sucesso!" -ForegroundColor Green
} else {
    Write-Host "  Aviso: push pode ter falhado" -ForegroundColor Yellow
    Write-Host "  $pushResult" -ForegroundColor Yellow
}

# 4. Resumo
Write-Host ""
Write-Host "[4/4] Resumo" -ForegroundColor Cyan
$dados = Get-Content "$camDir\data.json" | ConvertFrom-Json
$total = $dados.rows.Count
$conv = ($dados.rows | Where-Object { $_.valor -gt 0 }).Count
$fat = ($dados.rows | Measure-Object -Property valor -Sum).Sum

Write-Host "  Total contatos: $total" -ForegroundColor White
Write-Host "  Convertidos: $conv" -ForegroundColor Green
Write-Host "  Faturamento: R$ $($fat.ToString('N2'))" -ForegroundColor Yellow
Write-Host "  Ultima atualizacao: $($dados.updatedAt)" -ForegroundColor Gray
Write-Host ""
Write-Host "Dashboard atualizado!" -ForegroundColor Green
