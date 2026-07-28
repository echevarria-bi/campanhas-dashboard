@echo off
title Atualizar Campanhas
echo ========================================
echo   Atualizador de Campanhas
echo ========================================
echo.

set "BASEDIR=C:\Users\mar\OneDrive - SPADER DISTRIBUIDORA DE ALIMENTOS L\Area de Trabalho"
set "CAMDIR=%BASEDIR%\dashboards\campanhas"
set "SCRIPTDIR=%~dp0"

echo [1/3] Enriquecendo dados...
node "%SCRIPTDIR%enrich_campanhas.js"
if %ERRORLEVEL% neq 0 (
    echo ERRO ao enriquecer dados!
    pause
    exit /b 1
)
echo OK!
echo.

echo [2/3] Enviando para GitHub...
git add -A
git commit -m "feat: atualizacao manual %date% %time%" 2>nul
if %ERRORLEVEL% neq 0 (
    echo Nada novo para commitar.
) else (
    git push
)
echo.

echo [3/3] Resumo do data.json:
node -e "var d=require('./data.json');var c=d.rows.filter(r=>r.valor>0).length;var f=d.rows.reduce((s,r)=>s+r.valor,0);console.log('  Total contatos: '+d.rows.length);console.log('  Convertidos: '+c);console.log('  Faturamento: R\$ '+f.toFixed(2));"
echo.
echo Dashboard atualizado!
pause
