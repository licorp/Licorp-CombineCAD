@echo off
title Licorp_CombineCAD - Deploy
echo ========================================
echo LICORP COMBINECAD - DEPLOY
echo ========================================
echo.
powershell -ExecutionPolicy Bypass -NoProfile -File "%~dp0deploy-bundle.ps1"
pause
