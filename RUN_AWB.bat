@echo off
title PCCS AWB Automation
C:\AWB_TOOLS\WAITING-SLIPS_VERIFIER

if not exist node_modules (
    echo Installing dependencies - please wait...
    npm install
    echo.
)

node awb_checker.js
pause
