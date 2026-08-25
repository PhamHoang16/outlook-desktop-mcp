@echo off
REM Outlook Desktop MCP - Launcher
REM Usage:
REM   outlook-desktop-mcp.cmd mcp    Start the MCP server (stdio)
REM   outlook-desktop-mcp.cmd test       Run COM validation tests
REM   outlook-desktop-mcp.cmd test-unit  Run pytest unit tests (no Outlook)

setlocal

set PYTHON=%~dp0.venv\Scripts\python.exe

if not exist "%PYTHON%" (
    echo ERROR: Virtual environment not found. Run setup first. 1>&2
    exit /b 1
)

if "%1"=="mcp" (
    "%PYTHON%" -m outlook_desktop_mcp.server
) else if "%1"=="test" (
    "%PYTHON%" tests\phase1_com_test.py
) else if "%1"=="test-unit" (
    "%PYTHON%" -m pytest tests\unit -q
) else (
    echo Usage: outlook-desktop-mcp.cmd [mcp^|test^|test-unit] 1>&2
    exit /b 1
)
