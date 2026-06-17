@echo off
title Кафе Учет
cd /d C:\Users\Sair3n\Desktop\Cafe_System

:: Закрываем предыдущие запуски (если есть)
taskkill /f /im python.exe /fi "windowtitle eq streamlit*" 2>nul

echo Запуск приложения...
start /B streamlit run cafe_system_unified.py --server.port 8501 --server.headless true --browser.serverAddress localhost --browser.gatherUsageStats false

:: Ждем 4 секунды
timeout /t 4 /nobreak > nul

:: Открываем браузер только один раз
start http://localhost:8501

echo.
echo ========================================
echo   ☕ КАФЕ УЧЕТ - РАБОТАЕТ
echo ========================================
echo.
echo Приложение открыто в браузере
echo Не закрывайте это окно!
echo Для остановки нажмите Ctrl+C
echo.

pause