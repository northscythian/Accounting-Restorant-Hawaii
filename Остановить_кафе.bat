@echo off
title Остановка приложения
echo Остановка Streamlit...
taskkill /f /im python.exe /fi "windowtitle eq streamlit*"
echo Готово. Закройте окно браузера вручную.
pause