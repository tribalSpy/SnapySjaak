@echo off
title SnappySjaak Shadow App
cd /d "%~dp0shadow-app"
echo Starting SnappySjaak Shadow App...
echo.
npm start
echo.
echo Shadow app stopped. If it says the port is already in use, it is probably already open at:
echo http://127.0.0.1:4174
echo.
echo The app also prints a network address like http://192.168.x.x:4174 when it starts.
echo Use that address from another PC on the same Wi-Fi/LAN.
echo.
pause
