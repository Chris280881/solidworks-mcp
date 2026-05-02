@echo off
echo SolidWorks MCP Server — SSE Modus
echo ===================================
echo SSE-Endpunkt: http://localhost:8000/sse
echo Zum Beenden: Strg+C
echo.
"C:\Users\Chris\AppData\Local\Python\pythoncore-3.14-64\python.exe" "%~dp0server.py" --sse --port=8000
pause
