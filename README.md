# SolidWorks 2019 MCP Server

Ein Model Context Protocol (MCP) Server zur Automatisierung von SolidWorks 2019 via COM (pywin32).

## Features

Dieser Server ermöglicht es KI-Assistenten (wie Claude oder Gemini), SolidWorks direkt zu steuern. Zu den Funktionen gehören:

- **Status & Navigation**: Verbindung prüfen, geöffnete Dokumente listen und aktivieren.
- **Dateimanagement**: Neue Parts erstellen, Dateien öffnen, speichern und schließen.
- **Modellierung**: Erstellen von Quader-Primitiven, Skizzen (Linien, Kreise, Rechtecke) und Features (Extrusion, Schnitt, Verrundung, Fase).
- **Daten & Eigenschaften**: Auslesen von Masseneigenschaften (Masse, Volumen, Schwerpunkt), Bemaßungen und benutzerdefinierten Eigenschaften.
- **Export**: Export in Formate wie STL, STEP, IGES, DXF, DWG und PDF.
- **Zeichnungen**: Automatisierte Erstellung von Zeichnungen mit Standardansichten.

## Voraussetzungen

- **SolidWorks 2019** (oder kompatible Version, Version 27).
- **Python 3.10+**
- Erforderliche Pakete: `pywin32`, `mcp` (FastMCP).

## Installation

1. Repository klonen oder herunterladen.
2. Abhängigkeiten installieren:
   ```bash
   pip install pywin32 mcp
   ```
3. SolidWorks starten.

## Nutzung

### Als stdio-Server (Standard für Claude Desktop)
```bash
python server.py
```

### Als SSE-Server (HTTP)
```bash
python server.py --sse --port=8000
```
Oder die mitgelieferte `start_sse.bat` nutzen.

## Projektstruktur

- `server.py`: Hauptskript mit der MCP-Logik und den SolidWorks-Tools.
- `start_sse.bat`: Batch-Skript zum schnellen Starten des SSE-Servers.

## Lizenz
MIT
