# SolidWorks 2019 MCP Server

Ein Model Context Protocol (MCP) Server zur Automatisierung von SolidWorks 2019 via COM (pywin32).

## Features

Dieser Server ermöglicht es KI-Assistenten (wie Claude oder Gemini), SolidWorks direkt zu steuern. Zu den Funktionen gehören:

- **Status & Navigation**: Verbindung prüfen, geöffnete Dokumente listen und aktivieren.
- **Dateimanagement**: Neue Parts erstellen, Dateien öffnen, speichern und schließen.
- **Modellierung**: Vollautomatische Erstellung von Geometrie via Macro-Bridge (VBScript/VBA) für maximale Kompatibilität mit SolidWorks 2019.
- **Tools**:
    - `sw_create_box`: Erstellt einen zentrierten Quader.
    - `sw_drill_hole`: Erstellt zentrierte Bohrungen auf den Hauptachsen (Vorne, Oben, Rechts) inkl. "Through All" Support.
    - `sw_chamfer_all_edges`: Wendet eine Fase auf alle Kanten des Bauteils an.
- **Lokalisierung**: Vollständige Unterstützung für deutsche SolidWorks-Installationen (Ebenennamen, Feature-Typen).

## Technische Details (SW 2019)

Für eine stabile Automatisierung nutzt dieser Server:
1. **VBScript Bridge**: Komplexe Befehle wie `FeatureExtrusion2` und `FeatureCut3` werden über temporäre VBScripts ausgeführt, um COM-Parameter-Konflikte zu vermeiden.
2. **Explizite Parameter**: Befehle wie `InsertFeatureChamfer` nutzen die exakt ermittelte Anzahl von 8 Parametern.
3. **Zentrierung**: Bohrungen werden standardmäßig als Mid-Plane-Cuts ausgeführt, um Symmetrie zu gewährleisten.

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
