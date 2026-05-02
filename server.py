"""SolidWorks 2019 MCP Server — COM Automation via pywin32
Transports:
  stdio (Standard): python server.py
  SSE   (HTTP):     python server.py --sse [--port=8000]
"""

import sys
import pythoncom
import win32com.client
from mcp.server.fastmcp import FastMCP

mcp = FastMCP("SolidWorks 2019")

SW_PROGID = "SldWorks.Application.27"  # 2019 = Version 27
_sw = None

# Null-IDispatch für COM-Methoden die LPDISPATCH Nothing erwarten (z.B. SelectByID2 Callout)
_NULL_DISPATCH = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)


# ─── INTERNE HILFSFUNKTIONEN ────────────────────────────────────────────────

def get_sw() -> win32com.client.CDispatch:
    global _sw
    pythoncom.CoInitialize()
    if _sw is None:
        try:
            _sw = win32com.client.GetActiveObject(SW_PROGID)
        except Exception:
            raise RuntimeError("SolidWorks ist nicht geöffnet. Bitte starte SolidWorks zuerst.")
    return _sw


def active_doc():
    doc = get_sw().ActiveDoc
    if doc is None:
        raise RuntimeError("Kein aktives Dokument in SolidWorks.")
    return doc


def _call(obj, method, *args):
    """Ruft eine COM-Methode auf. In Python 3.14/pywin32 sind Zero-Arg-Methoden
    Properties (nicht mehr callable), daher: mit Args normal aufrufen, ohne Args
    als Property lesen."""
    attr = getattr(obj, method)
    if args:
        return attr(*args)
    return attr if not callable(attr) else attr()


def _doc_type_from_path(path: str) -> int:
    ext = path.lower().rsplit(".", 1)[-1]
    return {"sldprt": 1, "sldasm": 2, "slddrw": 3}.get(ext, 1)


def _save_doc(doc, path: str) -> str:
    try:
        ok = doc.SaveAs(path)
        return f"Gespeichert: {path}" if ok else f"Fehler beim Speichern (SaveAs)"
    except Exception:
        errors = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
        warnings = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
        ok = doc.Extension.SaveAs3(path, 0, 0, _NULL_DISPATCH, _NULL_DISPATCH, errors, warnings)
        return (f"Gespeichert: {path}" if ok
                else f"Fehler beim Speichern (Code: {errors.value})")


def _select_plane(doc, plane: str) -> bool:
    """Wählt eine Referenzebene aus — versucht DE/EN Namen, dann Feature-Iteration."""
    candidates = {
        "top":   ["Oben", "Top Plane", "Top", "Ebene1", "Plane1"],
        "front": ["Vorne", "Front Plane", "Front", "Ebene2", "Plane2"],
        "right": ["Rechts", "Right Plane", "Right", "Ebene3", "Plane3"],
    }
    plane_idx = {"top": 0, "front": 1, "right": 2}.get(plane.lower(), 0)
    names = candidates.get(plane.lower(), [plane])
    for name in names:
        if doc.Extension.SelectByID2(name, "PLANE", 0.0, 0.0, 0.0, False, 0, _NULL_DISPATCH, 0):
            return True
    # Fallback: Modellbaum iterieren und n-te RefPlane selektieren
    try:
        feat = doc.FirstFeature()
    except Exception:
        feat = doc.FirstFeature
    count = 0
    while feat:
        try:
            type_name = feat.GetTypeName2()
        except Exception:
            type_name = feat.GetTypeName2
        if str(type_name) in ("RefPlane", "PLANE", "HistoryFolder"):
            if str(type_name) != "HistoryFolder":
                if count == plane_idx:
                    feat.Select2(False, 0)
                    return True
                count += 1
        try:
            feat = feat.GetNextFeature()
        except Exception:
            feat = feat.GetNextFeature
    return False


def _collect_dims(feat, results: list):
    while feat:
        disp = feat.GetFirstDisplayDimension
        while disp:
            dim = disp.GetDimension
            results.append(f"{dim.FullName}: {dim.SystemValue * 1000:.4f} mm")
            disp = disp.GetNext
        sub = feat.GetFirstSubFeature
        if sub:
            _collect_dims(sub, results)
        feat = feat.GetNextFeature


def _find_feature(doc, name: str):
    feat = doc.GetFirstFeature
    while feat:
        if feat.Name == name:
            return feat
        feat = feat.GetNextFeature
    return None


# ─── STATUS ─────────────────────────────────────────────────────────────────

@mcp.tool()
def sw_status() -> str:
    """Prüft ob SolidWorks läuft und gibt Version + aktives Dokument zurück."""
    sw = get_sw()
    doc = sw.ActiveDoc
    doc_info = f"Aktives Dokument: {doc.GetPathName}" if doc else "Kein Dokument geöffnet"
    return f"SolidWorks läuft. Version: {sw.RevisionNumber}\n{doc_info}"


@mcp.tool()
def sw_list_open_docs() -> str:
    """Listet alle geöffneten Dokumente auf."""
    sw = get_sw()
    docs = sw.GetDocuments
    if not docs:
        return "Keine Dokumente geöffnet."
    type_map = {1: "Part", 2: "Assembly", 3: "Drawing"}
    lines = []
    for doc in docs:
        t = type_map.get(doc.GetType, "?")
        path = doc.GetPathName or doc.GetTitle
        lines.append(f"[{t}] {path}")
    return "\n".join(lines)


@mcp.tool()
def sw_activate_doc(name_or_path: str) -> str:
    """Aktiviert ein geöffnetes Dokument anhand von Dateiname oder Pfad.

    Args:
        name_or_path: Teilpfad oder Dateiname, z.B. "würfel.sldprt"
    """
    sw = get_sw()
    docs = sw.GetDocuments
    for doc in docs:
        path = doc.GetPathName or ""
        title = doc.GetTitle or ""
        if name_or_path.lower() in path.lower() or name_or_path.lower() in title.lower():
            sw.ActivateDoc3(title, True, 0, 0)
            return f"Aktiviert: {path or title}"
    return f"Dokument '{name_or_path}' nicht gefunden."


@mcp.tool()
def sw_get_doc_info() -> str:
    """Gibt Typ, Pfad und aktive Konfiguration des aktiven Dokuments zurück."""
    doc = active_doc()
    type_map = {1: "Part", 2: "Assembly", 3: "Drawing"}
    return (
        f"Typ: {type_map.get(doc.GetType, '?')}\n"
        f"Pfad: {doc.GetPathName}\n"
        f"Konfiguration: {doc.ConfigurationManager.ActiveConfiguration.Name}"
    )


@mcp.tool()
def sw_help() -> str:
    """Listet alle verfügbaren Tools mit Kurzbeschreibung."""
    return """SolidWorks MCP Server — Verfügbare Tools:

STATUS & NAVIGATION:
  sw_status              — Verbindung + aktives Dokument
  sw_list_open_docs      — Alle geöffneten Dokumente
  sw_activate_doc        — Dokument aktivieren
  sw_get_doc_info        — Typ, Pfad, Konfiguration
  sw_help                — Diese Übersicht

DATEIEN:
  sw_new_part            — Neues leeres Part erstellen
  sw_open_file           — Datei öffnen (.sldprt/.sldasm/.slddrw)
  sw_save                — Aktives Dokument speichern
  sw_save_as             — Unter neuem Pfad speichern
  sw_close               — Aktives Dokument schließen
  sw_export              — Export: STL, STEP, IGES, DXF, DWG, PDF

SCHNELL-MODELLIERUNG:
  sw_create_box          — Quader in einem Schritt (Sketch + Extrude)

SKETCH-TOOLS:
  sw_create_sketch       — Neue Skizze auf Ebene (top/front/right)
  sw_sketch_rectangle    — Rechteck in aktiver Skizze
  sw_sketch_circle       — Kreis in aktiver Skizze
  sw_sketch_line         — Linie in aktiver Skizze
  sw_close_sketch        — Skizze schließen

FEATURES:
  sw_extrude             — Boss-Extrude (Skizze → Volumenkörper)
  sw_extrude_cut         — Cut-Extrude (Material entfernen)
  sw_fillet              — Verrundung auf ausgewählten Kanten
  sw_chamfer             — Fase auf ausgewählten Kanten
  sw_list_features       — Modellbaum anzeigen
  sw_suppress_feature    — Feature unterdrücken
  sw_unsuppress_feature  — Feature aktivieren
  sw_rename_feature      — Feature umbenennen
  sw_rebuild             — Modell neu aufbauen (Rebuild)

MATERIAL & MASSENEIGENSCHAFTEN:
  sw_set_material        — Material setzen
  sw_get_mass_properties — Masse, Volumen, Oberfläche, Schwerpunkt

BEMAΒUNGEN:
  sw_list_dimensions     — Alle Bemaßungen auflisten
  sw_set_dimension       — Bemaßungswert ändern

EIGENSCHAFTEN:
  sw_get_custom_properties   — Benutzerdefinierte Eigenschaften lesen
  sw_set_custom_property     — Eigenschaft setzen
  sw_delete_custom_property  — Eigenschaft löschen

KONFIGURATIONEN:
  sw_list_configurations     — Alle Konfigurationen
  sw_activate_configuration  — Konfiguration aktivieren

ZEICHNUNG:
  sw_new_drawing         — Neue Zeichnung mit Standard-Ansichten + Speichern
  sw_add_view            — Einzelne Ansicht zur Zeichnung hinzufügen"""


# ─── DATEIEN ────────────────────────────────────────────────────────────────

@mcp.tool()
def sw_new_part(save_path: str) -> str:
    """Erstellt ein neues leeres Part-Dokument und speichert es.

    Args:
        save_path: Vollständiger Pfad, z.B. C:\\Users\\Chris\\Desktop\\teil.sldprt
    """
    sw = get_sw()
    template = sw.GetUserPreferenceStringValue(9)  # swDefaultTemplatePart
    doc = sw.NewDocument(template, 0, 0, 0)
    if doc is None:
        return "Fehler: Part konnte nicht erstellt werden."
    doc = sw.ActiveDoc or doc
    return _save_doc(doc, save_path)


@mcp.tool()
def sw_open_file(path: str, read_only: bool = False) -> str:
    """Öffnet eine SolidWorks-Datei (.sldprt, .sldasm, .slddrw).

    Args:
        path: Vollständiger Dateipfad
        read_only: Schreibgeschützt öffnen
    """
    sw = get_sw()
    errors = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    warnings = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    doc = sw.OpenDoc6(path, _doc_type_from_path(path), 1 if read_only else 0, "", errors, warnings)
    if doc is None:
        return f"Fehler beim Öffnen: {path} (Code: {errors.value})"
    return f"Geöffnet: {path}"


@mcp.tool()
def sw_save() -> str:
    """Speichert das aktive Dokument."""
    doc = active_doc()
    errors = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    warnings = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    doc.Save3(1, errors, warnings)
    return f"Gespeichert: {doc.GetPathName}"


@mcp.tool()
def sw_save_as(path: str) -> str:
    """Speichert das aktive Dokument unter einem neuen Pfad.

    Args:
        path: Neuer Speicherpfad
    """
    return _save_doc(active_doc(), path)


@mcp.tool()
def sw_close() -> str:
    """Schließt das aktive Dokument."""
    sw = get_sw()
    doc = active_doc()
    path = doc.GetPathName
    sw.CloseDoc(path)
    return f"Geschlossen: {path}"


# ─── SCHNELL-MODELLIERUNG ───────────────────────────────────────────────────

@mcp.tool()
def sw_extrude(depth_mm: float, sketch_name: str = "Sketch1", flip: bool = False) -> str:
    """Extrudiert eine Skizze zu einem Volumenkörper (Boss-Extrude)."""
    doc = active_doc()
    
    # Sicherstellen, dass die Skizze selektiert ist (Mark 1 wird oft für Extrusion benötigt)
    ok = doc.Extension.SelectByID2(sketch_name, "SKETCH", 0, 0, 0, False, 1, _NULL_DISPATCH, 0)
    if not ok:
        # Versuch ohne Typ-Angabe
        ok = doc.Extension.SelectByID2(sketch_name, "", 0, 0, 0, False, 1, _NULL_DISPATCH, 0)
    
    if not ok:
        return f"Fehler: Skizze '{sketch_name}' konnte nicht selektiert werden."
    
    h = float(depth_mm) / 1000.0
    try:
        # Empirisch ermittelt für dieses System: FeatureExtrusion mit 20 Parametern
        # Wir füllen mit Standardwerten auf
        feat = doc.FeatureManager.FeatureExtrusion(
            True, flip, False, 
            0, 0, h, 0.0,
            False, False, False, False, 0.0, 0.0,
            False, False, False, False, True,
            False, False # Die letzten beiden Parameter (oft Tight und Poly)
        )
    except Exception as e:
        return f"Fehler bei Extrusion (Exception): {str(e)}"

    if feat is None:
        # Letzter Versuch: FeatureExtrusion2 mit 23 Parametern (falls 20 nur zufällig kein Argument-Fehler war)
        try:
            feat = doc.FeatureManager.FeatureExtrusion2(
                True, flip, False, 0, 0, h, 0.0,
                False, False, False, False, 0.0, 0.0,
                False, False, False, False, 0.0, 0.0,
                True, True, False, False
            )
        except Exception:
            pass

    if feat is None:
        return "Fehler: Extrusion fehlgeschlagen (FeatureManager gab None zurück). Prüfe ob die Skizze geschlossen und gültig ist."
    
    doc.EditRebuild3()
    return f"Extrusion '{feat.Name}' erstellt: {depth_mm} mm"


@mcp.tool()
def sw_select_feature(name: str, append: bool = False, mark: int = 0) -> str:
    """Selektiert ein Feature im Modellbaum anhand seines Namens.

    Args:
        name:   Name des Features
        append: Zu bestehender Auswahl hinzufügen (True) oder Auswahl ersetzen (False)
        mark:   Selektions-Markierung (oft 0 oder 1 für Features)
    """
    doc = active_doc()
    feat = _find_feature(doc, name)
    if feat is None:
        return f"Feature '{name}' nicht gefunden."
    ok = feat.Select2(append, mark)
    return f"'{name}' selektiert." if ok else f"Fehler beim Selektieren von '{name}'."


@mcp.tool()
def sw_create_box(width_mm: float, depth_mm: float, height_mm: float, save_path: str) -> str:
    """Erstellt ein neues Part mit einem Quader (Skizze auf Oben-Ebene + Extrusion)."""
    sw = get_sw()
    template = sw.GetUserPreferenceStringValue(9)
    doc = sw.NewDocument(template, 0, 0, 0)
    if doc is None: return "Fehler: Part-Erstellung fehlgeschlagen."
    
    doc = sw.ActiveDoc
    if not _select_plane(doc, "top"): return "Fehler: Oben-Ebene nicht gefunden."

    # Skizze öffnen
    doc.SketchManager.InsertSketch(True)
    w2, d2 = float(width_mm)/2000.0, float(depth_mm)/2000.0
    doc.SketchManager.CreateCornerRectangle(-w2, -d2, 0, w2, d2, 0)
    
    # EXTRUSION SOFORT (während Skizze noch aktiv ist)
    # Wir probieren die 21-Parameter-Variante (SW 2019 Standard)
    h = float(height_mm) / 1000.0
    feat = None
    try:
        feat = doc.FeatureManager.FeatureExtrusion2(
            True, False, False, 
            0, 0, h, 0.0,
            False, False, False, False, 0.0, 0.0,
            False, False, False, False, 0.0, 0.0,
            True, True
        )
    except Exception:
        # Fallback auf 23 Parameter
        try:
            feat = doc.FeatureManager.FeatureExtrusion2(
                True, False, False, 
                0, 0, h, 0.0,
                False, False, False, False, 0.0, 0.0,
                False, False, False, False, 0.0, 0.0,
                True, True, False, False
            )
        except Exception:
            pass

    if feat is None:
        # Letzter Versuch: Skizze schließen und dann extrudieren
        doc.SketchManager.InsertSketch(True)
        # Name ermitteln (letztes Feature)
        sketch_name = "Skizze1"
        try:
            sketch_name = doc.GetActiveSketch2.GetFeature.Name
        except:
            pass
        return sw_extrude(height_mm, sketch_name)
    
    doc.EditRebuild3()
    return _save_doc(doc, save_path)


@mcp.tool()
def sw_extrude_cut(depth_mm: float, sketch_name: str = "", through_all: bool = False) -> str:
    """Schneidet eine Extrusion in das Modell (Cut-Extrude).

    Args:
        depth_mm:    Schnitttiefe in mm
        sketch_name: Skizzenname (leer = zuletzt aktive)
        through_all: Durch alles schneiden
    """
    doc = active_doc()
    if sketch_name:
        doc.Extension.SelectByID2(sketch_name, "SKETCH", 0, 0, 0, False, 0, _NULL_DISPATCH, 0)
    t1 = 1 if through_all else 0
    d1 = 0.0 if through_all else depth_mm / 1000.0
    feat = doc.FeatureManager.FeatureCut3(
        True, False, False,
        t1, 0,
        d1, 0.0,
        False, False,
        0.0, 0.0,
        False, False,
        False, False,
        False, False,
        0.0, 0.0,
        True, True, True,
        False, False,
        0, 0.0, False, False
    )
    if feat is None:
        return "Fehler: Cut-Extrude fehlgeschlagen."
    doc.EditRebuild3
    label = "durch alles" if through_all else f"{depth_mm} mm"
    return f"Cut-Extrude '{feat.Name}': {label}"


@mcp.tool()
def sw_fillet(radius_mm: float) -> str:
    """Fügt eine Verrundung auf vorher selektierten Kanten hinzu.
    Kanten müssen vor diesem Aufruf mit sw_select_entity gewählt sein.

    Args:
        radius_mm: Verrundungsradius in mm
    """
    doc = active_doc()
    feat = doc.FeatureManager.FeatureFillet(
        0,                   # Feature-Typ (0 = konstanter Radius)
        radius_mm / 1000.0,  # Radius
        False, False, False, # Tangentiell, Vollrund, Rundkante
        0, 0, 0              # Setback-Werte
    )
    if feat is None:
        return "Fehler: Verrundung fehlgeschlagen. Kanten zuerst auswählen."
    doc.EditRebuild3
    return f"Verrundung '{feat.Name}': r={radius_mm} mm"


@mcp.tool()
def sw_chamfer(distance_mm: float, angle_deg: float = 45.0) -> str:
    """Fügt eine Fase auf vorher selektierten Kanten hinzu.

    Args:
        distance_mm: Fasenabstand in mm
        angle_deg:   Fasenwinkel in Grad (Standard: 45°)
    """
    doc = active_doc()
    import math
    feat = doc.FeatureManager.InsertFeatureChamfer(
        0,                       # Typ (0 = Winkel-Abstand)
        False,                   # Vollrund
        distance_mm / 1000.0,
        math.radians(angle_deg),
        0.0,
        False, False
    )
    if feat is None:
        return "Fehler: Fase fehlgeschlagen. Kanten zuerst auswählen."
    doc.EditRebuild3
    return f"Fase '{feat.Name}': {distance_mm} mm @ {angle_deg}°"


@mcp.tool()
def sw_list_features() -> str:
    """Listet alle Features im Modellbaum des aktiven Dokuments auf."""
    doc = active_doc()
    lines = []
    feat = doc.GetFirstFeature
    while feat:
        lines.append(f"{feat.Name}  [{feat.GetTypeName2}]")
        feat = feat.GetNextFeature
    return "\n".join(lines) if lines else "Keine Features."


@mcp.tool()
def sw_suppress_feature(name: str) -> str:
    """Unterdrückt ein Feature im Modellbaum.

    Args:
        name: Feature-Name, z.B. "Boss-Extrude1"
    """
    doc = active_doc()
    feat = _find_feature(doc, name)
    if feat is None:
        return f"Feature '{name}' nicht gefunden."
    feat.Select2(False, -1)
    doc.EditSuppress2()
    return f"'{name}' unterdrückt."


@mcp.tool()
def sw_unsuppress_feature(name: str) -> str:
    """Aktiviert ein unterdrücktes Feature.

    Args:
        name: Feature-Name
    """
    doc = active_doc()
    feat = _find_feature(doc, name)
    if feat is None:
        return f"Feature '{name}' nicht gefunden."
    feat.Select2(False, -1)
    doc.EditUnsuppress2()
    return f"'{name}' aktiviert."


@mcp.tool()
def sw_rename_feature(old_name: str, new_name: str) -> str:
    """Benennt ein Feature im Modellbaum um.

    Args:
        old_name: Aktueller Name
        new_name: Neuer Name
    """
    doc = active_doc()
    feat = _find_feature(doc, old_name)
    if feat is None:
        return f"Feature '{old_name}' nicht gefunden."
    feat.Name = new_name
    return f"'{old_name}' → '{new_name}'"


@mcp.tool()
def sw_rebuild() -> str:
    """Baut das aktive Modell neu auf (Rebuild / Strg+B)."""
    active_doc().EditRebuild3()
    return "Rebuild abgeschlossen."


# ─── MATERIAL & MASSENEIGENSCHAFTEN ─────────────────────────────────────────

@mcp.tool()
def sw_set_material(material_name: str, library: str = "SolidWorks Materials") -> str:
    """Setzt das Material des aktiven Parts.

    Häufige Materialien:
      Stahl:      "Plain Carbon Steel", "Alloy Steel", "AISI 304"
      Aluminium:  "1060 Alloy", "6061 Alloy"
      Kunststoff: "ABS", "Nylon 6/10"

    Args:
        material_name: Materialname aus der SolidWorks-Datenbank
        library:       Bibliotheksname (Standard: "SolidWorks Materials")
    """
    doc = active_doc()
    try:
        doc.SetMaterialPropertyName2("", library, material_name)
        return f"Material gesetzt: {material_name} ({library})"
    except Exception as e:
        return f"Fehler: {e}\nTipp: Materialname genau prüfen (Groß-/Kleinschreibung)."


@mcp.tool()
def sw_get_mass_properties() -> str:
    """Gibt Masse, Volumen, Oberfläche und Schwerpunkt des aktiven Parts zurück."""
    doc = active_doc()
    mass_prop = doc.Extension.CreateMassProperty
    mass_prop.UseSystemUnits = True
    mass_prop.AddBodies(doc.GetBodies2(0, True))
    cx, cy, cz = mass_prop.CenterOfMass
    return (
        f"Masse:       {mass_prop.Mass * 1000:.4f} g\n"
        f"Volumen:     {mass_prop.Volume * 1e6:.4f} cm³\n"
        f"Oberfläche:  {mass_prop.SurfaceArea * 1e4:.4f} cm²\n"
        f"Schwerpunkt: ({cx*1000:.3f}, {cy*1000:.3f}, {cz*1000:.3f}) mm"
    )


# ─── BEMAΒUNGEN ─────────────────────────────────────────────────────────────

@mcp.tool()
def sw_list_dimensions() -> str:
    """Listet alle Bemaßungen im aktiven Dokument auf (Name + Wert)."""
    doc = active_doc()
    results = []
    _collect_dims(doc.GetFirstFeature(), results)
    return "\n".join(results) if results else "Keine Bemaßungen gefunden."


@mcp.tool()
def sw_set_dimension(name: str, value: float) -> str:
    """Ändert den Wert einer Bemaßung (in mm oder Grad).

    Args:
        name:  Bemaßungsname, z.B. "D1@Sketch1" oder "D1@Boss-Extrude1"
        value: Neuer Wert in mm
    """
    doc = active_doc()
    dim = doc.Parameter(name)
    if dim is None:
        return f"Bemaßung '{name}' nicht gefunden."
    dim.SystemValue = value / 1000.0
    doc.EditRebuild3
    return f"'{name}' → {value} mm"


# ─── CUSTOM PROPERTIES ──────────────────────────────────────────────────────

@mcp.tool()
def sw_get_custom_properties() -> str:
    """Liest alle benutzerdefinierten Eigenschaften des aktiven Dokuments."""
    doc = active_doc()
    mgr = doc.Extension.CustomPropertyManager("")
    names = mgr.GetNames
    if not names:
        return "Keine benutzerdefinierten Eigenschaften vorhanden."
    lines = []
    for name in names:
        val_out = ""
        resolved = ""
        mgr.Get5(name, False, val_out, resolved, False)
        lines.append(f"{name}: {resolved}")
    return "\n".join(lines)


@mcp.tool()
def sw_set_custom_property(name: str, value: str) -> str:
    """Setzt eine benutzerdefinierte Eigenschaft.

    Args:
        name:  Eigenschaftsname, z.B. "Bauteilnummer", "Werkstoff"
        value: Wert als Text
    """
    doc = active_doc()
    mgr = doc.Extension.CustomPropertyManager("")
    mgr.Add3(name, 30, value, 1)  # 30 = swCustomInfoText
    return f"'{name}' = '{value}'"


@mcp.tool()
def sw_delete_custom_property(name: str) -> str:
    """Löscht eine benutzerdefinierte Eigenschaft.

    Args:
        name: Eigenschaftsname
    """
    doc = active_doc()
    mgr = doc.Extension.CustomPropertyManager("")
    ok = mgr.Delete2(name)
    return f"'{name}' gelöscht." if ok else f"'{name}' nicht gefunden."


# ─── KONFIGURATIONEN ────────────────────────────────────────────────────────

@mcp.tool()
def sw_list_configurations() -> str:
    """Listet alle Konfigurationen des aktiven Dokuments auf."""
    doc = active_doc()
    names = doc.GetConfigurationNames
    active = doc.ConfigurationManager.ActiveConfiguration.Name
    return "\n".join(f"{'* ' if n == active else '  '}{n}" for n in names)


@mcp.tool()
def sw_activate_configuration(name: str) -> str:
    """Aktiviert eine Konfiguration.

    Args:
        name: Konfigurationsname
    """
    doc = active_doc()
    ok = doc.ShowConfiguration2(name)
    return f"Konfiguration '{name}' aktiviert." if ok else f"'{name}' nicht gefunden."


# ─── EXPORT ─────────────────────────────────────────────────────────────────

@mcp.tool()
def sw_export(output_path: str) -> str:
    """Exportiert das aktive Dokument. Format wird aus der Dateiendung ermittelt.

    Unterstützte Formate: .stl .step .stp .iges .igs .dxf .dwg .pdf .edrw

    Args:
        output_path: Ausgabepfad, z.B. C:\\Desktop\\teil.stl
    """
    doc = active_doc()
    errors = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    warnings = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    ok = doc.Extension.SaveAs3(output_path, 0, 0, None, None, errors, warnings)
    return f"Exportiert: {output_path}" if ok else f"Fehler (E:{errors.value}, W:{warnings.value})"


# ─── ZEICHNUNG ──────────────────────────────────────────────────────────────

@mcp.tool()
def sw_new_drawing(part_path: str, save_path: str, sheet_size: str = "A4L") -> str:
    """Erstellt eine neue Zeichnung mit Standard-Ansichten (Vorne, Oben, Rechts, Isometrie).

    Args:
        part_path:  Pfad zum Part (.sldprt)
        save_path:  Speicherpfad (.slddrw)
        sheet_size: "A4L" (quer), "A4P" (hoch), "A3L", "A3P" (Standard: A4L)
    """
    sw = get_sw()
    template = sw.GetUserPreferenceStringValue(14)  # swDefaultTemplateDrawing

    sizes = {
        "A4L": (0.2970, 0.2100),
        "A4P": (0.2100, 0.2970),
        "A3L": (0.4200, 0.2970),
        "A3P": (0.2970, 0.4200),
        "A2L": (0.5940, 0.4200),
    }
    w, h = sizes.get(sheet_size.upper(), (0.2970, 0.2100))

    doc = sw.NewDocument(template, 0, w, h)
    if doc is None:
        return "Fehler: Zeichnung konnte nicht erstellt werden."

    draw = win32com.client.CastTo(doc, "IDrawingDoc")

    # Ansichten: (Name, X, Y) — Positionen in Metern auf dem Blatt
    views = [
        ("*Front",     0.07,  h * 0.40),
        ("*Top",       0.07,  h * 0.65),
        ("*Right",     0.17,  h * 0.40),
        ("*Isometric", 0.20,  h * 0.65),
    ]

    added = []
    for name, x, y in views:
        try:
            view = draw.CreateDrawViewFromModelView3(part_path, name, x, y, 0)
            if view:
                added.append(name.lstrip("*"))
        except Exception:
            pass

    doc.EditRebuild3
    result = _save_doc(doc, save_path)
    return f"Zeichnung erstellt. Ansichten: {', '.join(added) or 'keine'}\n{result}"


@mcp.tool()
def sw_add_view(model_path: str, view_name: str, x_mm: float, y_mm: float) -> str:
    """Fügt eine Modellansicht in die aktive Zeichnung ein.

    Args:
        model_path: Pfad zum Part/Assembly (.sldprt/.sldasm)
        view_name:  "*Front", "*Top", "*Right", "*Isometric", "*Bottom", "*Left", "*Back"
        x_mm:       X-Position auf dem Blatt in mm
        y_mm:       Y-Position auf dem Blatt in mm
    """
    doc = active_doc()
    draw = win32com.client.CastTo(doc, "IDrawingDoc")
    view = draw.CreateDrawViewFromModelView3(model_path, view_name, x_mm / 1000.0, y_mm / 1000.0, 0)
    if view is None:
        return f"Fehler: Ansicht '{view_name}' konnte nicht hinzugefügt werden."
    return f"Ansicht '{view_name}' bei ({x_mm}, {y_mm}) mm eingefügt."


# ─── ENTRY POINT ────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if "--sse" in sys.argv:
        port = 8000
        for arg in sys.argv:
            if arg.startswith("--port="):
                port = int(arg.split("=")[1])
        print(f"SolidWorks MCP Server (SSE) läuft auf http://0.0.0.0:{port}")
        print(f"SSE-Endpunkt:  http://localhost:{port}/sse")
        print(f"Zum Beenden:   Strg+C")
        mcp.run(transport="sse", host="0.0.0.0", port=port)
    else:
        mcp.run(transport="stdio")
