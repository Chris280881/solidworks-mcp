Set swApp = GetObject(, "SldWorks.Application.27")
' Schließe alle Docs um sauber zu starten
swApp.CloseAllDocuments True

' Neues Part
template = swApp.GetUserPreferenceStringValue(9)
Set doc = swApp.NewDocument(template, 0, 0, 0)

' Selektiere Ebene Vorne (Standardname in DE oft "Ebene vorne")
' Wir probieren Index-basierte Selektion falls Name scheitert
Set feat = doc.FirstFeature
Do While Not feat Is Nothing
    If feat.GetTypeName2 = "RefPlane" Then
        feat.Select2 False, 0
        Exit Do
    End If
    Set feat = feat.GetNextFeature
Loop

doc.SketchManager.InsertSketch True
' Kreis im Ursprung, r=25mm
Set seg = doc.SketchManager.CreateCircle(0, 0, 0, 0.025, 0, 0)
doc.SketchManager.InsertSketch True

' Selektiere die neue Skizze (letztes Feature)
Set feat = doc.FirstFeature
Set lastSketch = Nothing
Do While Not feat Is Nothing
    If feat.GetTypeName2 = "ProfileFeature" Then Set lastSketch = feat
    Set feat = feat.GetNextFeature
Loop

lastSketch.Select2 False, 1
' Extrudiere 50mm
Set ext = doc.FeatureManager.FeatureExtrusion2(True, False, False, 0, 0, 0.05, 0, False, False, False, False, 0, 0, False, False, False, False, 0, 0, True, True, False, False)

If Not ext Is Nothing Then
    WScript.Echo "Success:" & ext.Name
Else
    WScript.Echo "Failed"
End If
