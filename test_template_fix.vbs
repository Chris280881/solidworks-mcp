Set swApp = GetObject(, "SldWorks.Application.27")
' KEIN CloseAll, wir wollen nur ein neues Dokument
template = "C:\ProgramData\SolidWorks\SOLIDWORKS 2019\templates\Teil.prtdot"
Set doc = swApp.NewDocument(template, 0, 0, 0)

If doc Is Nothing Then
    WScript.Echo "Template failed"
    WScript.Quit
End If

' Prüfe Typ (1 = Part)
If doc.GetType <> 1 Then
    WScript.Echo "Still not a Part! Type: " & doc.GetType
    WScript.Quit
End If

' Ebene Oben selektieren
' In DE: "Ebene oben"
ok = doc.Extension.SelectByID2("Ebene oben", "PLANE", 0, 0, 0, False, 0, Nothing, 0)

doc.SketchManager.InsertSketch True
doc.SketchManager.CreateCornerRectangle -0.025, -0.025, 0, 0.025, 0.025, 0
doc.SketchManager.InsertSketch True

' Selektiere Skizze1 (Mark 1)
ok = doc.Extension.SelectByID2("Skizze1", "SKETCH", 0, 0, 0, False, 1, Nothing, 0)

' Extrudiere
Set ext = doc.FeatureManager.FeatureExtrusion2(True, False, False, 0, 0, 0.05, 0, False, False, False, False, 0, 0, False, False, False, False, 0, 0, True, True, False, False)

If Not ext Is Nothing Then
    WScript.Echo "Success:" & ext.Name
Else
    WScript.Echo "Failed"
End If
