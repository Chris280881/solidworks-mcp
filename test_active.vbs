Set swApp = GetObject(, "SldWorks.Application.27")
Set doc = swApp.ActiveDoc
doc.Extension.SelectByID2 "Ebene oben", "PLANE", 0, 0, 0, False, 0, Nothing, 0
doc.SketchManager.InsertSketch True
doc.SketchManager.CreateCornerRectangle -0.025, -0.025, 0, 0.025, 0.025, 0
' Skizze bleibt offen!
Set ext = doc.FeatureManager.FeatureExtrusion2(True, False, False, 0, 0, 0.05, 0, False, False, False, False, 0, 0, False, False, False, False, 0, 0, True, True, False, False)
If Not ext Is Nothing Then
    WScript.Echo "Success:" & ext.Name
Else
    WScript.Echo "Failed"
End If
