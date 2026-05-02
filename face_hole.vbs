Set swApp = GetObject(, "SldWorks.Application.27")
Set doc = swApp.ActiveDoc
doc.ClearSelection2 True

' Top Face selektieren (Y=0.025)
' Wir nehmen (0, 0.025, 0)
ok = doc.Extension.SelectByID2("", "FACE", 0, 0.025, 0, False, 0, Nothing, 0)

If ok Then
    doc.SketchManager.InsertSketch True
    ' Da wir auf dem Face sind, ist der lokale Ursprung der Skizze im Zentrum des Face
    doc.SketchManager.CreateCircle 0, 0, 0, 0.0025, 0, 0
    ' Cut
    Set cut = doc.FeatureManager.FeatureCut3(True, False, False, 1, 0, 0.01, 0.01, False, False, False, False, 0, 0, False, False, False, False, 0, 0, True, True, False, False, 0, 0, False)
    If Not cut Is Nothing Then WScript.Echo "Success" Else WScript.Echo "Failed"
Else
    WScript.Echo "Face selection failed"
End If
