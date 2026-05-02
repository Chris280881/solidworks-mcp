Set swApp = GetObject(, "SldWorks.Application.27")
Set doc = swApp.ActiveDoc
doc.ClearSelection2 True

' Versuche Ebene oben zu finden
Set feat = doc.FirstFeature
count = 0
Do While Not feat Is Nothing
    If feat.GetTypeName2 = "RefPlane" Then
        If count = 1 Then
            feat.Select2 False, 0
            Exit Do
        End If
        count = count + 1
    End If
    Set feat = feat.GetNextFeature
Loop

doc.SketchManager.InsertSketch True
doc.SketchManager.CreateCircle 0, 0, 0, 0.0025, 0, 0
' Extrusion während Skizze aktiv ist
Set cut = doc.FeatureManager.FeatureCut3(True, False, False, 1, 0, 0.01, 0.01, False, False, False, False, 0, 0, False, False, False, False, 0, 0, True, True, False, False, 0, 0, False)

If Not cut Is Nothing Then
    WScript.Echo "Success"
Else
    WScript.Echo "Failed"
End If
