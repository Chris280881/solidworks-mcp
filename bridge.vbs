Set swApp = GetObject(, "SldWorks.Application.27")
Set doc = swApp.ActiveDoc
If Not doc Is Nothing Then
    ' Versuche alle Profile (Skizzen) im Modell zu finden
    Set feat = doc.FirstFeature
    Do While Not feat Is Nothing
        If feat.GetTypeName2 = "ProfileFeature" Then
            ' Selektiere die Skizze
            feat.Select2 False, 1
            ' Extrudiere
            Set ext = doc.FeatureManager.FeatureExtrusion2(True, False, False, 0, 0, 0.05, 0, False, False, False, False, 0, 0, False, False, False, False, 0, 0, True, True, False, False)
            If Not ext Is Nothing Then
                WScript.Echo "Success:" & ext.Name
                Exit Do
            End If
        End If
        Set feat = feat.GetNextFeature
    Loop
End If
