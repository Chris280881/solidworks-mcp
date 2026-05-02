Set swApp = GetObject(, "SldWorks.Application.27")
Set doc = swApp.ActiveDoc
Set shell = WScript.CreateObject("WScript.Shell")

' Liste der Skizzen für die Bohrungen
sketches = Array("Skizze8", "Skizze9", "Skizze10")

For Each sk In sketches
    ' 1. Skizze selektieren (Mark 1)
    ok = doc.Extension.SelectByID2(sk, "SKETCH", 0, 0, 0, False, 1, Nothing, 0)
    If ok Then
        ' 2. Cut-Extrude Dialog öffnen
        swApp.RunCommand 10, ""
        WScript.Sleep 1000 ' Warten bis Dialog da ist
        
        ' 3. "Durch alles" ist in SW 2019 oft der Standard oder per Shortcut erreichbar.
        ' Da wir T1=1 (Through All) wollen, versuchen wir ENTER.
        ' Falls "Blind" aktiv ist, bohrt er evtl nur 10mm.
        ' Wir senden "Durch alles" Shortcut falls möglich, oder TABS.
        
        ' Wir probieren: ENTER (Bestätigen)
        shell.SendKeys "{ENTER}"
        WScript.Sleep 500
    End If
Next

WScript.Echo "UI Automation completed."
