' ============================================================================
' lanceur-fusion-figurines.vbs
' Lance Fusion-Figurines-GUI.ps1 sans fenetre console et sans les erreurs
' classiques de PowerShell (execution policy bloquee, etc.). A garder dans
' le MEME dossier que Fusion-Figurines-GUI.ps1. Double-clique juste ce
' fichier .vbs pour demarrer l'appli.
' ============================================================================

Dim objFSO, objShell, strDossierScript, strCheminPs1

Set objFSO = CreateObject("Scripting.FileSystemObject")
strDossierScript = objFSO.GetParentFolderName(WScript.ScriptFullName)
strCheminPs1 = strDossierScript & "\Fusion-Figurines-GUI.ps1"

If Not objFSO.FileExists(strCheminPs1) Then
    MsgBox "Introuvable : " & strCheminPs1 & vbCrLf & vbCrLf & _
           "Verifie que ce .vbs est bien dans le meme dossier que Fusion-Figurines-GUI.ps1.", _
           vbCritical, "Fusion figurines - lanceur"
    WScript.Quit
End If

Set objShell = CreateObject("WScript.Shell")
objShell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & strCheminPs1 & """", 0, False
