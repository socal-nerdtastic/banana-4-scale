' identical to Banana4Scale.swp, just in a github readable format

Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2
Dim swAssy As SldWorks.AssemblyDoc
Dim swComp As SldWorks.Component2

Sub main()

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc

    If swModel Is Nothing Then
        MsgBox "No active document open.", vbCritical
        Exit Sub
    End If

    If swModel.GetType <> swDocASSEMBLY Then
        MsgBox "This macro only works in an Assembly document." & vbCrLf & "Please open or create an assembly first.", vbExclamation
        Exit Sub
    End If

    Set swAssy = swModel

    ' *** PLACE THE BANANA FILE IN THE MACRO FOLDER OR CHANGE THIS PATH TO YOUR BANANA PART FILE ***
    Dim bananaPath As String
    bananaPath = getBananaPath() ' <--- Edit this!
    ' I will not check if the path exists because the normal file missing dialog is enough

    ' Insert the banana component at the assembly origin (0,0,0)
    Set swComp = swAssy.AddComponent5(bananaPath, swAddComponentConfigOptions_e.swAddComponentConfigOptions_CurrentSelectedConfig, "", False, "", 0, 0, 0)

    If swComp Is Nothing Then
        MsgBox "Failed to insert the banana. Check the file path and ensure the part exists.", vbCritical
    Else
        MsgBox "Banana for scale inserted successfully!", vbInformation
    End If

    ' Rebuild and zoom to fit
    swModel.ForceRebuild3 False
    swModel.ViewZoomtofit2

End Sub

Function getBananaPath() As String
    macroPath = swApp.GetCurrentMacroPathName()
    lastSlashPos = InStrRev(macroPath, "\")
    If lastSlashPos = 0 Then lastSlashPos = InStrRev(macroPath, "/")
    folderPath = ""
    If lastSlashPos > 0 Then folderPath = Left(originalPath, lastSlashPos)
    getBananaPath = folderPath & "Banana.SLDPRT"
End Function

