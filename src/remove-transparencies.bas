Attribute VB_Name = "remove-transparencies"

'**********************
'ALEC HENKEN
'2025-03-13
'REMOVES ALL TRANSPARENT STATUSES FROM PARTS IN AN ASSEMBLY
'**********************

Dim swApp As Object
Dim swModel As SldWorks.ModelDoc2
Dim swAssembly As SldWorks.AssemblyDoc
Dim swComp As SldWorks.Component2
Dim swPart As SldWorks.ModelDoc2
Dim vComponents As Variant
Dim i As Integer

Sub main()
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    
    If swModel Is Nothing Then
        MsgBox "No active document found. Please open an assembly."
        Exit Sub
    End If
    
    If swModel.GetType <> swDocASSEMBLY Then
        MsgBox "The active document is not an assembly."
        Exit Sub
    End If
    
    Set swAssembly = swModel
    vComponents = swAssembly.GetComponents(False)
    
    For i = 0 To UBound(vComponents)
        Set swComp = vComponents(i)
        If Not swComp.IsSuppressed Then
            Set swPart = swComp.GetModelDoc2
            If Not swPart Is Nothing Then
                Dim transparencyState As Boolean
                transparencyState = swComp.VisibleInView2(swModel.ActiveView, swComponentVisibilityState_e.swComponentTransparency)
                
                If transparencyState Then
                    swComp.VisibleInView2 swModel.ActiveView, swComponentVisibilityState_e.swComponentOpaque
                End If
            End If
        End If
    Next i
    
    MsgBox "Transparency check and update completed."
End Sub
