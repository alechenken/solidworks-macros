Attribute VB_Name = "remove-with-thread-callout"

'**********************
'ALEC HENKEN
'2024-01-22
'UNCHECKS THE BOX "WITH THREAD CALLOUT" ON ALL HOLE WIZARD HOLES IN A PART
'**********************

Option Explicit

Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2
Dim swModelDocExt As SldWorks.ModelDocExtension
Dim swSelMgr As SldWorks.SelectionMgr
Dim swFeatMgr As SldWorks.FeatureManager

Dim toggleCount As Integer
Dim csCount As Integer

Sub main()

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    
    If Not swModel Is Nothing Then
        If swModel.GetType <> swDocumentTypes_e.swDocPART Then
            MsgBox "Please open a part."
            End
        End If
        
        Set swModelDocExt = swModel.Extension
        Set swSelMgr = swModel.SelectionManager
        Set swFeatMgr = swModel.FeatureManager
                
        TraverseFeatures
    Else
        MsgBox "Please open the model"
    End If
    
    swModel.ClearSelection2 True
    MsgBox "Total hole feature thread callouts unchecked: " & toggleCount
    MsgBox "Total countersinks disabled: " & csCount

End Sub

Private Sub TraverseFeatures()

Dim swFeat As SldWorks.Feature
Dim swSubFeat As SldWorks.Feature

    Set swFeat = swModel.FirstFeature
    toggleCount = 0
    csCount = 0
    
    Do While Not swFeat Is Nothing
        If swFeat.GetTypeName = "HoleWzd" Then
            ToggleCosmeticThreadType swFeat
        End If
        Set swFeat = swFeat.GetNextFeature
    Loop
    
End Sub

Private Sub ToggleCosmeticThreadType(swHoleWzdFeat As SldWorks.Feature)

Dim swHoleWzdFeatData As SldWorks.WizardHoleFeatureData2
    
    Debug.Print swHoleWzdFeat.Name
    Set swHoleWzdFeatData = swHoleWzdFeat.GetDefinition

    If swHoleWzdFeatData.CosmeticThreadType = swWzdHoleCosmeticThreadTypes_e.swCosmeticThreadWithCallout Then
        swHoleWzdFeatData.CosmeticThreadType = swWzdHoleCosmeticThreadTypes_e.swCosmeticThreadWithoutCallout
        swHoleWzdFeat.ModifyDefinition swHoleWzdFeatData, swModel, Nothing
        toggleCount = toggleCount + 1
    End If
        
    If swHoleWzdFeatData.NearSideCounterSink Then
        swHoleWzdFeatData.NearSideCounterSink = False
        csCount = csCount + 1
    End If
        
    If swHoleWzdFeatData.NearSideCounterSink Then
        swHoleWzdFeatData.NearSideCounterSink = False
        csCount = csCount + 1
    End If
            
End Sub
