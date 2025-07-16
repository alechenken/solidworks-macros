Attribute VB_Name = "set-all-docfonts"

'**********************
'ALEC HENKEN
'2025-07-15
'CHANGES ALL FONTS IN A DRAWING DOCUMENT TO CENTURY GOTHIC (OR YOUR CHOOSING)
'**********************

Dim swApp As Object
Dim Part As Object
Dim boolstatus As Boolean
Dim longstatus As Long, longwarnings As Long
Dim swTextFormat As SldWorks.TextFormat
Dim TextFormatObj As Object
Dim ModelDocExtension As ModelDocExtension

 

Sub main()

Set swApp = Application.SldWorks
Set Part = swApp.ActiveDoc
Set ModelDocExtension = Part.Extension

' Create array of all text format constants to update
Dim textFormats As Variant
ReDim textFormats(46)

' Populate the array with text format constants
textFormats(0) = swUserPreferenceTextFormat_e.swDetailingBalloonTextFormat
textFormats(1) = swUserPreferenceTextFormat_e.swDetailingAnnotationTextFormat
textFormats(2) = swUserPreferenceTextFormat_e.swDetailingAuxView_DelimiterTextFormat
textFormats(3) = swUserPreferenceTextFormat_e.swDetailingAuxView_LabelTextFormat
textFormats(4) = swUserPreferenceTextFormat_e.swDetailingAuxView_NameTextFormat
textFormats(5) = swUserPreferenceTextFormat_e.swDetailingAuxView_RotationTextFormat
textFormats(6) = swUserPreferenceTextFormat_e.swDetailingAuxView_ScaleTextFormat
textFormats(7) = swUserPreferenceTextFormat_e.swDetailingBendTextFormat
textFormats(8) = swUserPreferenceTextFormat_e.swDetailingBillOfMaterialTextFormat
textFormats(9) = swUserPreferenceTextFormat_e.swDetailingDatumTextFormat
textFormats(10) = swUserPreferenceTextFormat_e.swDetailingDetailLabelTextFormat
textFormats(11) = swUserPreferenceTextFormat_e.swDetailingDetailTextFormat
textFormats(12) = swUserPreferenceTextFormat_e.swDetailingDetailView_DelimiterTextFormat
textFormats(13) = swUserPreferenceTextFormat_e.swDetailingDetailView_LabelTextFormat
textFormats(14) = swUserPreferenceTextFormat_e.swDetailingDetailView_NameTextFormat
textFormats(15) = swUserPreferenceTextFormat_e.swDetailingDetailView_ScaleTextFormat
textFormats(16) = swUserPreferenceTextFormat_e.swDetailingDimensionTextFormat
textFormats(17) = swUserPreferenceTextFormat_e.swDetailingGeneralTableTextFormat
textFormats(18) = swUserPreferenceTextFormat_e.swDetailingGeometricToleranceTextFormat
textFormats(19) = swUserPreferenceTextFormat_e.swDetailingHoleTableTextFormat
textFormats(20) = swUserPreferenceTextFormat_e.swDetailingLocationLabelTextFormat
textFormats(21) = swUserPreferenceTextFormat_e.swDetailingMiscView_DelimiterTextFormat
textFormats(22) = swUserPreferenceTextFormat_e.swDetailingMiscView_NameTextFormat
textFormats(23) = swUserPreferenceTextFormat_e.swDetailingMiscView_ScaleTextFormat
textFormats(24) = swUserPreferenceTextFormat_e.swDetailingNoteTextFormat
textFormats(25) = swUserPreferenceTextFormat_e.swDetailingOrthoView_DelimiterTextFormat
textFormats(26) = swUserPreferenceTextFormat_e.swDetailingOrthoView_NameTextFormat
textFormats(27) = swUserPreferenceTextFormat_e.swDetailingOrthoView_ScaleTextFormat
textFormats(28) = swUserPreferenceTextFormat_e.swDetailingPunchTextFormat
textFormats(29) = swUserPreferenceTextFormat_e.swDetailingRevisionTableTextFormat
textFormats(30) = swUserPreferenceTextFormat_e.swDetailingSectionLabelDelimiterTextFormat
textFormats(31) = swUserPreferenceTextFormat_e.swDetailingSectionLabelLabelTextFormat
textFormats(32) = swUserPreferenceTextFormat_e.swDetailingSectionLabelNameTextFormat
textFormats(33) = swUserPreferenceTextFormat_e.swDetailingSectionLabelScaleTextFormat
textFormats(34) = swUserPreferenceTextFormat_e.swDetailingSectionLabelTextFormat
textFormats(35) = swUserPreferenceTextFormat_e.swDetailingSectionTextFormat
textFormats(36) = swUserPreferenceTextFormat_e.swDetailingSectionView_RotationTextFormat
textFormats(37) = swUserPreferenceTextFormat_e.swDetailingSurfaceFinishTextFormat
textFormats(38) = swUserPreferenceTextFormat_e.swDetailingTableTextFormat
textFormats(39) = swUserPreferenceTextFormat_e.swDetailingTitleBlockTableTextFormat
textFormats(40) = swUserPreferenceTextFormat_e.swDetailingViewArrowTextFormat
textFormats(41) = swUserPreferenceTextFormat_e.swDetailingViewTextFormat
textFormats(42) = swUserPreferenceTextFormat_e.swDetailingWeldSymbolTextFormat
textFormats(43) = swUserPreferenceTextFormat_e.swDetailingWeldSymbolTextRootInsideFont
textFormats(44) = swUserPreferenceTextFormat_e.swPointAxisCoordSystemLabelFontTextFormat
textFormats(45) = swUserPreferenceTextFormat_e.swPointAxisCoordSystemNameFontTextFormat
textFormats(46) = swUserPreferenceTextFormat_e.swSheetMetalBendNotesTextFormat

' Loop through each text format and apply the font settings. Replace Century Gothic with your preferred typeface name
Dim i As Integer
For i = 0 To UBound(textFormats)
    Set TextFormatObj = ModelDocExtension.GetUserPreferenceTextFormat(textFormats(i), 0)
    Set swTextFormat = TextFormatObj
    swTextFormat.TypeFaceName = "Century Gothic"
    boolstatus = ModelDocExtension.SetUserPreferenceTextFormat(textFormats(i), 0, swTextFormat)
Next i

' Display completion message to user
MsgBox "Switch to Century Gothic Completed. You may need to FORCE REBUILD to update kerning and appearance.", vbInformation + vbOKOnly, "Font Update"

End Sub

