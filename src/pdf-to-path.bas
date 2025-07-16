Attribute VB_Name = "pdf-to-path"
'**********************
'ALEC HENKEN
'2022-09-01
'GENERATES PDF FROM THE ACTIVE SESSION AND PUTS THEM IN THE PATH YOU INDICATE BELOW
'**********************

Dim swApp As SldWorks.SldWorks

Sub main()

    Set swApp = Application.SldWorks
    
    Dim swDraw As SldWorks.ModelDoc2
    
    Set swDraw = swApp.ActiveDoc
    
    If swDraw Is Nothing Then
        Err.Raise vbError, "", "Open drawing"
    End If
    
    Dim drawRevision As String
    drawRevision = swDraw.GetCustomInfoValue("", "Revision")
    
    If swDraw.GetType() = swDocumentTypes_e.swDocDRAWING Then
    
        Dim outFolder As String
        'UNCOMMENT THIS LINE TO ASK PATH EACH TIME
        outFolder = BrowseForFolder()
        'UNCOMMENT THIS LINE TO SPECIFY A SPECIFIC PATH AUTOMATICALLY EVERY TIME
        '******************TYPE YOUR PATH IN THE QUOTES BELOW*********************
        'outFolder = "C:\Users\alechenken\Documents\Drawing pdf Export"
        '*************************************************************************
        
        
        If Right(outFolder, 1) = "\" Then
            outFolder = Left(outFolder, Len(outFolder) - 1)
        End If
        
        If outFolder <> "" Then
        
            Dim outFileName As String
            outFileName = GetFileNameWithoutExtension(swDraw.GetPathName()) & " - " & drawRevision & ".pdf"
            
            Dim outFilePath As String
            outFilePath = outFolder & "\" & outFileName
            
            Dim errs As Long
            Dim warns As Long
            
            If False = swDraw.Extension.SaveAs(outFilePath, swSaveAsVersion_e.swSaveAsCurrentVersion, swSaveAsOptions_e.swSaveAsOptions_Silent, Nothing, errs, warns) Then
                Err.Raise vbError, "", "Failed to export PDF to " & outFile
            End If
            
            'UNCOMMENT THIS TO CLOSE DRAWING AFTER EXPORT
            'swApp.CloseDoc swDraw.GetTitle
            
        End If
    Else
        Err.Raise vbError, "", "Active document is not a drawing"
    End If
    
End Sub

Function GetFileNameWithoutExtension(filePath As String) As String
    GetFileNameWithoutExtension = Mid(filePath, InStrRev(filePath, "\") + 1, InStrRev(filePath, ".") - InStrRev(filePath, "\") - 1)
End Function

Function BrowseForFolder(Optional title As String = "Select Folder") As String
    
    Dim shellApp As Object
    
    Set shellApp = CreateObject("Shell.Application")
    
    Dim folder As Object
    Set folder = shellApp.BrowseForFolder(0, title, 0)
    
    If Not folder Is Nothing Then
        BrowseForFolder = folder.Self.Path
    End If
    
End Function
