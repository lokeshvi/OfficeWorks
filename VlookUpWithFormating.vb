Option Explicit

'wsParamSource = StagingWorksheet
'wsParamDestination = MasterWorksheet
Sub VLookUpWithFormating(wsParamSource As Worksheet, wsParamDestination As Worksheet, sParamSearchingRange As String, _
                        sParamToFind As String, sFoundColumn As String, sReplaceColumn As String, sToBePastedRange As String)
    
    Dim sFoundRange As String
    
    Dim range_to_copy As Range
    Dim range_to_paste As Range
    
    Dim ToFindString As String
        
    ToFindString = sParamToFind
    
    Application.CutCopyMode = True
        
    On Error Resume Next:
    sFoundRange = wsParamSource.Range(sParamSearchingRange).Find(ToFindString).Address
    
    Debug.Print sFoundRange
    
    Set range_to_copy = wsParamSource.Range(Replace(sFoundRange, sFoundColumn, sReplaceColumn))
    Set range_to_paste = wsParamDestination.Range(sToBePastedRange)
    
    range_to_copy.Copy
    range_to_paste.PasteSpecial xlPasteAllUsingSourceTheme
    
    Application.CutCopyMode = False

End Sub

