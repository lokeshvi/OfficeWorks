Option Explicit

Dim wb As Workbook
Dim wsMacroCombining As Worksheet
Dim wsMacroStaging As Worksheet

Dim wbMasterFile As Workbook
Dim wsMasterFile As Worksheet

Dim wbLookUpFile As Workbook
Dim wsLookUpFile As Worksheet

Dim sLastRowAddressTemp As String
Dim iCountOfTotalRecordsTemp As Integer

Sub CombiningData()

    Dim iMacroCombiningMasterFileRowStart As Integer
    Dim iMacroCombiningVLookUpFileRowStart As Integer
    
    'If you are changing the Data collection starting row from 10 to else where. Change the Values, and everything changes.
    'Use the before text to find out where all you have to change manually
    iMacroCombiningMasterFileRowStart = 3
    iMacroCombiningVLookUpFileRowStart = 12

    Set wb = ThisWorkbook
    Set wsMacroCombining = wb.Worksheets("Combiner")
    Set wsMacroStaging = wb.Worksheets("Staging")
    
    Application.ScreenUpdating = True
    

    Dim sLastRowAddressInVLookUpData As String
    Dim iCountOfVLookUpFilesNeededToLooped As Integer
    
    Call iToGetTheCountOfRecords(wsMacroCombining, iMacroCombiningVLookUpFileRowStart, "B")
    
    sLastRowAddressInVLookUpData = sLastRowAddressTemp
    iCountOfVLookUpFilesNeededToLooped = iCountOfTotalRecordsTemp
    
'    Debug.Print sLastRowAddressInVLookUpData & "---------------" & sLastRowAddressTemp
'
'    Debug.Print iCountOfVLookUpFilesNeededToLooped & "---------------" & iCountOfTotalRecordsTemp
    
    Dim sStagingLookUpFileWorksheetName As String
    Dim iStagingLookUpFileHeaderRowNum As Integer
    Dim iStagingLookUpFileDataRowNum As Integer
    Dim sStagingLookUpFileFilteringColumn As Integer
    Dim sStagingLookUpFileFilteringValue As String
    Dim sStagingLookUpFileUser As String
    Dim sStagingLookUpFileLocation As String
    
    Dim sLastRowAddressInStagingLookUpFile As String
    Dim iCountOfTotalRecordsInStagingLookUpFile As Integer
    Dim iTotalRecordsInStagingWorksheet As Integer
    
    iTotalRecordsInStagingWorksheet = 1
    
    
    Dim i As Integer
    Dim iLoopLastValue As Integer
    
    iLoopLastValue = (iMacroCombiningVLookUpFileRowStart - 1) + iCountOfVLookUpFilesNeededToLooped
    
    Call DeleteCells(wsMacroStaging)
    
    'Collecting Information From the Macro
    
    'If you are changing the Data collection starting row from 10 to else where. Change the i = 11 To 11 + iCountOfVLookUpFilesNeededToLooped part
    For i = iMacroCombiningVLookUpFileRowStart To iLoopLastValue
    
        sStagingLookUpFileWorksheetName = wsMacroCombining.Cells(i, 2)
        iStagingLookUpFileHeaderRowNum = wsMacroCombining.Cells(i, 3)
        iStagingLookUpFileDataRowNum = wsMacroCombining.Cells(i, 4)
        sStagingLookUpFileFilteringColumn = wsMacroCombining.Cells(i, 5)
        sStagingLookUpFileFilteringValue = wsMacroCombining.Cells(i, 6)
        sStagingLookUpFileUser = wsMacroCombining.Cells(i, 7)
        sStagingLookUpFileLocation = wsMacroCombining.Cells(i, 8)
        
        Set wbLookUpFile = Workbooks.Open(Filename:=sStagingLookUpFileLocation, ReadOnly:=True)
        Set wsLookUpFile = wbLookUpFile.Worksheets(sStagingLookUpFileWorksheetName)
        
        On Error Resume Next
        wsLookUpFile.ShowAllData
        
        Call iToGetTheCountOfRecords(wsLookUpFile, iStagingLookUpFileDataRowNum, "A")
        
        sLastRowAddressInStagingLookUpFile = sLastRowAddressTemp

        iCountOfTotalRecordsInStagingLookUpFile = iCountOfTotalRecordsTemp
        
        wsLookUpFile.Range("A" & iStagingLookUpFileHeaderRowNum).AutoFilter Field:=sStagingLookUpFileFilteringColumn, Criteria1:=sStagingLookUpFileFilteringValue
        
        Dim iFilterRowsCount As Integer
        iFilterRowsCount = Range("A" & Rows.Count).End(xlUp).Row
        
        If (iTotalRecordsInStagingWorksheet = 1) Then
            wsLookUpFile.Range("A" & iStagingLookUpFileHeaderRowNum & ":J" & iFilterRowsCount).SpecialCells(xlCellTypeVisible).Select
        Else
            wsLookUpFile.Range("A" & iStagingLookUpFileDataRowNum & ":J" & iFilterRowsCount).SpecialCells(xlCellTypeVisible).Select
        End If
        
        Selection.Copy
        
        wsMacroStaging.Range("A" & iTotalRecordsInStagingWorksheet).PasteSpecial xlPasteAllUsingSourceTheme
        
        Application.CutCopyMode = False
        
        wbLookUpFile.Close (False)
        
        Set wbLookUpFile = Nothing
        Set wsLookUpFile = Nothing
        
        Call iToGetTheCountOfRecords(wsMacroStaging, 1, "A")
        
        iTotalRecordsInStagingWorksheet = iCountOfTotalRecordsTemp + 1
        
        'Debug.Print sStagingLookUpFileUser & "-------" & sLastRowAddressInStagingLookUpFile & "-------" & iCountOfTotalRecordsInStagingLookUpFile
    
    Next i
    
    Dim sMasterFileLocation As String
    Dim sMasterFileWorksheetName As String
    Dim iMasterFileHeaderRowNum As Integer
    Dim iMasterFileDataRowNum As Integer
    
    sMasterFileLocation = wsMacroCombining.Cells(iMacroCombiningMasterFileRowStart, 3)
    sMasterFileWorksheetName = wsMacroCombining.Cells(iMacroCombiningMasterFileRowStart + 1, 3)
    iMasterFileHeaderRowNum = wsMacroCombining.Cells(iMacroCombiningMasterFileRowStart + 2, 3)
    iMasterFileDataRowNum = wsMacroCombining.Cells(iMacroCombiningMasterFileRowStart + 3, 3)
    
    Set wbMasterFile = Workbooks.Open(Filename:=sMasterFileLocation, ReadOnly:=False)
    Set wsMasterFile = wbMasterFile.Worksheets(sMasterFileWorksheetName)
    
    Call iToGetTheCountOfRecords(wsMasterFile, iMasterFileDataRowNum, "A")
    
    Debug.Print sLastRowAddressTemp
    Debug.Print iCountOfTotalRecordsTemp
    
    For i = iMasterFileDataRowNum To iMasterFileDataRowNum + iCountOfTotalRecordsTemp
    
        Call VLookUpWithFormating(wsMacroStaging, wsMasterFile, "A:A", wsMasterFile.Cells(i, 1), "A", "H", "H" & i)
        
        Call VLookUpWithFormating(wsMacroStaging, wsMasterFile, "A:A", wsMasterFile.Cells(i, 1), "A", "I", "I" & i)
        
    Next i
    
    'Cleaning Variables
    Set wb = Nothing
    Set wsMacroCombining = Nothing
    Set wsMacroStaging = Nothing
    
    Application.DisplayAlerts = False
    wbMasterFile.Save
    wbMasterFile.Close (True)
    Application.DisplayAlerts = True
    Set wbMasterFile = Nothing
    Set wsMasterFile = Nothing

    Set wbLookUpFile = Nothing
    Set wsLookUpFile = Nothing
    
    Debug.Print "done"
    
    'Application.ScreenUpdating = True

End Sub

'This Sub will give both the last row address and also the total number of records
Sub iToGetTheCountOfRecords(wsParam As Worksheet, iParamDataStartingRowNum As Integer, sParamColumnName As String)

    sLastRowAddressTemp = ""
    iCountOfTotalRecordsTemp = 0

    sLastRowAddressTemp = wsParam.Range(sParamColumnName & iParamDataStartingRowNum & ":A1048576").Find(What:="*", lookat:=xlWhole, SearchDirection:=xlPrevious).Address
        
    iCountOfTotalRecordsTemp = wsParam.Range("$" & sParamColumnName & "$" & iParamDataStartingRowNum & ":" & sLastRowAddressTemp).Rows.Count

End Sub

Sub DeleteCells(wsParam As Worksheet)
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
    wsParam.UsedRange.Delete
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
End Sub
