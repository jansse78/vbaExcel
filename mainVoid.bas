Attribute VB_Name = "mainVoid"
Public Sub main()

    Dim fileHandler As fileFolderHandler
    Dim stream_ As StreamHandler
    Dim container As DataContainer
    Dim newXl As excelHandler
    Dim xlMain As excelHandler
    Dim arrOrderIDs() As Long
    Dim arrOrders() As String
    
    Set fileHandler = New fileFolderHandler
    Set stream_ = New StreamHandler
    Set container = New DataContainer
    Set xlMain = New excelHandler
    xlMain.this_workBook
    container.container_Size = 17
    
    If (fileHandler.chooseFile) Then
        stream_.readCsvOther fileHandler, container
        container.sortList
    Else
        Exit Sub
    End If
                   
    fileHandler.deleteIfFound xlMain.wbPath & "\excelData.xlsx"
    fileHandler.deleteIfFound xlMain.wbPath & "\csvData.csv"
    fileHandler.deleteIfFound xlMain.wbPath & "\binData.bin"
    
    
    arrOrderIDs = container.throwRandomOrderIdArray(1000)
    arrOrders = container.convertObjectsToArray(container.arrayIdsToObjArray(arrOrderIDs))
    
    stream_.writeToCsv xlMain.wbPath & "\csvData.csv", arrOrders
    stream_.writeToBin xlMain.wbPath & "\binData.bin", container.arrayIdsToObjArray(arrOrderIDs)
    
    Set newXl = New excelHandler
    newXl.openNewExcel "datafile" & i + 1, "data"
    newXl.copyToExcel arrOrders, UBound(arrOrders), 14
    newXl.saveWbAndClose "excelData", xlMain.wbPath

        
    'container.toDebugPrint
        
    Set stream_ = Nothing
    Set fileHandler = Nothing
    Set container = Nothing

End Sub
