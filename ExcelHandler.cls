VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "excelHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private xlApp As Excel.Application
Private xlWb As Excel.Workbook
Private xlWs As Excel.Worksheet

Private Sub Class_Initialize()

    Set xlApp = Application

End Sub
Public Sub this_workBook()

    Set xlWb = ThisWorkbook

End Sub
Public Property Get wbPath()

    wbPath = xlWb.Path

End Property

Public Function sheetExists(sheet_name As String) As Boolean

    Dim sht As Excel.Worksheet

    On Error Resume Next

    Set sht = xlWb.Sheets(sheet_name)
       
    If (sht Is Nothing) Then
        sheetExists = False
    Else
        sheetExists = True
    End If
    
    Set sht = Nothing
    
    On Error GoTo 0
    
End Function
Public Function add_sheet(sheet_name As String) As Boolean

    If (sheetExists(sheet_name)) Then
        add_sheet = False
    Else
        Dim xlWsNew As Excel.Worksheet
        Set xlWsNew = xlWb.Sheets.Add
        xlWsNew.Name = sheet_name
    End If

End Function
Public Function delete_sheet(sheet_name As String) As Boolean

    If (sheetExists(sheet_name)) Then
        xlWb.Sheets(sheet_name).Delete
        delete_sheet = True
    Else
        delete_sheet = False
    End If

End Function
Public Sub renameSheet(shtIndex As Integer, newSheetName As String)

    xlWb.Sheets(shtIndex).Name = newSheetName

End Sub
Public Function openNewExcel(xlName As String, Optional newWsName As String)

    Set xlWb = xlApp.Workbooks.Add
    If (newWsName <> "") Then renameSheet 1, newWsName
    Set xlWs = xlWb.Worksheets(1)
        
End Function
Public Function saveWbAndClose(wbName, savePath As String)

    xlWb.SaveAs fileName:=wbName
    xlWb.Close False

End Function
Public Sub copyToExcel(arrayToCopy() As String, rows As Long, cols As Long)

    Dim rngToCopy As Range
    
    Set rngToCopy = xlWs.Range(Cells(1, 1), Cells(rows, cols))

    rngToCopy = arrayToCopy
    
    

End Sub
