VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StreamHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private txtStream As Scripting.TextStream
Private container As DataContainer
Private delim As String
Private Sub Class_Initialize()

    delim = ","
    lastFilled = 0

End Sub
Public Sub readCsv(file_handler As fileFolderHandler, data_container As DataContainer)

    Dim lineFromCsv As String
    Dim headerLine As String
    Set container = data_container
    
    container.container_Size = Me.countCsvLines(file_handler)

    Set txtStream = file_handler.obj_fso.OpenTextFile(file_handler.obj_file.Path)
    
    headerLine = txtStream.ReadLine
    Do While Not txtStream.AtEndOfLine
        lineFromCsv = txtStream.ReadLine
        lineToContainer lineFromCsv
    Loop

    txtStream.Close
    
End Sub
Public Sub readCsvOther(file_handler As fileFolderHandler, data_container As DataContainer)

    Dim lineFromCsv As String
    Set container = data_container
    
    container.container_Size = Me.countCsvLines(file_handler)

    Open file_handler.obj_file.Path For Input As #1
    
    Line Input #1, lineFromCsv
    
    Do Until EOF(1)
        Line Input #1, lineFromCsv
        lineToContainer lineFromCsv
    Loop
    
    Close #1
    
End Sub
Public Function countCsvLines(file_handler As fileFolderHandler)

    Dim lineFromCsv As String
    Dim rowNro As Long: rowNro = 0
    
    Open file_handler.obj_file.Path For Input As #1
    
    
    Do Until EOF(1)
        Line Input #1, lineFromCsv
        rowNro = rowNro + 1
    Loop
    
    Close #1
    
    countCsvLines = rowNro - 1

End Function
Private Sub lineToContainer(lineValue As String)

    Dim arrayLine() As String
    Dim location As modelLocation
    Dim sales As modelSales
    Dim order As modelOrder
    Dim num As Variant
    
    Set location = New modelLocation
    Set sales = New modelSales
    Set order = New modelOrder
    
    arrayLine = Split(lineValue, delim)
    
    location.region = arrayLine(0)
    location.Country = arrayLine(1)
    sales.unitsSold = CDbl(arrayLine(8))
    sales.unitPrice = CDec(Replace(arrayLine(9), ".", ","))
    sales.unitCost = CDbl(Replace(arrayLine(10), ".", ","))
    sales.totalRevenue = CDbl(Replace(arrayLine(11), ".", ","))
    sales.totalCost = CDbl(Replace(arrayLine(12), ".", ","))
    sales.totalProfit = CDbl(Replace(arrayLine(13), ".", ","))
    Set order.objLocation = location
    Set order.objSales = sales
    order.item_type = arrayLine(2)
    order.sales_channel = arrayLine(3)
    order.order_priority = arrayLine(4)
    order.order_date = CDate(arrayLine(5))
    order.order_ID = Trim(arrayLine(6))
    order.ship_date = CDate(arrayLine(7))
    container.objToContainer order
    
End Sub
Public Sub writeToCsv(filePath As String, arrayToCsv() As String)
    
    Dim i As Integer
    Dim j As Integer
    Dim lineToCsv As String

    Open filePath For Output As #2
    
    For i = 0 To UBound(arrayToCsv, 1)
        For j = 0 To UBound(arrayToCsv, 2)
            lineToCsv = lineToCsv & arrayToCsv(i, j)
            If (j < UBound(arrayToCsv, 2)) Then lineToCsv = lineToCsv & ","
        Next j
        Print #2, lineToCsv
        lineToCsv = ""
    Next i
    
    Close #2

End Sub
Public Sub writeToBin(filePath As String, objArrayToBin() As modelOrder)
    
    Dim i As Integer

    obj = FreeFile

    Open filePath For Binary Lock Read Write As #obj
    
    For i = 0 To UBound(objArrayToBin)
        Put #obj, , objArrayToBin(i).toString
    Next i
    
    Close #obj

End Sub
Private Sub Class_Terminate()

    Set txtStream = Nothing

End Sub
