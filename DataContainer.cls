VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private container() As modelOrder
Private sizeSet As Boolean
Private lastFilled As Long
Dim isSorted As Boolean
Private Sub Class_Initialize()

    lastFilled = -1
    sizeSet = False
    isSorted = False

End Sub
Public Property Get sorted() As Boolean

    sorted = isSorted

End Property
Public Property Get lastPosition() As Long

    lastPosition = lastFilled

End Property
Public Property Get fillSize()

    fillSize = lastFilled + 1

End Property
Public Property Let container_Size(value As Long)

    ReDim container(value)
    sizeSet = True
    
End Property
Public Property Get getObj(x As Long) As modelOrder

    Set getObj = container(x)

End Property
Public Function objToContainer(obj As modelOrder) As Boolean

    If (sizeSet And (lastFilled + 1) < UBound(container)) Then
        Set container(lastFilled + 1) = obj
        lastFilled = lastFilled + 1
        arrayToContainer = True
    Else
        'Err.Raise 513, "DataContainer.toContainer", "unable to insert data"
        arrayToContainer = False
    End If

End Function
Public Sub sortList()

    quickSort 0, Me.lastPosition
    isSorted = True
End Sub
Private Sub quickSort(left As Long, right As Long)

    If (left < right) Then
        Dim partitionPoint As Long: partitionPoint = quickSortPartition(left, right) 'haetaan pivotarvon oikea paikka matriisissa
        quickSort left, partitionPoint - 1 'vasemmanpuolen sorttaus
        quickSort partitionPoint + 1, right 'oikeanpuolen sorttaus
    End If

End Sub
Private Function quickSortPartition(left As Long, right As Long)

    Dim pivot As modelOrder: Set pivot = container(right)
    Dim i As Long: i = left - 1
    Dim j As Long
    
    For j = left To right
        If (j > right) Then Exit For
        If (container(j).order_ID <= pivot.order_ID) Then 'jos tutkittava j on pienempi tai yhtäiso kuin vertailuarvo (pivot)
            i = i + 1   'i pykälän eteenpäin
            swap i, j
        End If
    Next j
    
    quickSortPartition = i

End Function
Private Sub swap(i As Long, j As Long)

    Dim jOrder As modelOrder: Set jOrder = container(j)
    Set container(j) = container(i)
    Set container(i) = jOrder
    
End Sub
Public Sub toDebugPrint(Optional ammount As Long = 5)
    
    Dim i As Long
    
    If (ammount > Me.lastPosition) Then ammount = Me.lastPosition

    For i = 0 To ammount
        Debug.Print container(i).order_ID
    Next i

End Sub
Public Function throwRandomOrderIdArray(ammount As Integer)

    Dim returningArray() As Long
    Dim randNro As Long
    Dim i As Integer
    
    ReDim returningArray(ammount)
    
    For i = 0 To ammount
        randNro = ((lastFilled - 0 + 1) * Rnd + 0)
        returningArray(i) = container(randNro).order_ID
    Next i
    
    throwRandomOrderIdArray = returningArray
    

End Function
Public Function binarySearch(orderIdToSearch As Long)

    Dim middle As Long: middle = -1
    Dim begin As Long: begin = 0
    Dim last As Long: last = lastFilled
    
    Do While (begin <= last)
        middle = (begin + last) / 2
        
        If (orderIdToSearch < Trim(container(middle).order_ID)) Then
            last = middle - 1
        ElseIf (orderIdToSearch > Trim(container(middle).order_ID)) Then
            begin = middle + 1
        Else
            binarySearch = middle
            Exit Function
        End If
        
    Loop
    
    binarySearch = -1
    
End Function
Public Function arrayIdsToObjArray(orderIdArray() As Long) As modelOrder()

    Dim returningArray() As modelOrder
    Dim positionInArray As Long
    Dim fillPlace As Long: fillPlace = 0
    Dim i As Long

    ReDim returningArray(UBound(orderIdArray))

    For i = 0 To UBound(orderIdArray)
        positionInArray = Me.binarySearch(orderIdArray(i))
        If (positionInArray > -1) Then Set returningArray(fillPlace) = container(positionInArray)
        fillPlace = fillPlace + 1
    Next i
    
    fillPlace = fillPlace - 1
    
    If (UBound(returningArray) < UBound(orderIdArray)) Then ReDim Preserve returningArray(fillPlace)
    
    arrayIdsToObjArray = returningArray

End Function

Public Function convertObjectsToArray(orderArray() As modelOrder)

    Dim i As Long
    Dim e As Long
    Dim returningArray() As String
    
    On Error GoTo handler
    
    ReDim returningArray(UBound(orderArray) + 1, 13)
    
    returningArray(0, 0) = "Region"
    returningArray(0, 1) = "Country"
    returningArray(0, 2) = "Item_type"
    returningArray(0, 3) = "Sales_Channel"
    returningArray(0, 4) = "Order_priority"
    returningArray(0, 5) = "Order_date"
    returningArray(0, 6) = "Order_ID"
    returningArray(0, 7) = "Ship_date"
    returningArray(0, 8) = "Unit_sold"
    returningArray(0, 9) = "Unit_price"
    returningArray(0, 10) = "Unit_cost"
    returningArray(0, 11) = "Total_revenue"
    returningArray(0, 12) = "Total_cost"
    returningArray(0, 13) = "Total_profit"
     
    For i = 0 To UBound(orderArray)
        e = i + 1
        returningArray(e, 0) = orderArray(i).objLocation.region
        returningArray(e, 1) = orderArray(i).objLocation.Country
        returningArray(e, 2) = orderArray(i).item_type
        returningArray(e, 3) = orderArray(i).sales_channel
        returningArray(e, 4) = orderArray(i).order_priority
        returningArray(e, 5) = orderArray(i).order_date
        returningArray(e, 6) = orderArray(i).order_ID
        returningArray(e, 7) = orderArray(i).ship_date
        returningArray(e, 8) = orderArray(i).objSales.unitsSold
        returningArray(e, 9) = orderArray(i).objSales.unitPrice
        returningArray(e, 10) = orderArray(i).objSales.unitCost
        returningArray(e, 11) = orderArray(i).objSales.totalRevenue
        returningArray(e, 12) = orderArray(i).objSales.totalCost
        returningArray(e, 13) = orderArray(i).objSales.totalProfit
    Next i

handler:
    If Err.Number > 0 Then Err.Raise 513, Err.Source, "Ops, something happened that shouldnt have happened" & vbCrLf & "DataContainer.convertObjectsToArray round:" & i

    On Error GoTo 0
    convertObjectsToArray = returningArray



End Function
Private Sub Class_Terminate()
    
    lastFilled = -1
    sizeSet = False
    isSorted = False
    Erase container

End Sub
