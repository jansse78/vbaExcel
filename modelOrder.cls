VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "modelOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private sales As modelSales
Private itemType As String
Private salesChannel As String
Private orderPriority As String
Private orderDate As Date
Private orderId As Long
Private shipDate As Date
Private location As modelLocation
Public Property Get objSales() As modelSales
    Set objSales = sales
End Property
Public Property Set objSales(value As modelSales)
    Set sales = value
End Property
Public Property Get item_type() As String
    item_type = itemType
End Property
Public Property Let item_type(value As String)
    itemType = value
End Property
Public Property Get sales_channel() As String
    sales_channel = salesChannel
End Property
Public Property Let sales_channel(value As String)
    salesChannel = value
End Property
Public Property Get order_priority() As String
    order_priority = orderPriority
End Property
Public Property Let order_priority(value As String)
    orderPriority = value
End Property
Public Property Get order_date() As Date
    order_date = orderDate
End Property
Public Property Let order_date(value As Date)
    orderDate = value
End Property
Public Property Get order_ID() As Long
    order_ID = orderId
End Property
Public Property Let order_ID(value As Long)
    orderId = value
End Property
Public Property Get ship_date() As Date
    ship_date = shipDate
End Property
Public Property Let ship_date(value As Date)
    shipDate = value
End Property
Public Property Get objLocation() As modelLocation
    Set objLocation = location
End Property
Public Property Set objLocation(value As modelLocation)
    Set location = value
End Property
Public Property Get toString()

    toString = Me.order_ID & "," & Me.order_date & "," & Me.item_type
    

End Property

