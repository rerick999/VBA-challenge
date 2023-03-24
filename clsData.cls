VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'Dim mticker As String
Dim mstart_price As Double
Dim mend_price As Double
Dim mtotal_stock_volume As Double

Public Sub Class_Initialize()
    'Debug.Print "class initialized"
    mstart_price = -1
    'mend_price = 0
    'mtotal_stock_volume = 0
End Sub

'Public Property Let ticker(ByVal v As String)
'    mticker = v
'End Property

Public Property Let start_price(ByVal v As Double)
    If mstart_price = -1 Then
        'Debug.Print "setting start_price"
        mstart_price = v
    End If
    'Debug.Print mstart_price
End Property

Public Property Let end_price(ByVal v As Double)
    mend_price = v
End Property

Public Property Let update_stock_volume(ByVal v As Double)
    mtotal_stock_volume = mtotal_stock_volume + v
End Property

Public Property Get yearly_change()
    yearly_change = mend_price - mstart_price
End Property

Public Property Get percent_change()
    'percent_change = Round(yearly_change / mstart_price * 100#, 2)
    percent_change = Round(yearly_change / mstart_price, 4)
End Property

Public Property Get total_stock_volume()
    total_stock_volume = mtotal_stock_volume
End Property

    
