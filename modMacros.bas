Attribute VB_Name = "modMacros"
'note: requires reference to Microsoft Scripting Runtime
'summary variables
    Dim greatest_increase_percent As Double
    Dim greatest_increase_ticker As String
    Dim greatest_decrease_percent As Double
    Dim greatest_decrease_ticker As String
    Dim greatest_total_volume As LongLong
    Dim greatest_total_volume_ticker As String
    Dim COMMAS_IN_VOLUME As Boolean
    
Sub multiple_sheets()
    For Each ws In ActiveWorkbook.Sheets
        ws.Activate
        one_sheet
    Next ws
End Sub

Sub one_sheet()
    COMMAS_IN_VOLUME = False
    greatest_increase_percent = greatest_decrease_percent = 0
    greatest_total_volume = 0

    Dim rng As Range
    Set rng = ActiveSheet.Cells(1).CurrentRegion
    Dim dict As Dictionary
    Set dict = New Dictionary
    Dim key As Variant
    Dim obj As clsData
    
    If Not rng.Cells(1).Formula = "<ticker>" Then
        MsgBox "Please select a tab with <ticker> in the upper left corner!", vbCritical, "Warning"
        Exit Sub
    End If

    For Each cell In rng.Columns(1).Cells
        If cell.Row = 1 Then
            'pass
        Else
            key = cell.Formula
            If Not dict.Exists(key) Then
                Set dict(key) = New clsData
            End If
               
            Set obj = dict(key)
            obj.start_price = cell.Offset(0, 2).Value
            obj.end_price = cell.Offset(0, 5).Value
            obj.update_stock_volume = cell.Offset(0, 6).Value
        End If
    Next
    
    Cells(9).Formula = "Ticker"
    Cells(10).Formula = "Yearly Change"
    Cells(11).Formula = "Percent Change"
    Cells(12).Formula = "Total Stock Volume"
    
    i = 2
    For Each key In dict.Keys()
        Set obj = dict(key)
        Cells(i, 9).Formula = key
        yc = obj.yearly_change
        Cells(i, 10).Formula = yc
        If yc >= 0 Then
            Cells(i, 10).Interior.Color = RGB(0, 255, 0)
        Else
            Cells(i, 10).Interior.Color = RGB(255, 0, 0)
        End If
        'percent_change = obj.percent_change
        
        Cells(i, 11).Formula = obj.percent_change
        Cells(i, 11).NumberFormat = "0.00%"
        Cells(i, 12).Formula = obj.total_stock_volume
        update_summary_variables key, obj.percent_change, obj.total_stock_volume
        i = i + 1
    Next key
    
    Cells(16).Formula = "Ticker"
    Cells(17).Formula = "Value"
    
    Cells(2, 15).Formula = "Greatest % Increase"
    Cells(2, 16).Formula = greatest_increase_ticker
    Cells(2, 17).Formula = greatest_increase_percent
    Cells(2, 17).NumberFormat = "0.00%"
    
    Cells(3, 15).Formula = "Greatest % Decrease"
    Cells(3, 16).Formula = greatest_decrease_ticker
    Cells(3, 17).Formula = greatest_decrease_percent
    Cells(3, 17).NumberFormat = "0.00%"
    
    Cells(4, 15).Formula = "Greatest Total Volume"
    Cells(4, 16).Formula = greatest_total_volume_ticker
    Cells(4, 17).Formula = greatest_total_volume
    If COMMAS_IN_VOLUME = True Then Cells(4, 17).NumberFormat = "#,##0"
End Sub

Private Sub update_summary_variables(ByVal ticker As String, ByVal percent_change As Double, ByVal volume As Double)
    If percent_change >= 0 And percent_change > greatest_increase_percent Then
        greatest_increase_percent = percent_change
        greatest_increase_ticker = ticker
    End If
    
    If percent_change < 0 And percent_change < greatest_decrease_percent Then
        greatest_decrease_percent = percent_change
        greatest_decrease_ticker = ticker
    End If
    
    If volume > greatest_total_volume Then
        greatest_total_volume = volume
        greatest_total_volume_ticker = ticker
    End If
End Sub
