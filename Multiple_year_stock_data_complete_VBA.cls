VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Module2()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim annualChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim outputRow As Long
    Dim maxPercentIncreaseTicker As String
    Dim maxPercentIncreaseValue As Double
    Dim maxPercentDecreaseTicker As String
    Dim maxPercentDecreaseValue As Double
    Dim maxTotalVolumeTicker As String
    Dim maxTotalVolumeValue As Double
    
    For Each ws In ThisWorkbook.Worksheets
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        outputRow = 2
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Volume"
        
        maxPercentIncreaseTicker = ""
        maxPercentIncreaseValue = 0
        maxPercentDecreaseTicker = ""
        maxPercentDecreaseValue = 0
        maxTotalVolumeTicker = ""
        maxTotalVolumeValue = 0
        
        For i = 2 To lastRow
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                If i = 2 Then
                    isFirstRow = True
                Else
                    isFirstRow = False
                End If
                ticker = ws.Cells(i, 1).Value
                openPrice = ws.Cells(i, 3).Value
                totalVolume = ws.Cells(i, 7).Value
            End If
            
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                closePrice = ws.Cells(i, 6).Value
                annualChange = closePrice - openPrice
                If openPrice = 0 Then
                    percentChange = 0
                Else
                    percentChange = annualChange / openPrice
                End If
                
                ws.Cells(outputRow, 9).Value = ticker
                ws.Cells(outputRow, 10).Value = annualChange
                ws.Cells(outputRow, 11).Value = percentChange
                ws.Cells(outputRow, 12).Value = totalVolume
                
                If annualChange >= 0 Then
                    ws.Cells(outputRow, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(outputRow, 10).Interior.ColorIndex = 3
                End If
                
                If percentChange > maxPercentIncreaseValue Then
                    maxPercentIncreaseValue = percentChange
                    maxPercentIncreaseTicker = ticker
                End If
                
                If percentChange < maxPercentDecreaseValue Then
                    maxPercentDecreaseValue = percentChange
                    maxPercentDecreaseTicker = ticker
                End If
                
                If totalVolume > maxTotalVolumeValue Then
                    maxTotalVolumeValue = totalVolume
                    maxTotalVolumeTicker = ticker
                End If
                
                outputRow = outputRow + 1
                ws.Cells(2, 14).Value = "Greatest % Increase"
                ws.Cells(2, 15).Value = maxPercentIncreaseTicker
                ws.Cells(2, 16).Value = maxPercentIncreaseValue
                ws.Cells(3, 14).Value = "Greatest % Decrease"
                ws.Cells(3, 15).Value = maxPercentDecreaseTicker
                ws.Cells(3, 16).Value = maxPercentDecreaseValue
                ws.Cells(4, 14).Value = "Greatest Total Volume"
            End If
        Next i
    Next ws
End Sub

