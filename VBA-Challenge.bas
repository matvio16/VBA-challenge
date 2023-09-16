Attribute VB_Name = "Module1"
Sub stocks():

Dim i As Double
Dim lastRow As Double
Dim tickerCount As Double
Dim opening As Double
Dim closing As Double
Dim newTicker As Boolean
Dim totalVol As Double
Dim percent As Double
Dim incTicker As String
Dim decTicker As String
Dim incPercent As Double
Dim decPercent As Double
Dim highVol As Double
Dim volTicker As String
Dim ws As Worksheet

Dim currentTicker As String
Dim currentOpening As Double
Dim currentClosing As Double
Dim currentVolume As Double


For Each ws In Worksheets
    
    Worksheets(ws.Name).Activate
    incPercent = 0
    decPercent = 0
    highVol = 0
    
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    tickerCount = 1

    newTicker = True

    For i = 2 To lastRow
    
        'Not necessary but included for first part of the requirements
        currentTicker = Cells(i, 1).Value
        currentOpening = Cells(i, 3).Value
        currentClosing = Cells(i, 6).Value
        currentVolume = Cells(i, 7).Value
        
        If newTicker Then
            opening = Cells(i, 3).Value
            newTicker = False
            totalVol = 0
        End If
    
        totalVol = totalVol + Cells(i, 7).Value

        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
    
            tickerCount = tickerCount + 1
            Cells(tickerCount, 9).Value = Cells(i, 1).Value
            closing = Cells(i, 6).Value
            Cells(tickerCount, 10).Value = closing - opening
            Cells(tickerCount, 10).NumberFormat = "0.00"
            If Cells(tickerCount, 10).Value > 0 Then
                Cells(tickerCount, 10).Interior.ColorIndex = 4
                Cells(tickerCount, 11).Interior.ColorIndex = 4
            ElseIf Cells(tickerCount, 10).Value < 0 Then
                Cells(tickerCount, 10).Interior.ColorIndex = 3
                Cells(tickerCount, 11).Interior.ColorIndex = 3
            End If
        
            percent = (closing - opening) / opening
            Cells(tickerCount, 11).Value = percent
            Cells(tickerCount, 11).NumberFormat = "0.00%"
        
            If percent > incPercent Then
                incPercent = percent
                incTicker = Cells(i, 1).Value
            End If
        
            If percent < decPercent Then
                decPercent = percent
                decTicker = Cells(i, 1).Value
            End If
        
            Cells(tickerCount, 12).Value = totalVol
        
            If totalVol > highVol Then
                highVol = totalVol
                volTicker = Cells(i, 1).Value
            End If
        
            newTicker = True

        End If

    Next i

    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(2, 16).Value = incTicker
    Cells(2, 17).Value = incPercent
    Cells(2, 17).NumberFormat = "0.00%"
    Cells(3, 16).Value = decTicker
    Cells(3, 17).Value = decPercent
    Cells(3, 17).NumberFormat = "0.00%"
    Cells(4, 16).Value = volTicker
    Cells(4, 17).Value = highVol


    ActiveSheet.UsedRange.EntireColumn.AutoFit

Next

End Sub
