Attribute VB_Name = "Module1"
Sub challenge():

Dim gpi As Double 'Greatest Percent Increase
Dim gpd As Double 'Greatest Percent Derease
Dim gtsy As Single 'Greatest Total Stock Volume
Dim tsv As Single 'Total Stock Volume
Dim pc As Double 'Final Percent Change
Dim yci As Double 'Initial Yearly Changes
Dim ycf As Double 'Final Yearly Changes
Dim Tick As String 'Ticker Name
Dim ti As String 'Greatest Ticker Increase
Dim td As String 'Greatest Ticker Decrease
Dim Summary_Position As Integer

For Each ws In Worksheets

gpi = 0
gpd = 0
gtsy = 0
tsv = 0
yci = Cells(2, 3).Value
Summary_Position = 2


    
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row 'Countdown of rows of Whole Dataset
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
    ws.Cells(2, 14).Value = "Greatest Total Stock Volume"
    ws.Cells(3, 14).Value = "Greatest Percent Increase"
    ws.Cells(4, 14).Value = "Greatest Percent Decrease"

    For i = 2 To lastRow 'Going down ticker column
        
        If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then 'There is a change in ticker value.
            'Final Calculcations
            Tick = ws.Cells(i, 1).Value
            tsv = tsv + Cells(i, 7).Value
            ycf = ws.Cells(i, 6).Value - yci
            
            pc = ycf / yci
            
            
            'Printing of results
            ws.Range("I" & Summary_Position) = Tick
            ws.Range("J" & Summary_Position) = ycf
            ws.Range("K" & Summary_Position) = pc
            ws.Range("L" & Summary_Position) = tsv
        'Resetting Values
        Summary_Position = Summary_Position + 1
        tsv = 0
        yci = ws.Cells(i + 1, 3).Value
        Else
        
        'Totalling Stock Volume
        tsv = tsv + ws.Cells(i, 7).Value
        
        End If
    
    Next i
    'Greatest Changes Calculations
    finalRow = ws.Cells(Rows.Count, 9).End(xlUp).Row 'Countdown of rows of Calculated Dataset
    
    For j = 2 To lastRow
        If ws.Cells(j, 12) > gtsv Then
            gtsv = ws.Cells(j, 12)
            Tick = ws.Cells(j, 9)
        End If
    Next j
    For x = 2 To finalRow
        If ws.Cells(x, 10).Value > gpi Then
            gpi = ws.Cells(x, 11).Value
            ti = ws.Cells(x, 9).Value
        End If
        If ws.Cells(x, 11).Value < gpd Then
            gpd = ws.Cells(x, 11)
            td = ws.Cells(x, 9).Value
        End If
    Next x
    'Printing Highest value
    ws.Cells(2, 15).Value = Tick
    ws.Cells(3, 15).Value = ti
    ws.Cells(4, 15).Value = td
    ws.Cells(2, 16).Value = gtsv
    ws.Cells(3, 16).Value = gpi
    ws.Cells(4, 16).Value = gpd
    'Formatting cells
    For k = 2 To finalRow
        If ws.Cells(k, 10).Value < 0 Then
            ws.Cells(k, 10).Interior.Color = RGB(255, 0, 0)
        ElseIf ws.Cells(k, 10).Value > 0 Then
            ws.Cells(k, 10).Interior.Color = RGB(0, 255, 0)
        Else
            ws.Cells(k, 10).Interior.Color = xlNone
        End If
        ws.Cells(k, 11).NumberFormat = "0.00%"
    Next k
    ws.Cells(3, 16).NumberFormat = "0.00%"
    ws.Cells(4, 16).NumberFormat = "0.00%"
Next ws
    MsgBox ("Calculation complete!")
End Sub


