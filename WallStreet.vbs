Sub WallStreet()

Dim totVol As Double
totVol = 0
Dim i As Long
Dim j As Integer
j = 0
Dim yDelta As Double
yDelta = 0
Dim pDelta As Double
Dim time As Integer
Dim dDelta As Double
Dim avDelta As Double
Dim start As Long
start = 2
' Since we don't know which row will hold the last value: https://www.excelcampus.com/vba/find-last-row-column-cell/
Dim numRows As Long
numRows = Cells(Rows.Count,1).End(xlUp).Row
Cells(1,9).Value = "Ticker"
Cells(1,10).Value = "Yearly Change"
Cells(1,11).Value = "% Change"
Cells(1,12).Value = "Tot.Stock Volume"
For i = 2 To numRows
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        totVol = totVol + Cells(i, 7).Value
        If totVol = 0 Then
            Range("I" & 2 + j).Value = Cells(i, 1).Value
            Range("J" & 2 + j).Value = 0
            Range("K" & 2 + j).Value = "%" & 0
            Range("L" & 2 + j).Value = 0
        Else
            If Cells(start, 3) = 0 Then
                For open = start To i
                    If Cells(open, 3).Value <> 0 Then
                        start = open
                        Exit For
                    End If
                 Next open
            End If
            yDelta = (Cells(i, 6) - Cells(start, 3))
            pDelta = Round((yDelta / Cells(start, 3) * 100), 2)
            start = i + 1
            Range("I" & 2 + j).Value = Cells(i, 1).Value
            Range("J" & 2 + j).Value = Round(yDelta, 2)
            Range("K" & 2 + j).Value = "%" & pDelta
            Range("L" & 2 + j).Value = totVol
             If yDelta > 0 Then
                Range("J" & 2 + j).Interior.ColorIndex = 4
           
            ElseIf yDelta < 0 Then
                Range("J" & 2 + j).Interior.ColorIndex = 3
            
            Else
                Range("J" & 2 + j).Interior.ColorIndex = 0
            End If
        End If
        totVol = 0
        yDelta = 0
        j = j + 1
        time = 0
    Else
        totVol = totVol + Cells(i, 7).Value
    End If
Next i
' Hard Section
Cells(1,13).Value = "Ticker"
Cells(1,14).Value = "Value"
Cells(2,15).Value = "Greatest % Increase"
Cells(3,15).Value = "Greatest % Decrease"
Cells(4,15).Value = "Greatest Tot.Volume"
Cells(2,17) = "%" & WorksheetFunction.Max(Range("K:K")) * 100
Cells(3,17) = "%" & WorksheetFunction.Min(Range("K:K")) * 100
Cells(4,17) = WorksheetFunction.Max(Range("L:L"))
pInc = WorksheetFunction.Match(WorksheetFunction.Max(Range("K:K")), Range("K:K"), 0)
pDec = WorksheetFunction.Match(WorksheetFunction.Min(Range("K:K")), Range("K:K"), 0)
gVol = WorksheetFunction.Match(WorksheetFunction.Max(Range("L:L")), Range("L:L"), 0)
Cells(2,16) = Cells(pInc + 1, 9)
Cells(3,16) = Cells(pDec + 1, 9)
Cells(4,16) = Cells(gVol + 1, 9)
' Challenge - I commented it out since it currently doesn't fully work. Maybe if I spent more time on it...
' Dim work as Worksheet
' Dim j as Integer
' Dim tot as Double
' For each work in Worksheets
' j = 0
' tot = 0
' work.Cells(1,9).Value = "Ticker"
' work.Cells(1,10).Value = "Tot.Stock Volume"

End Sub