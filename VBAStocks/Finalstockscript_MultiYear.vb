
Sub AlphaTesting()

    Dim ws As Worksheet
        
     For Each ws In ThisWorkbook.Worksheets
     
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("Q2:Q3").Style = "Percent"
        
        Dim Ticker As String
        Dim Summary_Table_Row As Double
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim Volume As Double
        Dim TickerRow As Double
        Dim YearChange As Double
        'Dim MaxIncrease As Double
        Dim MaxValue As Double
        Dim MinValue As Double
        Dim GreatestVolume As Double
        Dim TestValueMax As Double
        Dim TestValueMin As Double
        Dim TestValueVol As Double
        MaxValue = -500
        MinValue = 0
        GreatestVolume = 500000000
        
        Volume = 0
        Summary_Table_Row = 2
        Ticker_Row = 2
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To LastRow

            ws.Cells(i, 11).Style = "Percent"

            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
                Volume = Volume + ws.Cells(i, 7).Value
                ClosePrice = ws.Cells(i, 6).Value
                OpenPrice = ws.Cells(i - TickerRow, 3).Value
                YearChange = ClosePrice - OpenPrice
                
                ws.Range("I" & Summary_Table_Row).Value = Ticker
                ws.Range("J" & Summary_Table_Row).Value = YearChange
                ws.Range("L" & Summary_Table_Row).Value = Volume
                
                If (OpenPrice = 0) Then
                ws.Range("K" & Summary_Table_Row).Value = "0"
                Else
                ws.Range("K" & Summary_Table_Row).Value = YearChange / OpenPrice
                End If
                
                If ws.Range("J" & Summary_Table_Row).Value >= 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.Color = vbGreen
                ElseIf ws.Range("J" & Summary_Table_Row).Value < 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.Color = vbRed
                End If
                
                Summary_Table_Row = Summary_Table_Row + 1
                Volume = 0
                TickerRow = 0
               

            Else
                Volume = Volume + ws.Cells(i, 7).Value
                TickerRow = TickerRow + 1
                
            End If
            
            Next i

            For i = 2 To LastRow
            TestValueMax = ws.Cells(i, 11).Value
            
            If TestValueMax > MaxValue Then
        
                MaxValue = TestValueMax
                ws.Range("P2").Value = ws.Cells(i, 9).Value
            
            End If

            
            TestValueMin = ws.Cells(i, 11).Value
            If TestValueMin < MinValue Then
            
                MinValue = TestValueMin
                ws.Range("P3").Value = ws.Cells(i, 9).Value

            End If

            
            TestValueVol = ws.Cells(i, 12).Value
            If TestValueVol > GreatestVolume Then
                GreatestVolume = TestValueVol
                ws.Range("P4").Value = ws.Cells(i, 9).Value

            End If

            
            ws.Range("Q2").Value = MaxValue
            ws.Range("Q3").Value = MinValue
            ws.Range("Q4").Value = GreatestVolume
           
            
            
        Next i


        
    Next ws


  
  End Sub












