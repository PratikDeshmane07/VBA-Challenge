Attribute VB_Name = "Module1"
Sub ticker()

    For Each ws In Worksheets

        Dim tickername As String
        Dim open_Price As Double
        Dim close_Price As Double
        Dim Yearly_Change As Double
        Dim Percent_Change As Double
        Dim Data As Integer
        Data = 2
        tickervolume = 0
        
        Dim price As Long
        price = 2
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percentage Change"
    
        Lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
        For i = 2 To Lastrow
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
                tickername = ws.Cells(i, 1).Value
            
                tickervolume = tickervolume + ws.Range("G" & i).Value
            
                ws.Range("I" & Data).Value = tickername
            
                ws.Range("L" & Data).Value = tickervolume
                
                open_Price = ws.Range("C" & price).Value
                close_Price = ws.Range("F" & i).Value
                Yearly_Change = close_Price - open_Price
                
                    If open_Price = 0 Then
                    
                        Percent_Change = 0
                        
                        Else
                            Percent_Change = Yearly_Change / open_Price
                    End If
                    
                    ws.Range("J" & Data).Value = Yearly_Change
                    ws.Range("J" & Data).NumberFormat = "$ 0.00"
                    ws.Range("K" & Data).Value = Percent_Change
                    ws.Range("K" & Data).NumberFormat = "0.00%"
                            
                            If ws.Range("J" & Data).Value > 0 Then
                            ws.Range("J" & Data).Interior.ColorIndex = 4
                        Else
                            ws.Range("J" & Data).Interior.ColorIndex = 3
                        End If
                        
                Data = Data + 1
                price = i + 1
            
                tickervolume = 0
        
            Else
                tickervolume = tickervolume + ws.Range("G" & i).Value
            
            End If

        Next i
 
        Dim Greatest_Increase As Double
        Dim Greatest_Decrease As Double
        Dim Greatest_Total As Double
        Dim Greatest_Increase_Ticker As String
        Dim Greatest_Decrease_Ticker As String
        Dim Greatest_Total_Ticker As String
        
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        Greatest_Increase = ws.Range("K2").Value
        Greates_Decrease = ws.Range("K2").Value
        Greatest_Total = ws.Range("L2").Value
        
        Ticker_Lastrow = ws.Cells(Rows.Count, "I").End(xlUp).Row
        
        For r = 2 To Ticker_Lastrow
        
            If ws.Range("K" & r + 1).Value > Greatest_Increase Then
            
                Greatest_Increase = ws.Range("K" & r + 1).Value
                Greatest_Increase_Ticker = ws.Range("I" & r + 1).Value
                
            ElseIf ws.Range("K" & r + 1).Value < Greatest_Decrease Then
            
                Greatest_Decrease = ws.Range("K" & r + 1).Value
                Greatest_Decrease_Ticker = ws.Range("I" & r + 1).Value
                
            ElseIf ws.Range("L" & r + 1).Value > Greatest_Total Then
            
                Greatest_Total = ws.Range("L" & r + 1).Value
                Greatest_Total_Ticker = ws.Range("I" & r + 1).Value
        
            End If
        Next r
        
        ws.Range("P2").Value = Greatest_Increase_Ticker
        ws.Range("P3").Value = Greatest_Decrease_Ticker
        ws.Range("P4").Value = Greatest_Total_Ticker
        ws.Range("Q2").Value = Greatest_Increase
        ws.Range("Q3").Value = Greatest_Decrease
        ws.Range("Q4").Value = Greatest_Total
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
 
    Next ws

End Sub
