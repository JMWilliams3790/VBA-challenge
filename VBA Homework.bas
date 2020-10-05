Attribute VB_Name = "Module1"

 Sub StockNumbers()
'Naming the columns
 'Naming Variables
 Dim addrow As Double
 Dim OpenYear As Double
 Dim CloseYear As Double
 Dim totstockvol As LongLong
 Dim GStockInc As Double
 Dim GStockDec As Double
 Dim GStockVol As LongLong
 Dim MaxIncTicket As String
 Dim MaxDecTicket As String
 Dim MaxVolTicket As String
 Dim ws_name As String
 
 LastRow = Cells(Rows.Count, 1).End(xlUp).Row
 
 
For Each ws In Worksheets
 addrow = 2
 ws_name = ws.Name
 OpenYear = 0
 CloseYear = 0
 totstockvol = 0
 ws.Range("I1").Value = "Ticker"
 ws.Range("J1").Value = "Yearly Change"
 ws.Range("K1").Value = "Percent Change"
 ws.Range("L1").Value = "Total Stock Volume"
 ws.Range("O2").Value = "Greatest % Increase"
 ws.Range("O3").Value = "Greatest % Decrease"
 ws.Range("O4").Value = "Greatest Total Volume"
 ws.Range("P1").Value = "Ticker"
 ws.Range("Q1").Value = "Value"
    
    'Begin for loop
    For i = 2 To LastRow
    
    
                
                'Search Through Column `1
                If OpenYear = 0 And ws.Cells(i + 1, 3) <> 0 Then
                    OpenYear = ws.Cells(i + 1, 3).Value
                End If
                
                
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value And ws.Cells(i, 1) <> 0 Then
                
                OpenYear = ws.Cells(i, 3)
            
            ElseIf ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value And ws.Cells(i, 1) <> 0 Then
            
                CloseYear = ws.Cells(i, 6)
                
                
            End If
            If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Or ws.Cells(i, 1).Value = ws.Cells(i - 1, 1) Then
             totstockvol = totstockvol + ws.Cells(i, 7).Value
             
            End If
            
            
           
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        
                'Write to column 9, and 10
                ws.Cells(addrow, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(addrow, 10).Value = CloseYear - OpenYear
                ws.Cells(addrow, 12).Value = totstockvol
                    
                      
                      'calculate percentage
                      
                      If ws.Cells(i, 3) <> 0 Then
                            ws.Cells(addrow, 11).Value = (CloseYear - OpenYear) / OpenYear
                            ws.Cells(addrow, 11).NumberFormat = "0.00%"
 
                      ElseIf ws.Cells(i, 3) = 0 Then
                            ws.Cells(addrow, 11) = "0"
                      End If
                      
                  'color cells
                  If ws.Cells(addrow, 10).Value > "0" Then
            
                    ws.Cells(addrow, 10).Interior.ColorIndex = 4
    
                
                  ElseIf ws.Cells(addrow, 10).Value < "0" Then
            
                    ws.Cells(addrow, 10).Interior.ColorIndex = 3
                  End If
                    'Find Greatest Percent Values
                    
           If ws.Cells(addrow, 11) > GStockInc Then
            GStockInc = ws.Cells(addrow, 11)
            MaxIncTicket = ws.Cells(addrow, 9)
            ElseIf ws.Cells(addrow, 11) < GStockDec Then
            GStockDec = ws.Cells(addrow, 11)
            MaxDecTicket = ws.Cells(addrow, 9)
            ElseIf ws.Cells(addrow, 12) > GStockVol Then
            GStockVol = ws.Cells(addrow, 12)
            MaxVolTicket = ws.Cells(addrow, 9)
            End If
            
            
            
            
            
            
            addrow = addrow + 1
            totstockvol = 0
            End If
          
                    'greatest volume
                    
                    
          
          
    Next i
ws.Range("Q2").Value = GStockInc
ws.Range("Q2").NumberFormat = "0.00%"
ws.Range("P2").Value = MaxIncTicket
ws.Range("Q3").Value = GStockDec
ws.Range("Q3").NumberFormat = "0.00%"
ws.Range("P3") = MaxDecTicket
ws.Range("Q4").Value = GStockVol
ws.Range("P4").Value = MaxVolTicket
GStockInc = 0
MaxIncTicket = ""
GStockDec = 0
GStockVol = 0
MaxDecTicket = ""
MaxVolTicket = ""


Next ws
End Sub

