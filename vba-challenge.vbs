VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stock()

    'To loop throught the workbook
    For Each ws In Worksheets
    
    'Set an intial variable for holding the ticket
    Dim ticker As String
    
    'Set an intial varibale for holding the total stock volume
    stocktotal = 0
         
    'Keep track of location of each ticker in the summary
    Dim tablecount As Integer
    tablecount = 2

      
    'Define lastrows
    lastrow1 = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
      
    'Print headers of Ticker, yearly change, perecentage chance, total stock vol, greatest % inc, dec and greatest total vol.
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percentage Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("N2").Value = "Greatest % increase"
    ws.Range("N3").Value = "Greatest % decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
    
    
     
        
    'loop through all the stock prices
    For i = 2 To lastrow1
            
            
            
            'Set Ticker to hold value
            ticker = ws.Cells(i, 1).Value
            
            'if ticket symbol is not the same as the next one
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                 
            'add the stock volume to total
            stocktotal = stocktotal + ws.Cells(i, 7).Value
            
            'print the ticker onto the summary table
            ws.Cells(tablecount, 9).Value = ticker
                                    
            'print the stock volume total onto the summary table
            ws.Cells(tablecount, 12).Value = stocktotal
            
            'Reset the stock total to 0
            stocktotal = 0
                      
            'set closing price
            closing = ws.Cells(i, 6).Value
            'MsgBox (Cells(i, 1).Value + " " + Str(closing) + " " + Str(Cells(i, 2).Value))
            
            'to calculate yearly change
            yearlychange = (closing - opening)
            'MsgBox (Str(yearlychange) + Cells(i, 1).Value)
            ws.Cells(tablecount, 10).Value = yearlychange
                        
        
            'to calculate percentage change
            If opening > 0 Then percentchange = yearlychange / opening
            'MsgBox (Str(percentchange) + Cells(i, 1).Value)
            ws.Cells(tablecount, 11).Value = (percentchange * 100) & "%"
            
            
            'Add one to row to table
            tablecount = tablecount + 1
            
                        
            'Search for the first row for each ticker
            ElseIf ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            opening = ws.Cells(i, 3).Value
            'MsgBox (Cells(i, 1).Value + " " + Str(opening) + " " + Str(Cells(i, 2).Value))
            
            'add the stock volume to total
            stocktotal = stocktotal + ws.Cells(i, 7).Value
            
            
            'add the stock volume to total
            Else: stocktotal = stocktotal + ws.Cells(i, 7).Value
            
            If yearlychange > 0 Then
            ws.Cells(tablecount, 10).Interior.ColorIndex = 4
            
            Else: ws.Cells(tablecount, 10).Interior.ColorIndex = 3
            
            
            
                  
            End If
            End If
    Next i
  
    lastrow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
  
        For i = 2 To lastrow2
                                       
            
            If (ws.Cells(i, 11).Value > 0 And ws.Cells(i, 11).Value > Greatestinc) Then
            Greatestinc = ws.Cells(i, 11).Value
            tickerGI = ws.Cells(i, 9).Value
            'MsgBox (greatestinc)
            'MsgBox (tickerGI)
            
            End If
            
            If (ws.Cells(i, 11).Value < 0 And ws.Cells(i, 11).Value < Greatestdec) Then
            Greatestdec = ws.Cells(i, 11).Value
            tickerGD = ws.Cells(i, 9).Value
            'MsgBox (greatestdec)
            'MsgBox (tickerGD)
            End If
            
            If ws.Cells(i, 12).Value > greatestvol Then
            greatestvol = ws.Cells(i, 12).Value
            tickerGT = ws.Cells(i, 9).Value
            'MsgBox (greatestvol)
            'MsgBox (tickerGT)
            End If
            
            ws.Range("O2").Value = tickerGI
            ws.Range("O3").Value = tickerGD
            ws.Range("O4").Value = tickerGT
            ws.Range("P2").Value = Greatestinc
            ws.Range("P3").Value = Greatestdec
            ws.Range("P4").Value = greatestvol
            
     Next i
    Next ws
    
End Sub



