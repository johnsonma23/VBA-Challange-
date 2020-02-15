Sub Stock()
For Each ws In Worksheets

    'Set the varibles to hold the Ticker"
Dim Ticker As String
    'Set the Varibles to hold the Ticker Volume'
Dim Volume As Double

    'Set the place Holder for the table'
 Dim Stock_Table As Integer
    'Place the row of the Stock'
Stock_Table = 2

Dim YearB As Double
Dim YearEnd As Double
Dim YearC As Double
Dim PercentC As Double
Dim LastRow As Long


LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'MsgBox (LastRow)


    





    'Set the loop"
Dim T As Long

For T = 2 To LastRow




  If T = 2 Then
   YearB = ws.Cells(T, 3).Value
    
    End If
    


    'Find the mismatch"
    
    If ws.Cells(1 + T, 1).Value <> ws.Cells(T, 1).Value Then
    Ticker = ws.Cells(T, 1).Value
    
    
        'Print to table Ticker to the table"
    ws.Range("J" & Stock_Table).Value = Ticker
    
    
    
    
    YearEnd = ws.Cells(T, 6).Value
    YearC = YearEnd - YearB
    
    
    
    
    ws.Range("K" & Stock_Table).Value = YearC
    
    If YearC < 0 Then
    ws.Range("K" & Stock_Table).Interior.ColorIndex = 3
    
    Else: ws.Range("K" & Stock_Table).Interior.ColorIndex = 4
    End If
    
    
    'Formate the color'
    
    
    
    ws.Range("L" & Stock_Table).Value = PercentC
    
    PercentC = (((YearEnd - YearB) / YearB) * 100)
    
    ws.Range("L" & Stock_Table).Value = PercentC
    
    
    
    YearB = ws.Cells(T + 1, 3).Value
    
     'Print the Volume'
    Volume = Volume + ws.Cells(T, 7).Value
    ws.Range("M" & Stock_Table).Value = Volume
    
    Stock_Table = Stock_Table + 1
    
    'Print the Yearly Change of the Ticker'
    
    
    
    'Print the Percent Change'
    


    
    Volume = 0
    
    
    
    Else
        'If the are the same add the volume'
        
    Volume = Volume + ws.Cells(T, 7).Value
    
    

    
    
    End If
    
    
    Next T
    
    Next ws
    
End Sub
