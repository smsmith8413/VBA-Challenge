Sub vbaChallengeWS()

'-----loop through all worksheets----------
For Each ws In Worksheets

  ' Set an initial variable for holding the ticker info
  Dim ticker As String

  ' Set an initial variable for holding the total stock volume
  Dim totalStock As Variant
  totalStock = 0

  ' set up a place to keep track of the data and assign headers
  Dim rowNum As Integer
  rowNum = 2
  
  ws.Range("i1").Value = "Ticker"
  ws.Range("j1").Value = "Yearly Change"
  ws.Range("k1").Value = "% Change"
  ws.Range("l1").Value = "Total Stock Volume"
  ws.Range("n2").Value = "Greatest % Increase"
  ws.Range("n3").Value = "Greatest % Descrease"
  ws.Range("n4").Value = "Greatest Total Volume"
  ws.Range("o1").Value = "Ticker"
  ws.Range("p1").Value = "Value"
  
  'dim open price, close price and yearly (might not have needed to do this)
  Dim openPrice As Double
  Dim closePrice As Double
  Dim yearlyChange As Double
  
'loop through to last row of data by setting last row variable for stock column (1)
  lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'--------------Loop throu all stocks and get total volume for each, begin populating summary section ---------------------- '

  For I = 2 To lastrow
    
    ' Check if we are still within the same ticker/stock, if it is not...
    If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then

      ' Set the ticker
      ticker = ws.Cells(I, 1).Value

      ' Add to the total stock volume
      totalStock = totalStock + ws.Cells(I, 7).Value
      
      ' Print the total Stock in the summary section
      ws.Range("i" & rowNum).Value = ticker

      ' Print the total stock in summary section
      ws.Range("l" & rowNum).Value = totalStock
        

      ' Reset the Stock Total
      totalStock = 0

    ' If the cell immediately following a row is the same stock...
    Else

      ' Add to the stock Total
      totalStock = totalStock + ws.Cells(I, 7).Value
         
    End If
    
'--------------get yearly change and percent change ----------------------
'*this works if the data is sorted in chronological order, oldest to newest, by stock. Consider sorting manually
'before running Sub or including a sort into the code if this could be an issue*

     ' Check if we are still within the same ticker/stock
    If ws.Cells(I + 1, 1).Value = ws.Cells(I, 1).Value Then

'if so, then
    'set open date
    openDate = ws.Cells(I, 2).Value
    
     'find row of open date by checking tot see if open date is less than cell after and before
    If (ws.Cells(I, 2).Value < ws.Cells(I - 1, 2).Value And ws.Cells(I, 2).Value < ws.Cells(I + 1, 2).Value) Then
    'set open price in column 3
    openPrice = ws.Cells(I, 3).Value
    'MsgBox (openDate & i)
    End If
    
    'if not, set to close price in column 6
   Else
   closePrice = ws.Cells(I, 6).Value
   'MsgBox (closePrice & i)
     'MsgBox (closePrice - openPrice)
   'set yearly change to close price - open
    yearlyChange = closePrice - openPrice
    'MsgBox (yearlyChange)
    
' Print the yearly change to the summary section
   ws.Range("j" & rowNum).Value = yearlyChange
   
   'dim % Change
    Dim percentChange As Double
    
    'account for any zero values in the data that could interrupt the calculation
        If yearlyChange = 0 Then
        percentChange = 0
        
        ElseIf openPrice = 0 Then
        percentChange = 1
        
        'calculate % change
        Else
        percentChange = yearlyChange / openPrice
        'MsgBox (percentChange)
        End If
    
    ' Print the % change to summary section
    ws.Range("k" & rowNum).Value = percentChange

    ' Add one to the summary table row
      rowNum = rowNum + 1
    End If
    
  Next I
  
  'set variable for stock in challenge section
  Dim ticker2 As String
  'set new last row based last row of summary section
  lastrow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
  
'---------Loop through summary section to get greatest % increase, % decrease, and total volume-------------------
  
  For w = 2 To lastrow2
  
  'find max change in % change column
    If ws.Cells(w, 11).Value = WorksheetFunction.Max(ws.Range("k2:k" & lastrow2)) Then
        ticker2 = ws.Cells(w, 9).Value
  'set greatest % increase
        greatestInc = ws.Cells(w, 11).Value
   'print stock and gpi
        ws.Range("o2") = ticker2
        ws.Range("p2") = greatestInc
        
  'find min change in % change column
    ElseIf ws.Cells(w, 11).Value = WorksheetFunction.Min(ws.Range("k2:k" & lastrow2)) Then
 'set greatest % increase
      ticker2 = ws.Cells(w, 9).Value
      greatestDec = ws.Cells(w, 11).Value
   'print stock and gpd
        ws.Range("o3") = ticker2
        ws.Range("p3") = greatestDec

  'find max volume
    ElseIf ws.Cells(w, 12).Value = WorksheetFunction.Max(ws.Range("l2:l" & lastrow2)) Then
   'set greatest total volume
      ticker2 = ws.Cells(w, 9).Value
      greatestVol = ws.Cells(w, 12).Value
   'print stock and gtv
        ws.Range("o4") = ticker2
        ws.Range("p4") = greatestVol
   
   End If
  
  Next w
 
 '----------------Loop through summary section and format columns
   For Z = 2 To lastrow2

'set decrease (less than 0) to red
    If ws.Cells(Z, 10).Value >= 0 Then
        ws.Cells(Z, 10).Interior.ColorIndex = 4
'set increase (greater than or equal ton 0) to green
    Else
        ws.Cells(Z, 10).Interior.ColorIndex = 3
        
 
     End If
  Next Z
  
 'format % values to percent (is there a way to do this when setting the variable?)
 ws.Range("k2:k" & lastrow).NumberFormat = "0.00%"
 ws.Range("p2:p3").NumberFormat = "0.00%"
 
 Next ws
 
 'the end!
 
End Sub

