Attribute VB_Name = "Module1"
Sub VBAStocks():
  ' Set a variable for the ticker symbol
  Dim Ticker As String
  
  'Set a variable for Opening Price
  Dim Opening As Double
  Opening = 0
  
  'Set an initial variable for Closing Price
  Dim Closing As Double
  Closing = 0
    
  'Set a variable for Yearly change
  Dim Yearly As Double
  Yearly = 0
  
  'Set a variable for Percent change
  Dim Percent As Double
  Percent = 0
  
  'Set a variable for Total Stock Volume
  Dim TSV As Double
  TSV = 0
  
  'Headers for Summary Table
     Range("I1").Value = "Ticker"
     Range("J1").Value = "Opening Price"
     Range("K1").Value = "Closing Price"
     Range("L1").Value = "Yearly Change"
     Range("M1").Value = "Percent Change"
     Range("N1").Value = "Total Stock Volume"
     
  ' Keep track of the location for ticker symbol in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
      
  ' Loop through all ticker symbols
  For i = 2 To 705714
    'Check the first row for ticker symbols
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        'Set the value for ticker symbols
        Ticker = Cells(i, 1).Value
        
        'Set the value for Total Stock Volume
        TSV = TSV + Cells(i, 7).Value
    
        'Display ticker symbol
        Range("I" & Summary_Table_Row).Value = Ticker
    
        'Display Total Stock Volume
        Range("N" & Summary_Table_Row).Value = TSV
        
        'Display information vertically
        Summary_Table_Row = Summary_Table_Row + 1
        
        'Initialize Total Stock Volume
        TSV = 0
        
        'If the following row has the same ticker symbol...
        
        Else
    
        'Add to the Total Stock Volume
        TSV = TSV + Cells(i, 7).Value
        
        'Search for opening date
        If Cells(i, 2).Value = "20140101" Then
        
        'Set value for opening price
        Opening = Cells(i, 3).Value
        
        'Display Opening Price
        Range("J" & Summary_Table_Row).Value = Opening
        End If
        End If
        
        'Search for closing date
        If Cells(i, 2).Value = "20141231" Then
        
        'Set value for closing date
        Closing = Cells(i, 6).Value
        
        'Display Closing Price
        Range("K" & Summary_Table_Row).Value = Closing
                       
        'Setting the value for Yearly change
        Yearly = (Closing - Opening)
                   
        'Display Yearly change
        Range("L" & Summary_Table_Row).Value = Yearly
        
        'Set value for Percent change
        Percent = Yearly / Opening
        
        'Display Percent change
        Range("M" & Summary_Table_Row).Value = Percent
        
        'Conditional formatting of Yearly change
        If Cells(i, 12) > 0 Then Cells(i, 12).Interior.Color = 10 'Green
        If Cells(i, 12) < 0 Then Cells(i, 12).Interior.Color = 3 'Red
        
        'Conditional formatting of Percent change
        If Cells(i, 13) > 0 Then Cells(i, 13).Interior.Color = 10 'Green
        If Cells(i, 13) < 0 Then Cells(i, 12).Interior.Color = 3 'Red
        
        End If
    Next i
 End Sub


