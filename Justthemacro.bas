Attribute VB_Name = "Module1"
Sub TestModule()

Dim a As Integer

Dim ws As Worksheet

' Loop through all of the worksheets in the active workbook.
For Each ws In Worksheets
    
             
    'Declaring variables
    'Also sets intitial value to 0 for quantities
    
    Dim TickerName As String
    
    'Dim i As Double
    
    Dim YearlyChange As Double
            YearlyChange = 0
            
    Dim PercentChange As Double
            PercentChange = 0
            
    Dim PercentChangeRounded As Double
    
            
    Dim TotalSV As Variant 'total stock volume
            TotalSV = 0
            
    Dim OpeningPrice As Double
           OpeningPrice = 0
           
    Dim ClosingPrice As Double
           ClosingPrice = 0
    
    Dim LastSummaryRow As Double
    
    Dim TickerMaxValue As Double
    
    Dim TickerMaxName As String
    
    Dim TickerMinValue As Double
    
    Dim TickerMinName As String
    
    Dim VolumeMaxValue As Double
    
    Dim VolumeMaxName As String
    
    
           
    'Daily change and average change add
    
    ws.Range("i1").Value = "ticker"
    ws.Range("j1").Value = "yearly change"
    ws.Range("k1").Value = "percent change"
    ws.Range("l1").Value = "total stock volume"
    ws.Range("o2").Value = "MaxChange"
    ws.Range("o3").Value = "MinChange"
    ws.Range("o4").Value = "MaxVolume"
    ws.Range("p1").Value = "Ticker"
    ws.Range("q1").Value = "Value"

    
    ' Keep track of the location for each credit card brand in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    'counts the number of rows
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    j = 0
    begin = 2
    

    ' Loop through each row
    For i = 2 To lastrow

        TotalSV = TotalSV + ws.Cells(i, 7).Value
        
        'Check if they have the same ticker symbol
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
                    
            'get yearly change in price
            ClosingPrice = ws.Cells(i, 6).Value
            OpeningPrice = ws.Cells(begin, 3).Value
            Change = ClosingPrice - OpeningPrice
            
            ws.Range("J" & Summary_Table_Row).Value = Round(Change, 2)
        
            PercentChange = Change / OpeningPrice
            ws.Range("K" & Summary_Table_Row).Value = Round(PercentChange, 2)
            ws.Range("K" & Summary_Table_Row).NumberFormat = "0%"
            
    
            
            'calculate the change next, equal to the ith row cells(i,6) - cells(begin,3)
            'percent change is dividing cells and multiplying by 100, rounding etc.
            'print results in report using summary
            
            
  If ws.Range("j" & Summary_Table_Row).Value > 0 Then
                ws.Range("j" & Summary_Table_Row).Interior.ColorIndex = 4
            Else
                ws.Range("j" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
                    
                        
                        
              
          
      
    'Set the Ticker name
            TickerName = ws.Cells(i, 1).Value
    
    'Set the opening price
            OpeningPrice = ws.Cells(i, 3).Value
    
    
      'Print the Ticker Name in the Summary Table
            ws.Range("I" & Summary_Table_Row).Value = TickerName
    
      'Print the Total Stock Volume to the Summary Table
            ws.Range("L" & Summary_Table_Row).Value = TotalSV
    
      'Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
    
    
      'Reset the Total Stock Volume
            TotalSV = 0
        


                

        End If




    Next i

               
Next ws
   
 

End Sub
