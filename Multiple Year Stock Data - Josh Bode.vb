Sub SummarizeAllSheets()

'By Josh Bode
'May 12, 2018

'The code loops through all sheets, analyzes stock prices and volume,
'and produces summary tables. The below code loops through the sheets
'and calls on a separate subroutine - SummarizeSheet - which does all
'of the work for the active sheet.

'Number of worksheets
Dim ws_count As Integer
ws_count = ActiveWorkbook.Worksheets.Count

'Summarize each worksheet
For w = 1 To ws_count
    Worksheets(w).Activate
    SummarizeSheet
Next w

'Message that we are done
MsgBox ("All " & w & " Sheets have been summarized")
    
End Sub


Sub SummarizeSheet()
   
'By Josh Bode
'May 12, 2018
   
'The Code analyses stock trading and produces:
'   1. A table summarizing the name, delta, %change, and volumes
'   2. A table identifying the stocks with the largest %increases, %decreases, and trade volume

'Main steps:
    '0. Prep work
        'Declare variables
        'Set seed values
        'Sort the data (to make sure it is in order)
        'Create Results table headers

    '1. Summarize all stocks in sheet
    '   a. Loop through all rows
    '   b. If we have a new stock (in next row) summarize results and reset values
    '2. Format the annual change depending on if its positive of negative
    '3. Identify stocks with greatest % increase, % decrease, and volume
    '4. Format results columns for readibility
    
'0. PREP WORK 

    'Declare variables
        Dim rowmax As Long
        Dim colmax As Long
        Dim stock As String
        Dim newstock As String
        Dim outputrow As Long
        Dim volume As Double
        Dim start As Double
        Dim pctchange As Double
        Dim delta As Double
    

    'Sort the data (to make sure it is in order)

        With ActiveSheet.Sort
            .SortFields.Add Key:=Range("A1"), Order:=xlAscending
            .SortFields.Add Key:=Range("B1"), Order:=xlAscending
            .SetRange Range("A1").CurrentRegion
            .Header = xlYes
            .Apply
        End With

    'Seed values for variables
        outputrow = 2
        volume = 0
        rowmax = ActiveSheet.UsedRange.Rows.Count
        colmax = ActiveSheet.UsedRange.Columns.Count
        stock = Cells(2, 1).Value
        start = Cells(2, 3).Value
        
    'Create Results Columns
        Cells(1, colmax + 2).Value = "Ticker"
        Cells(1, colmax + 3).Value = "Yearly Change"
        Cells(1, colmax + 4).Value = "Percent Change"
        Cells(1, colmax + 5).Value = "Total Stock Volume"

'1. Summarize All Stocks in Sheet

    For Row = 2 To rowmax
        stock = Cells(Row, 1).Value
        newstock = Cells(Row + 1, 1).Value
        
        volume = volume + Cells(Row, 7)
          
        'If we have a new stock name record values and reset
        If stock <> newstock Then
                          
            'Write stock name
            Cells(outputrow, colmax + 2).Value = stock
            
            'Write total change
            delta = Cells(Row, 6).Value - start
            Cells(outputrow, colmax + 3).Value = delta
            
            'Write percent change
            If start <> 0 Then
                pctchange = Cells(Row, 6).Value / start - 1
                Cells(outputrow, colmax + 4).Value = pctchange
            ElseIf start = 0 Then
                pctchange = 0
                Cells(outputrow, colmax + 4).Value = pctchange
            End If
            
            'Write volume
            Cells(outputrow, colmax + 5).Value = volume
             
            'Move to next output row
            outputrow = outputrow + 1
             
            'Reset start value
            start = Cells(Row + 1, 3).Value
            volume = 0
           
        End If
      
    Next Row

'2. Format the annual change depending on if its positive of negative
    endrow = outputrow - 1
    
    'Go through each row in output table and format color based on result
    For Row = 2 To endrow
        positive = Cells(Row, colmax + 3) >= 0
        
        If positive = "True" Then
            Cells(Row, colmax + 3).Interior.Color = RGB(0, 235, 30)
        
        Else
            Cells(Row, colmax + 3).Interior.Color = RGB(235, 20, 20)
        
        End If
        
    Next Row

'3. Identify stocks with greatest % increase, % decrease, and volume
    
    'Set up variables
    Dim max_change As Double
    Dim min_change As Double
    Dim max_volume As Double
    Dim name_max_change As String
    Dim name_min_change As String
    Dim name_max_volume As String
    
    max_change = 0
    min_change = 0
    max_volume = 0
    
    'Go through all and keep track
    For Row = 2 To endrow
        If Cells(Row, colmax + 4).Value >= max_change Then
            max_change = Cells(Row, colmax + 4).Value
            name_max_change = Cells(Row, colmax + 2).Value
        End If
        
        If Cells(Row, colmax + 4).Value < min_change Then
            min_change = Cells(Row, colmax + 4).Value
            name_min_change = Cells(Row, colmax + 2).Value
        End If
        
        If Cells(Row, colmax + 5).Value >= max_volume Then
            max_volume = Cells(Row, colmax + 5).Value
            name_max_volume = Cells(Row, colmax + 2).Value
        End If
    Next Row
    
    'Summarize results
    Cells(1, colmax + 7).Value = "Metric"
    Cells(1, colmax + 8).Value = "Ticker"
    Cells(1, colmax + 9).Value = "Value"
    
    Cells(2, colmax + 7).Value = "Greatest % Increase"
    Cells(2, colmax + 8).Value = name_max_change
    Cells(2, colmax + 9).Value = max_change
    
    Cells(3, colmax + 7).Value = "Greatest % Decrease"
    Cells(3, colmax + 8).Value = name_min_change
    Cells(3, colmax + 9).Value = min_change
       
    Cells(4, colmax + 7).Value = "Greatest Volume"
    Cells(4, colmax + 8).Value = name_max_volume
    Cells(4, colmax + 9).Value = max_volume


'4. Format results columns for readibility
     Range("J:J").NumberFormat = "$###,###.00"
     Range("K:K").NumberFormat = "0.00%"
     Range("L:L").NumberFormat = "###,###,###,###"
     Range("P2").NumberFormat = "0.00%"
     Range("P3").NumberFormat = "0.00%"
     Range("P4").NumberFormat = "###,###,###,###"
     Columns("L").ColumnWidth = 20
     Columns("N").ColumnWidth = 20
End Sub
