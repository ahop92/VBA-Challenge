Attribute VB_Name = "Module11"
Sub YearlyStockEval()

    Dim nexti As Integer
    Dim summary_row_index As Integer
    Dim summary_col_index As Integer
    Dim worksheet_counter As Integer
    Dim total_worksheets As Integer
    
    Dim yearly_pchange As Double
    Dim opening_price As Double
    Dim closing_price As Double
    
    Dim lastrow As Long
    Dim lastcol As Long
    Dim i As Long
    Dim j As Long
    
    Dim total_stockvolume As Variant
    Dim yearly_pchange_prct As Variant
    Dim greatest_prct_increase As Variant
    Dim greatest_prct_decrease As Variant
    Dim greatest_total_stockvolume As Variant
    Dim compare_increase As Variant
    Dim compare_decrease As Variant
    Dim compare_stockvolume As Variant
      
    Dim tickersymbol As String
    Dim next_tickersymbol As String
    Dim greatest_increase_ticker As String
    Dim greatest_decrease_ticker As String
    Dim greatest_stockvolume_ticker As String
    
    

    'Find the number of worksheets in the workbook
    'ref: https://excelchamps.com/vba/count-sheets/
 
    total_worksheets = ThisWorkbook.Sheets.Count
    'MsgBox ("There are " & total_worksheets & " worksheets total in this workbook")


    'Creating the primary for loop that will iterate through each worksheet where we want to complete the data analysis

    For worksheet_counter = 1 To total_worksheets

        Worksheets(worksheet_counter).Select
        'MsgBox ("You are currently appending worksheet number " & worksheet_counter)

        'Find the last row, and column of the current worksheet
        'ref: https://www.excelcampus.com/vba/find-last-row-column-cell/
        'ref: automateexcel.com/vba/xldown-xlup-xltoright-xltoleft/
    
        lastrow = Range("A1").End(xlDown).Row
        lastcol = Range("A1").End(xlToRight).Column
        'MsgBox ("There are " & lastrow & " in this sheet")
        'MsgBox ("There are " & lastcol & " in this sheet")
        
        
        'Setup Summary Table Columns
        
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        
        
        'Initializing variables for looping through the raw data
    
        total_stockvolume = 0
        yearly_pchange = 0
        summary_row_index = 2
        summary_col_index = 9
        opening_price = Cells(2, 3).Value
    
      
        'For loop to move from the first row to the last row of the raw data on the current worksheet
     
        For i = 2 To lastrow
        
            tickersymbol = Cells(i, 1).Value
            next_i = i + 1
            next_tickersymbol = Cells(next_i, 1).Value
            
            
            'MsgBox (tickersymbol)
            'MsgBox (next_tickersymbol)
            
            
            'Adding to the total stock volume for a given ticker and determining when the ticker symbol has changed in the loop
            
            
            total_stockvolume = Cells(i, lastcol).Value + total_stockvolume
             
            
            If tickersymbol <> next_tickersymbol Then
            
            
                'Extract difference between the beginning of year open price and end of year close price for the year and depost in the summary table
                
                closing_price = Cells(i, lastcol - 1).Value
                
                
                'Convert that change into a percentage
                
                If opening_price = 0 Then
                
                    yearly_pchange = closing_price
                    yearly_pchange_prct = "N/A - Opening price was 0"
                    Cells(summary_row_index, 11).HorizontalAlignment = xlRight
                    
                    Else
                    
                    yearly_pchange = closing_price - opening_price
                    yearly_pchange_prct = Format((yearly_pchange / opening_price), "Percent")
                
                End If
                
                
                Cells(summary_row_index, 9).Value = tickersymbol
                Cells(summary_row_index, 10).Value = yearly_pchange
                Cells(summary_row_index, 11).Value = yearly_pchange_prct
                Cells(summary_row_index, 12).Value = total_stockvolume
                
                
                'Conditional formatting the cells (red for negative change, green for positive)
            
            
                If yearly_pchange > 0 Then
                
                    Cells(summary_row_index, 10).Interior.ColorIndex = 4
                    
                    ElseIf yearly_pchange < 0 Then
                    
                    Cells(summary_row_index, 10).Interior.ColorIndex = 3
                    
                    Else
                    
                    Cells(summary_row_index, 10).Interior.ColorIndex = 2
                
                End If
                    
                
                
                'MsgBox (tickersymbol & "  is a stock symbol in the list")
                'MsgBox ("The current total stock volume is " & total_stockvolume)
                'MsgBox ("The calculated price change for this ticker is " & yearly_pchange)
                'MsgBox (i & " " & lastcol)
                
                
                'Moving to the next row in the summary table, establishing the beginning of the year opening price for the next ticker, and reseting the total stock volume zum to zero
                
                summary_row_index = summary_row_index + 1
                opening_price = Cells(next_i, 3).Value
                total_stockvolume = 0
            
            
            End If
        
        Next i
        
        
        'For loop through the summary table to identify greatest percent increase, greatest decrease
        'and greatest total volume
        
        lastrow = Range("I1").End(xlDown).Row
        greatest_prct_increase = Range("K2").Value
        greatest_prct_decrease = Range("K2").Value
        greatest_total_stockvolume = Range("L2").Value
        compare_increase = 0
        compare_decrease = 0
        compare_stockvolume = 0
        greatest_increase_ticker = " "
        greatest_decrease_ticker = " "
        greatest_stockvolume_ticker = " "
        
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        
        
        'MsgBox (greatest_prct_increase & " " & greastest_prct_decrease & " " & greastest_total_volume)
        
        
        For i = 2 To lastrow
        
            Do While Cells(i + 1, 11).Value = "N/A - Opening price was 0"
            
                i = i + 1
            
            Loop
            
        
            compare_increase = Cells(i + 1, 11).Value
            compare_decrease = Cells(i + 1, 11).Value
            compare_stockvolume = Cells(i + 1, 12).Value
            
            'MsgBox ("Current: " & greastest_prct_increase & " " & greatest_prct_decrease & " " & greatest_total_stockvolume)
            'MsgBox ("Next: " & compare_increase & " " & compare_decrease & " " & compare_stockvolume)
            
            
            If compare_increase > greatest_prct_increase Then
            
                greatest_prct_increase = compare_increase
                greatest_increase_ticker = Cells(i + 1, 9).Value
            
            End If
            
            If compare_decrease < greatest_prct_decrease Then
                
                greatest_prct_decrease = compare_decrease
                greatest_decrease_ticker = Cells(i + 1, 9).Value
            
            End If
            
            If compare_stockvolume > greatest_total_stockvolume Then
            
                greatest_total_stockvolume = compare_stockvolume
                greatest_stockvolume_ticker = Cells(i + 1, 9).Value
            
            End If
        
        Next i
        
        
        
        Cells(2, 17).Value = Format(greatest_prct_increase, "Percent")
        Cells(3, 17).Value = Format(greatest_prct_decrease, "Percent")
        Cells(4, 17).Value = greatest_total_stockvolume
        Cells(2, 16).Value = greatest_increase_ticker
        Cells(3, 16).Value = greatest_decrease_ticker
        Cells(4, 16).Value = greatest_stockvolume_ticker
        
        
    
    Next worksheet_counter
    
        




End Sub

