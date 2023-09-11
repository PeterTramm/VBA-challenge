Attribute VB_Name = "Module1"
Sub Summary_Table():
    'Set an inital value for holding ticker name
    Dim Ticker_name As String
    
    'Set an inital variable for holding Open Price at start of year
    Dim OpenPrice As Double
    OpenPrice = 1
    
    'Set an inital variable for holding Closing price at end of year
    Dim ClosePrice As Double
    ClosePrice = 1
    
    'Set an inital variable for holding Percent change for each ticker name
    Dim PChange As Double
    
    'Set an inital varaible for holding stock volume
    Dim Stock_volume As Double
    Stock_volume = 0
    
    
    'Equation to calculate percent change
    ' (New Price - Old Price)/Old price * 100 for stock increase
    
   'Finding last row for spreadsheet
   Dim LastRow As Double
   LastRow = Cells(Rows.Count, 1).End(xlUp).Row
   
    
    'Loop through all worksheets
    For Each ws In Worksheets
    
       'Add in Summary table column names for each worksheet
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        'Add second Sumary table for each worksheet
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
       
       'Keep Track of location for each Ticker name on in summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
          'Loop through all Tickers in one sheet
       For I = 2 To LastRow
            
            'Add Ticker name to Summary table
           
            'Check if it the first of the ticker name, if so add ticker name to summary table
            If ws.Cells(I, 1).Value <> ws.Cells(I - 1, 1) Then
                
                'Save open price value for later calculation
                OpenPrice = ws.Cells(I, 3).Value
               
            'Check if it the last ticker name
            ElseIf ws.Cells(I, 1).Value <> ws.Cells(I + 1, 1) Then
                'Save ticker_name
                Ticker_name = ws.Cells(I, 1).Value
                
                'Save Close price for later calculation
                ClosePrice = ws.Cells(I, 6).Value
                
                'Calculate Yearly change
                YChange = ClosePrice - OpenPrice
                
                'Calculate Percentage change
                PChange = YChange / OpenPrice
                PChange2 = PChange
              
                'Add in ticker name and total volume into summary table
                ws.Range("I" & Summary_Table_Row).Value = Ticker_name
                ws.Range("L" & Summary_Table_Row).Value = Stock_volume
              
                
               'Adding Yearly Change for ticker to summary table
                ws.Range("J" & Summary_Table_Row).Value = YChange
                
                'Adding colour format for yearly change'
                If YChange > 0 Then
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                ElseIf YChange < 0 Then
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
    
               'Adding Percent Change for ticker to summary table
                ws.Range("K" & Summary_Table_Row).Value = PChange
                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                
                
                'Add row to summary table row
                Summary_Table_Row = Summary_Table_Row + 1
                
                'reset stock volume
                Stock_volume = 0
                
            'Adding volume stock to volume stock counter
            ElseIf Cells(I, 1).Value = Cells(I + 1, 1) Then
                Stock_volume = Stock_volume + CDbl(Cells(I, 7).Value)
                End If
        Next I
        
        'Bonus
        Dim Highest_Percent_increase As Double
        Highest_Percent_increase = 0
        
        Dim Highest_Percent_decrease As Double
        Highest_Percent_decrease = 0
        
        Dim Highest_Total_volume As Double
        Highest_Total_volume = 0
        
        'Loop through summary table
        For I = 2 To LastRow
        
            'Finding highest percentage increase
            If ws.Cells(I, 11) > Highest_Percent_increase Then
            Highest_Percent_increase = ws.Cells(I, 11).Value
            ws.Range("Q2").Value = Highest_Percent_increase
            ws.Range("P2").Value = ws.Cells(I, 9).Value
            Else
                End If
            
            'Finding highest percentage decrease
            If ws.Cells(I, 11) < Highest_Percent_decrease Then
            Highest_Percent_decrease = ws.Cells(I, 11).Value
            ws.Range("Q3").Value = Highest_Percent_decrease
            ws.Range("P3").Value = ws.Cells(I, 9).Value
            Else
                End If
            
            'Finding highest total volume
            If ws.Cells(I, 12) > Highest_Total_volume Then
            Highest_Total_volume = ws.Cells(I, 12).Value
            ws.Range("Q4").Value = Highest_Total_volume
            ws.Range("P4").Value = ws.Cells(I, 9).Value
            Else
                End If
        Next I
        
        'Formating percentages for column Q
         ws.Range("Q2").NumberFormat = "0.00%"
         ws.Range("Q3").NumberFormat = "0.00%"
        'Formating columns to fit text in cells
        ws.Columns("A:Q").AutoFit
    Next ws
     
End Sub




