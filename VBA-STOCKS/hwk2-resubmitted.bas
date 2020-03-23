Attribute VB_Name = "Module1"
'Ticker symbol
'yearly change from opening price at the beginning of a given year to the closing price at the end of that year
'percent change from opening price at the beginning of a given year to the closing price at the end of that year
'total stock volume of the stock
'apply conditional formatting to highlight positive change in green and negative change in red

Sub stockAnalyzer()

 For Each ws In Worksheets
 

 'Definitions:
  Dim TickerSymbol As String
  Dim open_price_beg_year As Double
  Dim close_price_end_year As Double
  Dim Stock_Total As LongLong
  Dim sameticker_ctr As Integer
  Dim yearly_amount_change As Double
  Dim yearly_percent_change As Double
  Dim Summary_Table_Row As Integer
  Dim Last_Row As Long
  Dim temp_greatest_inc As Long
  Dim temp_greatest_inc_ticker As String
  Dim temp_greatest_dec As Long
  Dim temp_greatest_dec_ticker As String
  Dim greatest_idx As Integer
  Dim i As Long
  
  
  Summary_Table_Row = 2
  sameticker_ctr = 0
  Stock_Total = 0
  
  'apply changes to all sheets
  
  'For Each ws In Worksheets
  
     Dim worksheetName As String
    
     ws.Activate
     Debug.Print ws.Name
     
     'insert headers
     ws.Range("I1").Value = "Ticker"
     ws.Range("J1").Value = "Yearly Change"
     ws.Range("K1").Value = "percent Change"
     ws.Range("L1").Value = "Total Stock Volume"
     
     'insert total results table
     ws.Range("N2").Value = "Greatest % Increase"
     ws.Range("N3").Value = "Greatest % decrease"
     ws.Range("N4").Value = "Greatest Total Volume"
     ws.Range("O1").Value = "Ticker"
     ws.Range("P1").Value = "Value"
  
     Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
  ' Main loop through all tickers

       For i = 2 To Last_Row
  
      ' Check if we are about to change to a new ticker symbol, if so, then process output to sheets for the current ticker
    
         If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value And ws.Cells(i + 1, 1).Value <> " " Then

      ' print Stock Total
      
         Stock_Total = Stock_Total + ws.Cells(i, 7).Value
           
      'opening price from beginning of year
         If ws.Cells(i + 1, 1).Value <> "" Then
             open_price_beg_year = ws.Cells(i - sameticker_ctr, 3).Value
          
         End If
      
         'obtain closing price at the end of year and print it in summary table along with total volume for the current ticker
         close_price_end_year = ws.Cells(i, 6).Value
         ws.Range("I" & Summary_Table_Row).Value = ws.Cells(i, 1)
         ws.Range("L" & Summary_Table_Row).Value = Stock_Total
      
         'calculate yearly amount change and move it to summary table for the current ticker
         yearly_amount_change = close_price_end_year - open_price_beg_year
         ws.Range("J" & Summary_Table_Row).Value = yearly_amount_change

         'depending if the value is positive or negative color cells with green or red accordingly
      
         If (ws.Cells(Summary_Table_Row, 10).Value > 0) Then
            ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
         Else
            ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
         End If

        'calculate yearly_percent_change and move it to the summary table
         If open_price_beg_year > 0 Then
            yearly_percent_change = ((yearly_amount_change * 100) / open_price_beg_year)
         End If

         ws.Range("K" & Summary_Table_Row).Value = yearly_percent_change
      
         'increase the row counter for the summary table
          Summary_Table_Row = Summary_Table_Row + 1
      
         ' Reset the Stock Total
          Stock_Total = 0
          sameticker_ctr = 0

      Else
     
         'if the next ticker is the same as the current and not the last row, then increase counters and process next row.
         If ws.Cells(i, 1) <> " " Then
         ' Add to the Stock Total
          Stock_Total = Stock_Total + ws.Cells(i, 7).Value
          sameticker_ctr = sameticker_ctr + 1
       End If
       

      End If

    Next i
    
    'CHALLENGES: Your solution will also be able to return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume".
    'find the greatest percent increase
    For inc = 2 To Summary_Table_Row

          If ws.Cells(inc, 11).Value > 0 Then
          
               If temp_greatest_inc < ws.Cells(inc, 11) Then
                  temp_greatest_inc = ws.Cells(inc, 11).Value
                  temp_greatest_inc_ticker = ws.Cells(inc, 9).Value
               End If
              
          End If
          
     Next inc

    'find the greatest percent decrease
     For dec = 2 To Summary_Table_Row

        If ws.Cells(dec, 11).Value < 0 Then
          
               If temp_greatest_dec > ws.Cells(dec, 11) Then
                  temp_greatest_dec = ws.Cells(dec, 11).Value
                  temp_greatest_dec_ticker = ws.Cells(dec, 9).Value
               End If
              
          End If
          
       Next dec
       
    'find the greatest total volume
     For totvol = 2 To Summary_Table_Row
          
         If temp_greatest_totvol < ws.Cells(totvol, 12) Then
              temp_greatest_totvol = ws.Cells(totvol, 12).Value
              temp_greatest_totvol_ticker = ws.Cells(totvol, 9).Value
         End If
          
       Next totvol
             
       'print all to the greatest table
       ws.Cells(2, 15) = temp_greatest_inc_ticker
       ws.Cells(2, 16) = temp_greatest_inc
       ws.Cells(3, 15) = temp_greatest_dec_ticker
       ws.Cells(3, 16) = temp_greatest_dec
       ws.Cells(4, 15) = temp_greatest_totvol_ticker
       ws.Cells(4, 16) = temp_greatest_totvol
       
       
    Next ws
   
End Sub



