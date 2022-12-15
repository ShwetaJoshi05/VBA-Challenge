Attribute VB_Name = "StockAnalysis"
Sub StockAnalysis()
        
      For Each ws In Worksheets
                
          Dim WorksheetName As String
          'current row
          Dim i, j, row As Long
          'column
          Dim column As Long
          'Ticker counter to fill Ticker row
          Dim TickerIndex As Long
         'Initial value of ticker opening value
          Dim open_value As Double
          'Initial value of ticker closing value
           Dim close_value As Double
          'Last row in column A
           Dim LastRow As Long
          'last row column I
           Dim LastRowI As Long
           'Variable for greatest increase calculation
           Dim GreatIncrease As Double
           'Variable for greatest decrease calculation
           Dim GreatDecrease As Double
           'Variable for greatest total volume
           Dim GreatVolume As Double
          'Variable for percent change calculation
           Dim PerChange As Double

           'Get the WorksheetName
           WorksheetName = ws.Name
                
           'Determine the Last Row in column
            LastRow = ws.Cells(Rows.Count, 1).End(xlUp).row
                  
           'Create column headers
            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Yearly Change"
            ws.Cells(1, 11).Value = "Percent Change"
            ws.Cells(1, 12).Value = "Total Stock Volumn"
            ws.Cells(2, 15).Value = "Greatest % Increase"
            ws.Cells(3, 15).Value = "Greatest % Decrease"
            ws.Cells(4, 15).Value = "Greatest Total Volume"
            ws.Cells(1, 16).Value = "Ticker"
            ws.Cells(1, 17).Value = "Value"
                 
           'Set Ticker Counter to first row
            TickerIndex = 2
                 
           'Set start row to 2
            row = 2
                            
           ' Loop through rows in the column
            For j = 2 To LastRow
                    
            'Check if ticker name changed
             If ws.Cells(j + 1, 1).Value <> ws.Cells(j, 1).Value Then
                           
                  'Write ticker in column I (#9)
                   ws.Cells(TickerIndex, 9).Value = ws.Cells(j, 1).Value
                           
                   'Calculate and write Yearly Change in column J (#10)
                   ws.Cells(TickerIndex, 10).Value = ws.Cells(j, 6).Value - ws.Cells(row, 3).Value
                           
                        'conditional formating
                        If ws.Cells(TickerIndex, 10).Value < 0 Then
                        'set cell background color to red
                        ws.Cells(TickerIndex, 10).Interior.ColorIndex = 3
                        Else
                        'set cell background color to green
                        ws.Cells(TickerIndex, 10).Interior.ColorIndex = 4
                        End If
                           
                   'Calculate and write percent change in column K (#11)
                        If ws.Cells(row, 3).Value <> 0 Then
                        PerChange = ((ws.Cells(j, 6).Value - ws.Cells(row, 3).Value) / ws.Cells(row, 3).Value)
                               
                        'Percent formating
                         ws.Cells(TickerIndex, 11).Value = Format(PerChange, "Percent")
                        
                         Else
                         ws.Cells(TickerIndex, 11).Value = Format(0, "Percent")
                         End If
                                                   
                     'Calculate and write total volume in column L (#12)
                     Dim s As String
                     's = "G" + CStr(row) + ":" + "G" + CStr(j)
                     
                      ws.Cells(TickerIndex, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(row, 7), ws.Cells(j, 7)))
                      'ws.Cells(TickerIndex, 12).Value = Sum(Range(ws.Cells(row, 7), ws.Cells(i, 7)))
                           
                      'Increase TickCount by 1
                       TickerIndex = TickerIndex + 1
                                                                                 
                      'Set new start row of the ticker block
                       row = j + 1
                 End If
                     
            Next j
                             
          'find last row in column I
          LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).row
                 
          GreatIncrease = ws.Cells(2, 11).Value
          GreatDecrease = ws.Cells(2, 11).Value
         GreatVolume = ws.Cells(2, 12).Value
                 
                 
                   'Loop for summary
                    For i = 2 To LastRowI
                    
                        'For greatest total volume--check if next value is larger--if yes take over a new value and populate ws.Cells
                        If ws.Cells(i, 12).Value > GreatVolume Then
                        GreatVolume = ws.Cells(i, 12).Value
                        ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                        
                        Else
                        
                        GreatVolume = GreatVolume
                        
                        End If
                        
                        'For greatest increase--check if next value is larger--if yes take over a new value and populate ws.Cells
                        If ws.Cells(i, 11).Value > GreatIncrease Then
                        GreatIncrease = ws.Cells(i, 11).Value
                        ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                        
                        Else
                        
                        GreatIncrease = GreatIncrease
                        
                        End If
                        
                        'For greatest decrease--check if next value is smaller--if yes take over a new value and populate ws.Cells
                        If ws.Cells(i, 11).Value < GreatDecrease Then
                        GreatDecrease = ws.Cells(i, 11).Value
                        ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                        
                        Else
                        
                        GreatDecrease = GreatDecrease
                        
                        End If
                        
                    'Write summary results in ws.Cells
                    ws.Cells(2, 17).Value = Format(GreatIncrease, "Percent")
                    ws.Cells(3, 17).Value = Format(GreatDecrease, "Percent")
                    ws.Cells(4, 17).Value = Format(GreatVolume, "Scientific")
                    
                    Next i
                    
                    'Djust column width automatically
                    Worksheets(WorksheetName).Columns("A:Z").AutoFit
                 
         
         Next ws
        
End Sub
    
    
    
