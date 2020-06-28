Attribute VB_Name = "Module1"

Sub Stock_Data()

  Dim ws As Worksheet
  
  'LOOP THROUGH ALL SHEETS
 
  For Each ws In Worksheets
 
        'set an intital variable
        
         Dim ticker As String
    
    
        'set an intial varaible for holding the total stock volume
        
         Dim Total_Stock_Volume As Double
        
         Total_Stock_Volume = 0
    
    
        ' keep track of the location of each stock ticker type in the summary
         
          Dim Summary_Table_Row As Long
         
          Summary_Table_Row = 2
          
     
        ' sets an inital variable
        
         Dim Pctchange As Double
         
         Dim Yearlychange As Double
         
         Dim Yearclosed As Double
         
         Dim Yearopen As Double
         
         Dim Start As Long
         
         Dim non_zero_row As Long
         
     
     
    
        ' finds the last row
    
         Dim Lastrow As Long
         Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
         
         ' The first row after the title row
         
           Start = 2
    
         ' keept track of percentage change in the summery
         
           Pctchange = 0
    
         'loop through all the rows in the ticker column starting from the top
    
          For i = 2 To Lastrow
    
    
                   ' checks when the current ticker is different from the next one
        
                If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1) Then
                
                        
                        'set the ticker
                    
                         ticker = ws.Cells(i, 1).Value
                        
                        'Add to the total stock Volume
            
                         Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                                     
                                     
            
                        
                            If Total_Stock_Volume = 0 Then
           
                                 
                                 
                                  ws.Range("I" & Summary_Table_Row).Value = 0
             
                                
                                
                                 ws.Range("J" & Summary_Table_Row).Value = 0
              
              
                                 ws.Range("K" & Summary_Table_Row).Value = 0
                                 ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
             
                                 ws.Range("L" & Summary_Table_Row).Value = 0
              
                           Else
                                 'Find the first non zero starting value
                                 
                                     If ws.Cells(Start, 3).Value = 0 Then
                     
                                        For non_zero_row = Start To i
                        
                                            If ws.Cells(non_zero_row, 3) <> 0 Then
                         
                                                Start = non_zero_row
                         
                                                Exit For
                                
                                            End If
                         
                                        Next non_zero_row
                         
                   
                                     End If
                                    
                   
                                    Yearopen = ws.Cells(Start, 3).Value
            
                                    ' sets the yearclosed
                                    Yearclosed = ws.Cells(i, 6).Value
            
                                    ' The yealychange is the difference between Yearclosed and Yearopen
            
                                    Yearlychange = Yearclosed - Yearopen
            
                                   
                                   ' The Percentage change is the yearlychange divided by the year opening price
            
                                    Pctchange = Yearlychange / Yearopen
            
            
                                     
            
                                     
                                     'Print the tickers in the summery
                                      ws.Range("I" & Summary_Table_Row).Value = ticker
             
                                     'Print the year closed in the summery
             
                                      ws.Range("N" & Summary_Table_Row).Value = Yearclosed
             
                                      ws.Range("M" & Summary_Table_Row).Value = Yearopen
             
                                     'Print the yearlychange in the summery
                                    
                                     ws.Range("J" & Summary_Table_Row).Value = Yearlychange
             
                                     ' If the yearly change is postive it shades the cells green(4) if it is negative red(3)
             
                                                If Yearlychange >= 0 Then
              
                                                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                 
                                                Else
                 
                                                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                 
                                                End If
                 
                 
             
                                   'Print the Percentage change in the summery
             
                                    ws.Range("K" & Summary_Table_Row).Value = Pctchange
                                    ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
             
             
             
                                    'Print the total stock volume in the summery
                                    
                                    ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
                                    
                                    Start = i + 1
        
                                    
                                    
            
            
                            End If
        
                                    'Reset all the following values
                                    
                                    Total_Stock_Volume = 0
                                    
                                    Yearlychange = 0
                                    
                                    Summary_Table_Row = Summary_Table_Row + 1
            
            
                                 
            
        
         ' If the following cell in the row is the same ticker
        
        Else
        
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
            
            
           
            
        
        
        End If
    
    
    
    
    Next i
    
          'Using the built in Min/Max function to find the greatest increase and decrease
    
            ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & Lastrow)) * 100
    
            ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & Lastrow)) * 100
    
            ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & Lastrow))
    
    
            'using the built in Match function in VBA to find the exact ticker type by looking up the MAX and Min value in the designated range
    
    
            increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & Lastrow)), ws.Range("K2:K" & Lastrow), 0)
    
            decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & Lastrow)), ws.Range("K2:K" & Lastrow), 0)
    
            volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & Lastrow)), ws.Range("L2:L" & Lastrow), 0)
    
            'It puts the exact ticker type matched into specifci cell
    
            ws.Range("P2") = ws.Cells(increase_number + 1, 9)
    
            ws.Range("P3") = ws.Cells(decrease_number + 1, 9)
    
            ws.Range("P4") = ws.Cells(volume_number + 1, 9)
    
    
    
    
    
    
    
    
 Next ws
 
 MsgBox ("Fixes Complete")
    
    
       
End Sub




