# VBA-challenge
Basic

Sub Stocks()
  
  
 For Each ws In Worksheets
         
     ' Create a Variable to Hold File Name, Last Row, and Year
        Dim WorksheetName As String
        
         WorksheetName = ws.Name
        
          
        
        
       
        Dim Closing As Double
        Dim Opening As Double
        Dim Ticker As String
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        Dim Yearly_Change As Double
        
        Dim Percent_Change As Double
        
        Dim Total_Volume As Double
        
        Dim Summary_table As Integer
        
        Summary_Table_Row = 2
        
        ws.Cells(1, 10).Value = "Ticker"
        ws.Cells(1, 11).Value = "Yearly Change"
        
        ws.Cells(1, 12).Value = "Percentage Change"
        ws.Cells(1, 13).Value = "Volume"
        ws.Columns("L").AutoFit
        ws.Columns("K").AutoFit
        ws.Columns("L").AutoFit
        ws.Columns("M").AutoFit
        
        
        
        Dim Summary_Table2 As Integer
        
        Summary_Table_Row2 = 2
       
       'add summary table titles
        
        ws.Cells(1, 17).Value = "Ticker"
        ws.Cells(1, 18).Value = "Value"
        ws.Cells(2, 16).Value = "Greatest % Increase"
        ws.Cells(3, 16).Value = "Greates % Decrease"
        ws.Cells(4, 16).Value = "Greatest Total Volume"
        ws.Columns("P").AutoFit
        ws.Columns("Q").AutoFit
        ws.Columns("R").AutoFit
        
     
 
    
    
   
 'loop through all rows
 
 For i = 2 To LastRow
 
 'loop through all columns
   
    
    
        
              
        
       
    
    
    
            'check when ticker is different
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                            
                'grab the ticker, put in summary table
                                
                Ticker = ws.Cells(i, 1).Value
                
                'grab the ticker, put in summary table
                
                ws.Range("J" & Summary_Table_Row).Value = Ticker
                
                
                Opening = ws.Cells(i, 3).Value And Closing = ws.Cells(i + 252, 6).Value
                
                
                Yearly_Change = ws.Cells(i + 252, 6).Value - ws.Cells(i, 3).Value
                
                
                ' store yeary change in summary table
                                
                
                ws.Range("K" & Summary_Table_Row).Value = Yearly_Change
                
                
                
                
                'Caluclate prenetage change
                
                
                Percent_Change = ws.Cells(i + 252, 6).Value - ws.Cells(i, 3).Value / ws.Cells(i, 3).Value
                
                
                ' Store Precentage Change Calculation
                
                ws.Range("L" & Summary_Table_Row).Value = Percent_Change
               
       
                
                


                       
            
                'sum total volume per ticker
                
                Total_Volume = Total_Volume + ws.Cells(i, 7).Value
                
                'take total sum of volume and put in summary table
                
                ws.Range("M" & Summary_Table_Row).Value = Total_Volume
                
                
                
                'Range("L" & Summary_Table_Row).Style = "Percentage"
                                
                           
                'rest
                
                
                Summary_Table_Row = Summary_Table_Row + 1
                
                 
                                
                Total_Volume = 0
            
            Else
                Total_Volume = Total_Volume + ws.Cells(i, 7).Value
            
             
            
            
                 
                
                
                End If
       
     
    
   
        
    Next i
    
    
Next ws

                
                
End Sub
    


