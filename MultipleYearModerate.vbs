Attribute VB_Name = "Module2"
Sub MultipleYearModerate()
   
  ' ' --------------------------------------------
  ' LOOP THROUGH ALL SHEETS
  ' --------------------------------------------
  For Each ws In Worksheets
         
    ' Created a Variable to Hold Last Row
    Dim LastRow As Long

    ' Determine the Last Row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
          
    ' Define Ticker Summary Table
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
  
    ' Set variable for holding the Ticker Symbol
    Dim Ticker_Symbol As String

    ' Set an initial variable for Ticker Yearly change
    Dim Ticker_Yearly_Change As Variant
    Ticker_Yearly_Change = 0

    ' Set an initial variable for Ticker Percent Change
    Dim Ticker_Percent_Change As Double
        
    ' Set an initial variable for holding the Ticker Total Stock Volume
    Dim Ticker_Total_Volume As Double
    Ticker_Total_Volume = 0

    ' Keep track of the location for each Ticker Symbol in the summary table
    Dim Ticker_Summary_Table_Row As Integer
    Ticker_Summary_Table_Row = 2
   
    
    ' Set variable for Open and Close Price
    Dim Open_Price As Variant
    Open_Price = ws.Cells(2, 3).Value
    
    ' Dim Close_Price As Variant
    Dim Close_Price As Variant
    Close_Price = 0
 
    
    ' Loop through all Ticker Symbols
        For i = 2 To LastRow
                                            
            ' Check if we are still within the Ticker, if it is not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                 
                ' Set the Ticker Symbol
                Ticker_Symbol = ws.Cells(i, 1).Value
                        
                ' Set the Open Price
                ' Open_Price = ws.Cells(i - 261, 3).Value
                                
                'Set the Close Price
                Close_Price = ws.Cells(i, 6).Value
                                            
                ' Get Ticker Yearly Change
                Ticker_Yearly_Change = Close_Price - Open_Price

                If Open_Price <> 0 Then

                ' Get the Ticker Percent Change
                Ticker_Percent_Change = (Close_Price - Open_Price) / Open_Price
                
                Else
                
                End If
                
                ' Add to the Ticker Total Stock Volume
                Ticker_Total_Volume = Ticker_Total_Volume + ws.Cells(i, 7).Value
            
                ' Print the Ticker Symbol in the Summary Table
                ws.Range("I" & Ticker_Summary_Table_Row).Value = Ticker_Symbol
            
                ' Print Ticker Yearly Change
                ws.Range("J" & Ticker_Summary_Table_Row).Value = Ticker_Yearly_Change
            
                ' Print Ticker Percent Change
                ws.Range("K" & Ticker_Summary_Table_Row).Value = Ticker_Percent_Change
                            
                ' Print the Ticker Total Stock Volume to the Summary Table
                ws.Range("L" & Ticker_Summary_Table_Row).Value = Ticker_Total_Volume

                ' Add one to the Ticker summary table row
                Ticker_Summary_Table_Row = Ticker_Summary_Table_Row + 1
                          
                'Add one to Open Price row
                'Open_Price = 0
                
                ' Reset the Ticker Total Stock Volume count
                Ticker_Total_Volume = 0
            
                ' Reset the Ticker Yearly Change
                Ticker_Yearly_Change = 0
                
                ' Reset the Ticker Percent Change
                Ticker_Percent_Change = 0
                
                ' Add one to reposition Open Price
                Open_Price = ws.Cells(i + 1, 3).Value
                                                       

            ' If the cell immediately following a row is the same Ticker...
            Else

            ' Add to the Ticker Total Stock Volume
            Ticker_Total_Volume = Ticker_Total_Volume + ws.Cells(i, 7).Value
                       
            End If
                      
        Next i
        
        ' Autofit to display data
        ws.Columns("I:L").AutoFit
        ws.Columns("K").NumberFormat = "0.00%"
        
    Next ws
    
End Sub


