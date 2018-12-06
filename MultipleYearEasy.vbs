Attribute VB_Name = "Module1"
Sub MultipleYearEasy()
   
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
    ws.Cells(1, 10).Value = "Total Stock Volume"
  
    ' Set an initial variable for holding the Ticker Symbol
    Dim Ticker_Symbol As String

    ' Set an initial variable for holding the Ticker Total Stock Volume
    Dim Ticker_Total_Volume As Double
    Ticker_Total_Volume = 0

    ' Keep track of the location for each Ticker Symbolin the summary table
    Dim Ticker_Summary_Table_Row As Integer
    Ticker_Summary_Table_Row = 2

    ' Loop through all Ticker Symbols
        For i = 2 To LastRow

            ' Check if we are still within the Ticker, if it is not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            ' Set the Ticker Symbol
            Ticker_Symbol = ws.Cells(i, 1).Value

            ' Add to the Ticker Total Stock Volume
            Ticker_Total_Volume = Ticker_Total_Volume + ws.Cells(i, 7).Value

            ' Print the Ticker Symbol in the Summary Table
            ws.Range("I" & Ticker_Summary_Table_Row).Value = Ticker_Symbol

            ' Print the Ticker Total Stock Volume to the Summary Table
            ws.Range("J" & Ticker_Summary_Table_Row).Value = Ticker_Total_Volume

            ' Add one to the Ticker summary table row
            Ticker_Summary_Table_Row = Ticker_Summary_Table_Row + 1
      
            ' Reset the Ticker Total Stock Volume count
            Ticker_Total_Volume = 0

            ' If the cell immediately following a row is the same Ticker...
            Else

            ' Add to the Ticker Total Stock Volume
            Ticker_Total_Volume = Ticker_Total_Volume + ws.Cells(i, 7).Value

            End If

        Next i
        
        ' Autofit to display data
        ws.Columns("I:J").AutoFit
        
    Next ws
    
End Sub
