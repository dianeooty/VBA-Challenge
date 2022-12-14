Attribute VB_Name = "Module1"
Sub WorksheetLoop()

For Each ws In Worksheets

'Declaring variables and types
Dim Lrow As Long
Dim OpenP As Double
Dim CloseP As Double
Dim Year As Double
Dim Percent As Double
Dim i As Long
Dim ticker As String
Dim SumRow As Double
Dim Vol As Double
Dim Max As Double
Dim Min As Double
Dim gv As Double

'Setting up the headers for the two new summary tables
ws.Range("J1").Value = "Ticker"
ws.Range("K1").Value = "Yearly Change"
ws.Range("L1").Value = "Percent Change"
ws.Range("M1").Value = "Total Stock Volume"
ws.Range("Q1").Value = "Ticker"
ws.Range("R1").Value = "Value"
ws.Range("P2").Value = "Greatest % Increase"
ws.Range("P3").Value = "Greatest % Decrease"
ws.Range("P4").Value = "Greatest Total Volume"


'Grabbing the row count of column A
Lrow = ws.Cells(Rows.Count, "A").End(xlUp).Row

'Setting row count to start at 2
SumRow = 2

'Setting up start values for 2nd summary table
Max = 0
Min = 0
gv = 0

    'Using for loop to check each row starting at row 2 in column A for ticker names
    For i = 2 To Lrow
    
        'Using If statement to check if the row above the current does not contain the same
        'ticker name. If not, then current has the open price and adding a new ticker name
        If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
            OpenP = ws.Cells(i, 3).Value
            ticker = ws.Cells(i, 1).Value
            
            'Assigning a value of 0 for stock volume since this is a different ticker name
            Vol = 0
            'Adding the stock volume of the current row to the previous volume total
            Vol = Vol + ws.Cells(i, 7).Value

        'Closing if statement
        End If
        
        'Using another if statement to check if the current row has the same ticker
        'as the row above and below.
        'Continue to add to stock volume with each row if True
        If ws.Cells(i - 1, 1).Value = ws.Cells(i + 1, 1).Value Then
            Vol = Vol + ws.Cells(i, 7).Value
            
        'Closing if statement
        End If
        
        'Using another if statement to check if the row below the current does not contain
        'the same ticker.  If True, then current has the close price and a new ticker name added
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            CloseP = ws.Cells(i, 6).Value
            
            'Now we have the open and close price to calculate yearly and percent change
            Year = CloseP - OpenP
            Percent = (Year / OpenP) * 100 / 100

            'Continue to add to the stock volume
            Vol = Vol + ws.Cells(i, 7).Value
            
            'Assigning the values for rows in the summary table with each loop round
            ws.Cells(SumRow, 10).Value = ticker
            ws.Cells(SumRow, 11).Value = Year
            ws.Cells(SumRow, 12).Value = Percent
            ws.Cells(SumRow, 13).Value = Vol
                               
            'Conditional statements to determine the max & min percent change and max stock volume
            'Grabbing the ticker name and values to print in 2nd summary table
            If Percent >= Max Then
                ws.Cells(2, 18).Value = ws.Cells(SumRow, 12).Value
                Max = Percent
                ws.Cells(2, 17).Value = ws.Cells(SumRow, 10).Value
            End If
             
            If Percent <= Min Then
                ws.Cells(3, 18).Value = ws.Cells(SumRow, 12).Value
                Min = Percent
                ws.Cells(3, 17).Value = ws.Cells(SumRow, 10).Value
            End If
            
            If Vol >= gv Then
                ws.Cells(4, 18).Value = ws.Cells(SumRow, 13).Value
                gv = Vol
                ws.Cells(4, 17).Value = ws.Cells(SumRow, 10).Value
            End If
                                          
            'Adding to sumrow for new count
            SumRow = SumRow + 1
                                         
        'Closing If statement
        End If
                          
    'Next statement to go round the loop
    Next i

    'Formatting the worksheet
    'changing number format to percentage for the percent change column
    ws.Range("L:L").NumberFormat = "0.00%"
    'Changing number format for the second summary table
    ws.Range("R2:R3").NumberFormat = "0.00%"
    
    'Deleting the current format conditions for this range
    
    'With statement to format column k interior color based on value
    With ws.Range("K:K")
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Yearly Change"""
        .FormatConditions(1).Interior.Color = RGB(255, 255, 255)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
        .FormatConditions(2).Interior.Color = RGB(255, 0, 0)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
        .FormatConditions(3).Interior.Color = RGB(0, 255, 0)
    End With

    'Changing the column size for the two summary tables to autofit the data added
    ws.Range("J:R").Columns.AutoFit

'Going to the next worksheet
Next ws
 
End Sub



