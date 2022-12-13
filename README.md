# VBA Challenge
> This challenge was a Module 2 Challenge for my Data Analytics and Visualization Boot Camp.  

## Table of Contents
* [General Info](#general-information)
* [Technologies Used](#technologies-used)
* [Screenshots](#screenshots)
* [Usage](#usage)
* [Project Status](#project-status)
* [Room for Improvement](#room-for-improvement)
* [Acknowledgements](#acknowledgements)
* [Contact](#contact)


## General Information
I was tasked with writing a scriot that analyzes generated stock data of the years 2018, 2019 and 2020. For all worksheets in the workbook, the script outputs two summary tables in each sheet(year), with conditional formatting to highlight positive changes in green and negative changes in red, all in one run.  The first summary table includes a column of the unique ticker symbols, the yearly change of the opening price to the closing price, the percent change of the opening price to the closng price and the total stock volume of that year.  The second summary includes the stock with the greatest percent increase, greatest percent decrease and greatest stock volume with the corresponding ticker symbols for each sheet(year).


## Technologies Used
Excel and Visual Basic editor


## Screenshots
![2018](https://user-images.githubusercontent.com/117790100/207243880-041e7efd-ffcf-47c1-83a4-be099843748c.png)
![2019](https://user-images.githubusercontent.com/117790100/207243885-e87d355a-284b-44cd-b8ff-695f148974a2.png)
![2020](https://user-images.githubusercontent.com/117790100/207243898-ed4a8d21-c983-46b3-a0e5-5bc45342ee1c.png)


## Usage
To analyze the stock data and output the two summary tables for each worksheet in the workbook, run this script.

`Sub WorksheetLoop()

For Each ws In Worksheets

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

ws.Range("J1").Value = "Ticker"
ws.Range("K1").Value = "Yearly Change"
ws.Range("L1").Value = "Percent Change"
ws.Range("M1").Value = "Total Stock Volume"
ws.Range("Q1").Value = "Ticker"
ws.Range("R1").Value = "Value"
ws.Range("P2").Value = "Greatest % Increase"
ws.Range("P3").Value = "Greatest % Decrease"
ws.Range("P4").Value = "Greatest Total Volume"

Lrow = ws.Cells(Rows.Count, "A").End(xlUp).Row

SumRow = 2
Max = 0
Min = 0
gv = 0

    For i = 2 To Lrow
 
        If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
            OpenP = ws.Cells(i, 3).Value
            ticker = ws.Cells(i, 1).Value
            Vol = 0
            Vol = Vol + ws.Cells(i, 7).Value
        End If
        
        If ws.Cells(i - 1, 1).Value = ws.Cells(i + 1, 1).Value Then
            Vol = Vol + ws.Cells(i, 7).Value
        End If
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            CloseP = ws.Cells(i, 6).Value
            Year = CloseP - OpenP
            Percent = (Year / OpenP) * 100 / 100

            Vol = Vol + ws.Cells(i, 7).Value
            
            ws.Cells(SumRow, 10).Value = ticker
            ws.Cells(SumRow, 11).Value = Year
            ws.Cells(SumRow, 12).Value = Percent
            ws.Cells(SumRow, 13).Value = Vol
                               
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
                                          
            SumRow = SumRow + 1
                                         
        End If
                          
    Next i

    ws.Range("L:L").NumberFormat = "0.00%"
            'Changing number format for the second summary table
    ws.Range("R2:R3").NumberFormat = "0.00%"

    With ws.Range("K:K")
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Yearly Change"""
        .FormatConditions(1).Interior.Color = RGB(255, 255, 255)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
        .FormatConditions(2).Interior.Color = RGB(255, 0, 0)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
        .FormatConditions(3).Interior.Color = RGB(0, 255, 0)
    End With
    
    ws.Range("J:R").Columns.AutoFit

Next ws
 
End Sub`


## Project Status
Project is complete and no longer being worked on.


## Acknowledgements
Many thanks to my amazing tutor, learning instructors, and TAs for guiding me through this learning process.


## Contact
Created by Diane Guzman
 


