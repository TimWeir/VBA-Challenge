Attribute VB_Name = "Module1"
Sub Stock()

'Loop worksheets
Dim sheet As Worksheet
For Each sheet In Worksheets


'Declare variables for table results
Dim TickerName As String
Dim Great_Increase As String
Dim Great_Decrese As String
Dim Great_Volume As String
Dim YearStart As Double
Dim YearClose As Double
Dim YearChange As Double
Dim PercentChange As Double
Dim Stock_Volume As Double
Stock_Volume = 0

'Declare variables for table Structure
Dim Table_Row As Integer
Table_Row = 2

'Declare variable for <ticker> column end
Dim ColEnd As Long
'Get last row for table headers
ColEnd = sheet.Cells(Rows.Count, 1).End(xlUp).Row

'Count worksheets in the workbook
Sheet_Count = ActiveWorkbook.Worksheets.Count

'Table header labels
sheet.Cells(1, 10).Value = "Ticker"
sheet.Cells(1, 11).Value = "Yearly Change"
sheet.Cells(1, 12).Value = "Yearly Percentage"
sheet.Cells(1, 13).Value = "Yearly Volume"
sheet.Cells(1, 16).Value = "Bonus Metric Label"
sheet.Cells(1, 17).Value = "Ticker Name"
sheet.Cells(1, 18).Value = "Bonus Metric Result"
sheet.Cells(2, 16).Value = "Greatest % increase"
sheet.Cells(3, 16).Value = "Greatest % decrease"
sheet.Cells(4, 16).Value = "Greatest total volume"

'Table calculation variables
Dim RowRef As Double
RowRef = 2

'Loop structure for Ticker name
For i = 2 To ColEnd

        'Find last instance of each
        If sheet.Cells(i + 1, 1).Value <> sheet.Cells(i, 1).Value Then
        
            'Retreive the ticker name
            TickerName = sheet.Cells(i, 1).Value
            
            'Add ticker volume to Stock_Volume
            Stock_Volume = Stock_Volume + sheet.Cells(i, 7).Value
            
            'Deposit the ticker name in the summary
            sheet.Cells(Table_Row, 10) = TickerName
            
            'Retrieve opening price
            YearStart = sheet.Cells(RowRef, 3).Value
            
            'Retrieve closing price
            YearClose = sheet.Cells(i, 6).Value
                
            'Calcluate yearly change from opening price
            YearChange = YearClose - YearStart
            
            'Deposit Year Change
            sheet.Cells(Table_Row, 11) = YearChange
            
            'Color Change YearChange cell
            If sheet.Cells(Table_Row, 11).Value < 0 Then
                sheet.Cells(Table_Row, 11).Interior.Color = vbRed
            Else
                sheet.Cells(Table_Row, 11).Interior.Color = vbGreen
            End If
            
            'Calculate the change percentage and deposit
            If YearStart = 0 Then
                sheet.Cells(Table_Row, 12) = 1
            Else
                sheet.Cells(Table_Row, 12) = (YearChange / YearStart)
            End If

            'Deposit the ticker volume
            sheet.Cells(Table_Row, 13) = Stock_Volume
                        
            'Adjust table row to next row
            Table_Row = Table_Row + 1
            'Adjust rowref for yearstart capture
            RowRef = i + 1
            'Reset ticker volume amount
            Stock_Volume = 0
        Else
            Stock_Volume = Stock_Volume + sheet.Cells(i, 7).Value
        End If
                           
Next i

'Declare variable for <ticker> column end
Dim ColEnd2 As Long
'Get last row for table headers
ColEnd2 = sheet.Cells(Rows.Count, 10).End(xlUp).Row

For n = 2 To ColEnd2
    If sheet.Cells(n, 12) = sheet.Range("R2") Then
        Great_Increase = sheet.Cells(n, 10)
    End If
Next n
     
For o = 2 To ColEnd2
     If sheet.Cells(o, 12) = sheet.Range("R3") Then
        Great_Decrease = sheet.Cells(o, 10)
    End If
Next o

For p = 2 To ColEnd2
     If sheet.Cells(p, 13) = sheet.Range("R4") Then
        Great_Volume = sheet.Cells(p, 10)
    End If
Next p


'Deposit and format Greatest Increase
sheet.Cells(2, 17).Value = Great_Increase
sheet.Cells(2, 18).Value = WorksheetFunction.Max(sheet.Range("L2:L" & ColEnd))
sheet.Cells(2, 18).NumberFormat = "#,##0.00%"

'Deposit and format Greatest Decrease
sheet.Cells(3, 17).Value = Great_Decrease
sheet.Cells(3, 18).Value = WorksheetFunction.Min(sheet.Range("L2:L" & ColEnd))
sheet.Cells(3, 18).NumberFormat = "#,##0.00%"

'Deposit and format Greatest Volume
sheet.Cells(4, 17).Value = Great_Volume
sheet.Cells(4, 18).Value = WorksheetFunction.Max(sheet.Range("M2:M" & ColEnd))
sheet.Cells(4, 18).NumberFormat = "#,##0"

'Auto-fit columns for bonus metrics
sheet.Range("P1:P4").EntireColumn.AutoFit
sheet.Range("Q1:Q4").EntireColumn.AutoFit

'Auto-fit columns for change/percent/volume columns
sheet.Range("K2:K" & ColEnd).NumberFormat = "#,##0.00"
sheet.Range("K2:K" & ColEnd).EntireColumn.AutoFit

sheet.Range("L2:L" & ColEnd).NumberFormat = "#,###.00%"
sheet.Range("L2:L" & ColEnd).EntireColumn.AutoFit

sheet.Range("M2:M" & ColEnd).NumberFormat = "#,##0"
sheet.Range("M2:M" & ColEnd).EntireColumn.AutoFit

Next sheet

End Sub
