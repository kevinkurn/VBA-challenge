Attribute VB_Name = "Module1"
Sub SharesCalcFinal()
    'Create a loop in the worksheet
    For Each ws In Worksheets
        
        'sort the data first
        ws.Range("A1").CurrentRegion.Sort Key1:=ws.Range("A1"), Order1:=xlAscending, Header:=xlYes
    
        'set up all variables (ticker no., Opening Value, Closing Value, stock volume, etc)
        Dim ticker As String
        Dim OPValue As Double
        Dim CLValue As Double
        Dim Volume As Double
        Dim SummaryRow As Integer
        Dim ChangeValue As Double
        
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row 'set up last row of the data
        SummaryRow = 2 'starting row on activity output
        
        'set up loop to group share performance based on ticker
        For i = 2 To lastrow
            'set up if logic
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                'generate an output on ticker value, yearly change, and % change
                ticker = ws.Cells(i, 1).Value
                
                'update the OPValue, ClValue, and Volume
                OPValue = OPValue + ws.Cells(i, 3).Value
                CLValue = CLValue + ws.Cells(i, 6).Value
                Volume = Volume + ws.Cells(i, 7).Value
                
                'generate summary table (ticker, yearly change, percent change & stock volume
                ws.Range("I" & SummaryRow).Value = ticker
                ws.Range("J" & SummaryRow).Value = CLValue - OPValue
                
                'in the case whereby either one of opening value or closing value is zero, then "zero value will be generated"
                If (OPValue Or CLValue) <> 0 Then
                    ws.Range("K" & SummaryRow).Value = CLValue / OPValue - 1
                Else
                    ws.Range("K" & SummaryRow).Value = 0
                End If
                    
                ws.Range("L" & SummaryRow).Value = Volume
                
                'colour code relevant cells
                ChangeValue = ws.Range("J" & SummaryRow).Value
                
                    If ChangeValue > 0 Then
                    ws.Range("J" & SummaryRow, "K" & SummaryRow).Interior.ColorIndex = 4
                    
                    Else
                    ws.Range("J" & SummaryRow, "K" & SummaryRow).Interior.ColorIndex = 3
                
                    End If
                
                'change the format of the relevant cells
                ws.Range("K" & SummaryRow).NumberFormat = "0.00%"
                ws.Range("L" & SummaryRow).NumberFormat = "#,##0"
                
                'add 1 to summaryRow to allow the above program to list down the output
                SummaryRow = SummaryRow + 1
                
                'Reset Opening Value, Closing Value, and Volume to zero to start counting from a new ticker
                OPValue = 0
                CLValue = 0
                Volume = 0
            
            Else
                OPValue = OPValue + ws.Cells(i, 3).Value
                CLValue = CLValue + ws.Cells(i, 6).Value
                Volume = Volume + ws.Cells(i, 7).Value
             
            End If
            
        Next i
        
        'Update the headings for summary table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
    '------------------------
    'Bonus - extract the greatest % increase, greatest % decrease, and greatest total volume along ticker number
    '------------------------
    
    'Determine the new variables
        lastrowPer = ws.Cells(Rows.Count, 11).End(xlUp).Row
        lastrowVol = ws.Cells(Rows.Count, 12).End(xlUp).Row
    
    'Update the headings
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
    
    'Input the greatest % increase, greatest % decrease, and greatest total volume from summary table
        ws.Range("Q2").Value = WorksheetFunction.Max(ws.Range("K2:K" & lastrowPer))
        ws.Range("Q3").Value = WorksheetFunction.Min(ws.Range("K2:K" & lastrowPer))
        ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L2:L" & lastrowVol))
    
    'Fixed the numbering format
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        ws.Range("Q4").NumberFormat = "#,##0"
    
    'Input the ticker value based on the the value generated in the above
        Dim L As Integer
    
        For L = 2 To lastrowPer
            If ws.Range("Q2").Value = ws.Cells(L, 11).Value Then
                ws.Range("P2").Value = ws.Cells(L, 9).Value
        
            End If
    
            If ws.Range("Q3").Value = ws.Cells(L, 11).Value Then
                ws.Range("P3").Value = ws.Cells(L, 9).Value
        
            End If
        
            If ws.Range("Q4").Value = ws.Cells(L, 12).Value Then
                ws.Range("P4").Value = ws.Cells(L, 9).Value
        
            End If
        
        Next L
        
    Next ws
    
End Sub



