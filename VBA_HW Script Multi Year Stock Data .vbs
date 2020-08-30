Sub WorksheetLoop()
    
' Loop through all sheets
Dim ws As Worksheet
For Each ws In ActiveWorkbook.Worksheets
   ws.Activate
    ' Determine last row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Set Headers for output
    Cells(1, "I").Value = "Ticker"
    Cells(1, "J").Value = "Yearly Change ($)"
    Cells(1, "K").Value = "Percent Change"
    Cells(1, "L").Value = "Total Stock Volumn"
    
    'Create Variables for Values
    Dim Ticker As String
    Dim Yearly_Change As Double
    Dim Open_Price As Double
    Dim Closing_Price As Double
    Dim Percent_Change As Double
    Dim Volumn As Double
        Volumn = 0
    Dim Summary_Row As Long
        Summary_Row = 2
    Open_Price = Cells(2, 3).Value
        'Loop through ticker symbol
        For r = 2 To LastRow
            If Cells(r + 1, 1).Value <> Cells(r, 1).Value Then
                'find values
                Ticker = Cells(r, 1).Value
                Cells(Summary_Row, 9).Value = Ticker
                
                Close_Price = Cells(r, 6).Value
                Yearly_Change = Close_Price - Open_Price
                Cells(Summary_Row, 10).Value = Yearly_Change
                If Open_Price <> 0 Then
                    Percent_Change = (Yearly_Change / Open_Price)
                End If
                Cells(Summary_Row, 11).Value = Format(Percent_Change, "Percent")
                
                
                Volumn = Volumn + Cells(r, 7).Value
                Cells(Summary_Row, 12).Value = Volumn
                
                'Add row to summary table
                Summary_Row = Summary_Row + 1
                'reset values
                Open_Price = Cells(r + 1, 3)
                Volumn = 0
                
                Else
                    Volumn = Volumn + Cells(r, 7).Value
                    
                End If
        
        Next r
        
        ' Determine last row of Yearly Change column per worksheet
        JLastRow = ws.Cells(Rows.Count, 10).End(xlUp).Row
        For j = 2 To JLastRow
            If (Cells(j, 10).Value > 0 Or Cells(j, 10).Value = 0) Then
                Cells(j, 10).Interior.ColorIndex = 4
            ElseIf Cells(j, 10).Value < 0 Then
                Cells(j, 10).Interior.ColorIndex = 3
            End If
        Next j
        

    'create summary table for challenge
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volumn"
    
 For s = 2 To JLastRow
    If Cells(s, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & JLastRow)) Then
        Cells(2, 16).Value = Cells(s, 9).Value
        Cells(2, 17).Value = Format(Cells(s, 11).Value, "Percent")
    ElseIf Cells(s, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & JLastRow)) Then
        Cells(3, 16).Value = Cells(s, 9).Value
        Cells(3, 17).Value = Format(Cells(s, 11).Value, "Percent")
    ElseIf Cells(s, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & JLastRow)) Then
        Cells(4, 16).Value = Cells(s, 9).Value
        Cells(4, 17).Value = Cells(s, 12).Value
    End If
    
    Next s
    

        
    Next ws
        
End Sub