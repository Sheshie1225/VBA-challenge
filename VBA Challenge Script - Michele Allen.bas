Attribute VB_Name = "Module1"
Sub Wall_Street():


For Each ws In Worksheet

'Input Column Name

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

'Autofit the columns
Columns("I:L").AutoFit


Dim Ticker As String
Dim YearlyChange As Double
Dim PercentChange As Double
Dim Volume As Integer
Dim YearOpen As Double
Dim YearClose As Double
Dim Summary_Table_Row As Double
Dim lastrow As Double

lastrow = ws.Cells(Rows.Count, 1).End(x1Up).Row

'Setup integers for loop
Summary_Table_Row = 2

For i = 2 To lastrow

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
        Ticker = ws.Cells(i, 1).Value
        Volume = Volume + ws.Cells(i, 7).Value
        
        ws.Range("I" & Summary_Table_Row).Value = Ticker
        ws.Range("L" & Summary_Table_Row).Value = Volume
        
        Volume = 0
        
       YearClose = ws.Cells(i, 6)
        
        If YearOpen = 0 Then
            YearlyChange = 0
            PercentChange = 0
        Else:
            YearlyChange = YearClose - YearOpen
            PercentChange = (YearClose - YearOpen) / YearOpen
        End If
        

            ws.Range("J" & Summary_Table_Row).Value = YearlyChange
            
            ws.Range("K" & Summary_Table_Row).Value = PercentChange
            ws.Range("K" & Summary_Table_Row).Style = "Percent"
            ws.Range("K" And Summary_Table_Row).NumberFormat = "0.00%"
            
            Summary_Table_Row = Summary_Table_Row + 1
        
        ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1) Then
            YearOpen = ws.Cells(i, 3)
            
            
        Else: Volume = Volume + ws.Cells(i, 7).Value
        
        End If
        
          
    'Move to next worksheet
    Next i



End Sub

            
            
        
