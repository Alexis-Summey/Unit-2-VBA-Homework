Attribute VB_Name = "Module1"
Sub stock_data()

'let's name some variables

Dim ws As Worksheet
Dim ticker As String
Dim vol As Integer
Dim Year_Open As Double
Dim Year_Close As Double
Dim Percent_Change As Double
Dim Summary_Table_Row As Integer

'run through worksheets


Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

' Staring on my loops

Dim Ticker_Name As String

Dim Yearly_Change As Double

Yearly_Change = 0

Dim Percent_Change As Double

Percent_Change = 0

Dim Volume As Integer

Volume = 0

Dim Year_Open As Double

Year_Open = 0

Dim Year_Close As Double

Year_Close = 0

Dim Summary_Table_Row As Integer

Summary_Table_Row = 2


For i = 2 To 705714

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    Ticker_Name = ws.Cells(i, 1).Value
    
    Year_Close = Year_Close + Cells(i, 6).Value
    
    Year_Open = Year_Open + Cells(i, 3).Value
    
    Yearly_Change = Year_Close - Year_Open
    
    Percent_Change = ((Year_Close - Year_Open) / Year_Close)
    
    Volume = Volume + Cells(i, 7).Value
    
    
' putting the values in the summary table

        
    Range("I" & Summary_Table_Row).Value = Ticker_Name
    
    Range("J" & Summar_Table_Row).Value = Yearly_Change
    
    Range("K" & Summary_Table_Row).Value = Percent_Change
    
    Range("L" & Summary_Table_Row).Value = Volume
    
Else
   Ticker_Name = ws.Cells(i, 1).Value
    Year_Open = Year_Open + ws.Cells(i, 3).Value
    Year_Close = Year_Close + ws.Cells(i, 6).Value
    Volume = Volume + ws.Cells(i, 7).Value

End If

Next i



End Sub
