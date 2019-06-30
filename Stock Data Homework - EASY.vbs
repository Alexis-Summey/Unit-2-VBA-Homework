Attribute VB_Name = "Module2"
Sub Yearly_Stock_Data()

Dim Ticker_Type As String

Dim Volume_Total As Double

Volume_Total = 0

Dim Summary_Table_Row As Integer

Summary_Table_Row = 2

For i = 2 To 705714
    
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        Ticker_Type = Cells(i, 1).Value
        
        Volume_Total = Volume_Total + Cells(i, 7).Value
        
        Range("I" & Summary_Table_Row).Value = Ticker_Type
        
        Range("J" & Summary_Table_Row).Value = Volume_Total
        
        Summary_Table_Row 1
    
            Volume_Total = 0
        
    Else
        Volume_Total = Volume_Total + Cells(i, 7).Value
    
    End If

Next i

    
    
    
End Sub
