Attribute VB_Name = "Module1"
Sub Stocks_Loops()
    Dim Ticker_Name As String
    Dim Summary_Table_Row As Integer
    Dim Total_Volume As Double
    Dim Opening_Price As Double
    Dim Closing_Price As Double
    Dim Percent_Change As Double
    Dim Yearly_Change As Double
    Dim Last_Row As Long
    Dim Max_Increase As Double
    Dim Max_Ticker As String
    Dim Min_Increase As Double
    Dim Min_Ticker As String
    Dim Volume_Increase As Double
    Dim Volume_Ticker As String

    Summary_Table_Row = 2
    Total_Volume = 0
    Opening_Price = Cells(2, 3).Value
    Closing_Price = 0
    Percent_Change = 0
    Last_Row = Cells(Rows.Count, 1).End(xlUp).Row
    Max_Increase = 0
    Min_Increase = 9999999999#
    Volume_Increase = 0

    Range("I" & 1).Value = "Ticker"
    Range("J" & 1).Value = "Yearly Change"
    Range("K" & 1).Value = "Percent Change"
    Range("L" & 1).Value = "Total Stock Volume"
    Range("O" & 2).Value = "Greatest % Increase"
    Range("O" & 3).Value = "Greatest % Decrease"
    Range("O" & 4).Value = "Greatest Total Volume"
    Range("P" & 1).Value = "Ticker"
    Range("Q" & 1).Value = "Value"

    For i = 2 To Last_Row
        Total_Volume = Total_Volume + Cells(i, 7).Value
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Ticker_Name = Cells(i, 1).Value
            Closing_Price = Cells(i, 6).Value
            Yearly_Change = Closing_Price - Opening_Price
            If Opening_Price > 0 Then
                Percent_Change = Yearly_Change / Opening_Price
            Else
                Percent_Change = 0
            End If
            If Yearly_Change > Max_Increase Then
                Max_Increase = Yearly_Change
                Max_Ticker = Cells(i, 1).Value
            End If
            If Yearly_Change < Min_Increase Then
                Min_Increase = Yearly_Change
                Min_Ticker = Cells(i, 1).Value
            End If
            If Total_Volume > Volume_Increase Then
                Volume_Increase = Total_Volume
                Volume_Ticker = Cells(i, 1).Value
            End If
            Range("I" & Summary_Table_Row).Value = Ticker_Name
            Range("J" & Summary_Table_Row).Value = Yearly_Change
            Range("J" & Summary_Table_Row).NumberFormat = "0.00"
            Range("K" & Summary_Table_Row).Value = Percent_Change
            Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
            Range("L" & Summary_Table_Row).Value = Total_Volume
            If Yearly_Change < 0 Then
                Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            Else
                Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            End If
            Summary_Table_Row = Summary_Table_Row + 1
            Opening_Price = Cells(i + 1, 3).Value
            Total_Volume = 0
        End If
    Next i
    Range("P" & 2).Value = Max_Ticker
    Range("P" & 3).Value = Min_Ticker
    Range("P" & 4).Value = Volume_Ticker
    Range("Q" & 2).Value = Max_Increase
    Range("Q" & 2).NumberFormat = "0.00%"
    Range("Q" & 3).Value = Min_Increase
    Range("Q" & 3).NumberFormat = "0.00%"
    Range("Q" & 4).Value = Volume_Increase
End Sub


