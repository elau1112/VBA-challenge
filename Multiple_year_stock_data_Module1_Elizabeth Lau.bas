Attribute VB_Name = "Module1"
Sub stock()

For Each ws In Worksheets

Dim Ticker As String
Dim Yearly_Change As Double
Dim Yearly_Open As Double
Dim Percent_Change As Double
Dim Total_Stock_Volume As Double

Dim Start_Date As Double
Dim End_Date As Double

Dim Max_Increase As Double
Dim Max_Ticker As String
Dim Max_Decrease As Double
Dim Min_Ticker As String
Dim Max_TotalVolume As Double
Dim Max_TickerVolume As String


Total_Stock_Volume = 0
Yearly_Change = 0
Yearly_Open = 0
Percent_Change = 0

Start_Date = 0
End_Date = 0

Dim Summary_Table_Row As Integer
Summary_Table_Row = 2
j = 2

LR = ws.Cells(Rows.Count, "A").End(xlUp).Row

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

ws.Range("M1").Value = "Start Date of the Year"
ws.Range("N1").Value = "End Date of the Year"

ws.Range("P2").Value = "Greatest % increase"
ws.Range("P3").Value = "Greatest % decrease"
ws.Range("P4").Value = "Greatest total volume"
ws.Range("Q1").Value = "Ticker"
ws.Range("R1").Value = "Value"

For i = 2 To LR
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Ticker = Cells(i, "A").Value
        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, "G").Value
        Yearly_Open = ws.Cells(j, "C").Value
        Yearly_Change = ws.Cells(i, "F").Value - Yearly_Open
        
        Start_Date = ws.Cells(j, "B").Value
        End_Date = ws.Cells(i, "B").Value
        

        If Yearly_Change <> 0 & Yearly_Open <> 0 Then
            Percent_Change = Yearly_Change / Yearly_Open
        Else
            Percent_Change = 0
        End If

j = i + 1
'MsgBox (j)

ws.Range("I" & Summary_Table_Row).Value = Ticker
ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
    If Yearly_Change > 0 Then
        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
    ElseIf Yearly_Change < 0 Then
        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
    Else
        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 0
    End If
ws.Range("K" & Summary_Table_Row).Value = Percent_Change
ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"

ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume

ws.Range("M" & Summary_Table_Row).Value = Start_Date
ws.Range("N" & Summary_Table_Row).Value = End_Date

Summary_Table_Row = Summary_Table_Row + 1
Yearly_Change = 0
Yearly_Open = 0
Total_Stock_Volume = 0

Start_Date = 0
End_Date = 0

Else
Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, "G").Value

End If
Next i

LRC = ws.Cells(Rows.Count, "K").End(xlUp).Row



Max_Increase = WorksheetFunction.Max(ws.Range("K2", "K" & LRC))
Max_Ticker = "Not Found Yet"

Max_Decrease = WorksheetFunction.Min(ws.Range("K2", "K" & LRC))
Min_Ticker = "Not Found Yet"

Max_TotalVolume = WorksheetFunction.Max(ws.Range("L2", "L" & LRC))
Max_TickerVolume = "Not Found Yet"

For x = 2 To LRC
    If ws.Range("K" & x) = Max_Increase Then
        Max_Ticker = ws.Range("I" & x)
    Else
        Max_Ticker = Max_Ticker
    End If
Next x

For y = 2 To LRC
    If ws.Range("K" & y) = Max_Decrease Then
        Min_Ticker = ws.Range("I" & y)
    Else
        Min_Ticker = Min_Ticker
    End If
Next y

For w = 2 To LRC
    If ws.Range("L" & w) = Max_TotalVolume Then
        Max_TickerVolume = ws.Range("I" & w)
    Else
        Max_TickerVolume = Max_TickerVolume
    End If
Next w

ws.Range("Q2").Value = Max_Ticker
ws.Range("R2").Value = Max_Increase
ws.Range("R2").NumberFormat = "0.00%"

ws.Range("Q3").Value = Min_Ticker
ws.Range("R3").Value = Max_Decrease
ws.Range("R3").NumberFormat = "0.00%"

ws.Range("Q4").Value = Max_TickerVolume
ws.Range("R4").Value = Max_TotalVolume

Next ws
End Sub
