Attribute VB_Name = "Module1"
Sub stock()

Sheets.Add.Name = "Combined Stocks"
Sheets("Combined Stocks").Move Before:=Sheets(1)
Set Combined_sheet = Worksheets("Combined Stocks")

For Each ws In Worksheets
LastRow = Combined_sheet.Cells(Rows.Count, "A").End(xlUp).Row + 1
'MsgBox (LastRow)

LastRowState = ws.Cells(Rows.Count, "A").End(xlUp).Row - 1

'MsgBox (LastRowState)

Combined_sheet.Range("A" & LastRow & ":G" & ((LastRowState - 1) + LastRow)).Value = ws.Range("A2:G" & (LastRowState + 1)).Value

'MsgBox ("A" & LastRow & ":G" & (LastRowState - 1) + LastRow)
'MsgBox ("A2:G" & (LastRowState + 1))

Next ws
Combined_sheet.Range("A1:G1").Value = Sheets(2).Range("A1:G1").Value
Combined_sheet.Columns("A:G").AutoFit

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
Percent_Change = 0
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2
j = 2

LR = Combined_sheet.Cells(Rows.Count, "A").End(xlUp).Row
'MsgBox (LR)

Combined_sheet.Range("I1").Value = "Ticker"
Combined_sheet.Range("J1").Value = "Yearly Change"
Combined_sheet.Range("K1").Value = "Percent Change"
Combined_sheet.Range("L1").Value = "Total Stock Volume"

Combined_sheet.Range("M1").Value = "Start Date of the Year"
Combined_sheet.Range("N1").Value = "End Date of the Year"

Combined_sheet.Range("P2").Value = "Greatest % increase"
Combined_sheet.Range("P3").Value = "Greatest % decrease"
Combined_sheet.Range("P4").Value = "Greatest total volume"
Combined_sheet.Range("Q1").Value = "Ticker"
Combined_sheet.Range("R1").Value = "Value"

For i = 2 To LR
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Ticker = Cells(i, "A").Value
        Total_Stock_Volume = Total_Stock_Volume + Combined_sheet.Cells(i, "G").Value
        Yearly_Open = Combined_sheet.Cells(j, "C").Value
        Yearly_Change = Combined_sheet.Cells(i, "F").Value - Yearly_Open
        
        Start_Date = Combined_sheet.Cells(j, "B").Value
        End_Date = Combined_sheet.Cells(i, "B").Value
        

        If Yearly_Change <> 0 Then
            Percent_Change = Yearly_Change / Yearly_Open
            
        Else
            Percent_Change = 0
        End If

j = i + 1
'MsgBox (j)

Combined_sheet.Range("I" & Summary_Table_Row).Value = Ticker
Combined_sheet.Range("J" & Summary_Table_Row).Value = Yearly_Change
    If Yearly_Change > 0 Then
        Combined_sheet.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
    ElseIf Yearly_Change < 0 Then
        Combined_sheet.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
    Else
        Combined_sheet.Range("J" & Summary_Table_Row).Interior.ColorIndex = 0
    End If
Combined_sheet.Range("K" & Summary_Table_Row).Value = Percent_Change
Combined_sheet.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"

Combined_sheet.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume

Combined_sheet.Range("M" & Summary_Table_Row).Value = Start_Date
Combined_sheet.Range("N" & Summary_Table_Row).Value = End_Date

Summary_Table_Row = Summary_Table_Row + 1
Yearly_Change = 0
Total_Stock_Volume = 0

Start_Date = 0
End_Date = 0

Else
Total_Stock_Volume = Total_Stock_Volume + Combined_sheet.Cells(i, "G").Value


End If
Next i

LRC = Combined_sheet.Cells(Rows.Count, "K").End(xlUp).Row
'MsgBox (LRC)


Max_Increase = WorksheetFunction.Max(Combined_sheet.Range("K2", "K" & LRC))
Max_Ticker = "Not Found Yet"

Max_Decrease = WorksheetFunction.Min(Combined_sheet.Range("K2", "K" & LRC))
Min_Ticker = "Not Found Yet"

Max_TotalVolume = WorksheetFunction.Max(Combined_sheet.Range("L2", "L" & LRC))
Max_TickerVolume = "Not Found Yet"

For x = 2 To LRC
    If Combined_sheet.Range("K" & x) = Max_Increase Then
        Max_Ticker = Combined_sheet.Range("I" & x)
    Else
        Max_Ticker = Max_Ticker
    End If
Next x

For y = 2 To LRC
    If Combined_sheet.Range("K" & y) = Max_Decrease Then
        Min_Ticker = Combined_sheet.Range("I" & y)
    Else
        Min_Ticker = Min_Ticker
    End If
Next y

For w = 2 To LRC
    If Combined_sheet.Range("L" & w) = Max_TotalVolume Then
        Max_TickerVolume = Combined_sheet.Range("I" & w)
    Else
        Max_TickerVolume = Max_TickerVolume
    End If
Next w

Combined_sheet.Range("Q2").Value = Max_Ticker
Combined_sheet.Range("R2").Value = Max_Increase
Combined_sheet.Range("R2").NumberFormat = "0.00%"

Combined_sheet.Range("Q3").Value = Min_Ticker
Combined_sheet.Range("R3").Value = Max_Decrease
Combined_sheet.Range("R3").NumberFormat = "0.00%"

Combined_sheet.Range("Q4").Value = Max_TickerVolume
Combined_sheet.Range("R4").Value = Max_TotalVolume

End Sub




