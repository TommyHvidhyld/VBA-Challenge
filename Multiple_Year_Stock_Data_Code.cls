VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Alphabetical_Testing()
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim i As Long
    Dim Ticker As String
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Stock_Volume As Double
    Dim Summary_Table_Row As Integer
    Dim BegYear As Double
    Dim EndYear As Double
    Dim Greatest_Increase As Double
    Dim Greatest_Decrease As Double
    Dim Greatest_Volume As Double
    
    For Each Sheet In Application.Worksheets
        Sheet.Activate
        
        Set ws = Application.ActiveSheet
        
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        
        Range("L1").ColumnWidth = 15
        Range("K1").ColumnWidth = 12
        Range("J1").ColumnWidth = 11
        Range("P1").ColumnWidth = 15
        Range("N1").ColumnWidth = 17
        
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        Summary_Table_Row = 2
        Stock_Volume = 0
        BegYear = Cells(2, 3).Value
        For i = 2 To LastRow
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                Ticker = Cells(i, 1).Value
                    Range("I" & Summary_Table_Row).Value = Ticker
                EndYear = Cells(i, 6).Value
                Yearly_Change = EndYear - BegYear
                    Range("J" & Summary_Table_Row).Value = Yearly_Change
                Percent_Change = (Yearly_Change / BegYear)
                    Range("K" & Summary_Table_Row).Value = Percent_Change
                    Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                Stock_Volume = Stock_Volume + Cells(i, 7).Value
                    Range("L" & Summary_Table_Row).Value = Stock_Volume
                Summary_Table_Row = Summary_Table_Row + 1
                Stock_Volume = 0
                BegYear = Cells(i + 1, 3).Value
            Else
                Stock_Volume = Stock_Volume + Cells(i, 7).Value
            End If
        Next i
        
        For i = 2 To LastRow
            If Cells(i, 10).Value > 0 Then
                    Cells(i, 10).Interior.ColorIndex = 4
            ElseIf Cells(i, 10).Value < 0 Then
                    Cells(i, 10).Interior.ColorIndex = 3
            End If
        Next i
        
        Range("O1") = "Ticker"
        Range("P1") = "Value"
        Range("N2") = "Greatest % Increase"
        Range("N3") = "Greatest % Decrease"
        Range("N4") = "Greatest Total Volume"
        
         Greatest_Increase = Application.WorksheetFunction.Max(Range("K2:K" & LastRow).Value)
            Cells(2, 16).Value = Greatest_Increase
            Cells(2, 16).NumberFormat = "0.00%"
        Greatest_Decrease = Application.WorksheetFunction.Min(Range("K2:K" & LastRow).Value)
            Cells(3, 16).Value = Greatest_Decrease
            Cells(3, 16).NumberFormat = "0.00%"
        Greatest_Volume = Application.WorksheetFunction.Max(Range("L2:L" & LastRow).Value)
            Cells(4, 16).Value = Greatest_Volume
        
        Range("O2") = WorksheetFunction.XLookup(Greatest_Increase, [K:K], [I:I], 0)
        Range("O3") = WorksheetFunction.XLookup(Greatest_Decrease, [K:K], [I:I], 0)
        Range("O4") = WorksheetFunction.XLookup(Greatest_Volume, [L:L], [I:I], 0)
    Next
    
End Sub
