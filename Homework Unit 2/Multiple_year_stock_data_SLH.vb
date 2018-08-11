Sub StockAnalysis()

' This subroutine executes both of the subroutines below for all of the worksheets in the workbook

Dim ws As Worksheet

 For Each ws In ThisWorkbook.Worksheets
     Call TotalStockVolume(ws)
 Next


End Sub




Sub TotalStockVolume(ws As Worksheet)

' Loop through data and summarize total volume by ticker symbol

Dim TotalStockVolume As Double
Dim LastRow As Long
Dim TickerSymbol As String
Dim TickerStockVolume As Double
Dim OpeningPrice As Double
Dim ClosingPrice As Double
Dim YearlyChange As Double
Dim PercentChange As Double



With ws


    ' Insert Headers for the Summary Area

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 13).Value = "Opening Price"
    ws.Cells(1, 14).Value = "Closing Price"

    ' Determine the last row in the worksheet

    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row


    ' Initialize the summary area row number and stock total volume

    VolumeSummaryRow = 2
    TickerStockVolume = 0

    OpeningPrice = .Range("C2")
    ClosingPrice = .Range("F2")

    For I = 2 To LastRow
   
      If .Cells(I + 1, 1).Value <> .Cells(I, 1).Value Then
    
           
          ' Put the ticker symbol in the summary area
           TickerSymbol = .Cells(I, 1).Value
           .Cells(VolumeSummaryRow, 9).Value = TickerSymbol
        
         ' Put the ticker volume in the summary area
       
           TickerStockVolume = TickerStockVolume + .Cells(I, 7).Value
           .Cells(VolumeSummaryRow, 12).Value = TickerStockVolume
         
         ' Put the Opening Price in the summary area
        
           .Cells(VolumeSummaryRow, 13).Value = OpeningPrice
             
         ' Put the Closing Price in the summary area
        
           ClosingPrice = .Range("F" & I)
          .Cells(VolumeSummaryRow, 14).Value = ClosingPrice
                 
         ' Calculate the Yearly and Percentage Change from Year Opening to Year Closing
       
           YearlyChange = Round(ClosingPrice - OpeningPrice, 4)
           .Cells(VolumeSummaryRow, 10).Value = YearlyChange
           .Cells(VolumeSummaryRow, 10).NumberFormat = "0.00"
         
         ' If the yearly price change is negative then color the cell red, otherwise green
         
           If YearlyChange < 0 Then
         
              .Cells(VolumeSummaryRow, 10).Interior.ColorIndex = 3
           Else
         
              .Cells(VolumeSummaryRow, 10).Interior.ColorIndex = 4
         
           End If
         
           If OpeningPrice <> 0 Then
              PercentChange = Round(YearlyChange / OpeningPrice, 4)
           Else
              PercentChange = 0
           End If
             
           .Cells(VolumeSummaryRow, 11).Value = PercentChange
           .Cells(VolumeSummaryRow, 11).NumberFormat = "0.00%"
         
                  
          ' Increment the summary area by one
        
           VolumeSummaryRow = VolumeSummaryRow + 1
        
          ' Reset the ticker symbol total volume back to 0
        
           TickerStockVolume = 0
        
          ' Get the new row number for the new symbol
        
           NewRowNumber = I + 1
        
          ' Store the Opening Price for the new symbol
        
           OpeningPrice = .Range("C" & I + 1)
        
      Else
           TickerStockVolume = TickerStockVolume + .Cells(I, 7).Value
           ClosingPrice = .Range("C" & I)
        
      End If
    
      Next I

 ' AutoFit the Column widths
 
 ws.Select
 .Columns("J:N").Select
 Selection.EntireColumn.AutoFit


 ' Call the subroutine that prints out the high, low, total volume numbers
 
 Call TotalGrid(ws)
 
 ' Zoom out on the worksheet to make it easier to read
 
 ActiveWindow.Zoom = 85
 
End With


End Sub


Sub TotalGrid(ws)


' Get the Greatest % increase, Greatest % Decrease and Greatest total volume.

Dim TopIncrease As Double
Dim TopDecrease As Double
Dim TopVolume As Double
Dim TopTickerIncrease As String
Dim TopTickerDecrease As String
Dim TopTickerVolume As String
Dim TotalGridRange As Range
Dim TotalGridRows As Integer
Dim PercentChangeRange As Range
Dim LastGridRow As Long

LastGridRow = ws.Cells(Rows.Count, 9).End(xlUp).Row

' Insert the cell headers

ws.Range("P2").Value = "Greatest % Increase"
ws.Range("P3").Value = "Greatest % Decrease"
ws.Range("P4").Value = "Greatest Total Volume"
ws.Range("Q1").Value = "Ticker"
ws.Range("R1").Value = "Value"

' Find and format the grid values

TopIncrease = Application.WorksheetFunction.Max(ws.Range("K:K"))
MinIncrease = Application.WorksheetFunction.Min(ws.Range("K:K"))
TopVolume = Application.WorksheetFunction.Max(ws.Range("L:L"))

ws.Range("R2").Value = TopIncrease
ws.Range("R3").Value = MinIncrease
ws.Range("R2").NumberFormat = "0.00%"
ws.Range("R3").NumberFormat = "0.00%"

ws.Range("R4").Value = TopVolume
ws.Range("R4").NumberFormat = "#,##0"

' Get the ticker symbols associated with the maximum increase, decrease and volume

 
TopTickerIncrease = Application.WorksheetFunction.Index(ws.Columns(9), Application.WorksheetFunction.Match(Application.WorksheetFunction.Max(ws.Range("K:K")), ws.Range("K:K"), False))
TopTickerDecrease = Application.WorksheetFunction.Index(ws.Columns(9), Application.WorksheetFunction.Match(Application.WorksheetFunction.Min(ws.Range("K:K")), ws.Range("K:K"), False))
TopTickerVolume = Application.WorksheetFunction.Index(ws.Columns(9), Application.WorksheetFunction.Match(Application.WorksheetFunction.Max(ws.Range("L:L")), ws.Range("L:L"), False))

ws.Range("Q2").Value = TopTickerIncrease
ws.Range("Q3").Value = TopTickerDecrease
ws.Range("Q4").Value = TopTickerVolume


Set PercentChangeRange = ws.Range("I2:J" & LastGridRow)

' TopTickerIncrease = Application.WorksheetFunction.VLookup(TopIncrease, PercentChangeRange, 2, False)
' Range("Q2").Value = TopTickerIncrease

'  TopTickerIncrease = Application.WorksheetFunction.VLookup(TopIncrease, PercentChangeRange, 2, False)
'  Range("Q2").Value = TopTickerIncrease


' ws.Select
' ws.Columns("P:R").Select
' Selection.EntireColumn.AutoFit

ws.Range("P:R").Columns.AutoFit


End Sub










