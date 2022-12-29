VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Sub Stock_market()


Dim ws As Worksheet


For Each ws In Worksheets



ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"


Dim Ticker As String
Ticker = " "
Dim Ticker_volume As Double
Ticker_volume = 0

Dim Greatest_Increase_percent As Double
Greatest_Increase_percent = 0
Dim Greatest_Increase_Ticker As String

Dim Greatest_Decrease_percent As Double
Greatest_Decrease_percent = 0
Dim Greatest_Decrease_Ticker As String

Dim Greatest_Total_Volume As Double
Greatest_Total_Volume = 0
Dim Greatest_Volume_Ticker As String


Dim Lastrow As Long
Dim i As Long

Dim firsRow As Boolean: firsRow = True


Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row


Dim open_price As Double
open_price = 0
Dim close_price As Double
close_price = 0
Dim price_change As Double
price_change = 0
Dim price_change_percent As Double
price_change_percent = 0

Dim TickerRow As Long: TickerRow = 1


For i = 2 To Lastrow


If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    TickerRow = TickerRow + 1
    Ticker = ws.Cells(i, 1).Value
    ws.Cells(TickerRow, "I").Value = Ticker

    

    close_price = ws.Cells(i, 6).Value
    ws.Cells(TickerRow, "J").Value = close_price - open_price
    
    ws.Cells(TickerRow, "J").Interior.ColorIndex = IIf(ws.Cells(TickerRow, "J").Value > 0, 4, 3)
    
    price_change_percent = ((close_price - open_price) / open_price)
    ws.Cells(TickerRow, "K").Value = FormatPercent(price_change_percent)
    ws.Cells(TickerRow, "k").Interior.ColorIndex = IIf(ws.Cells(TickerRow, "k").Value > 0, 4, 3)
    
    
    If price_change_percent > 0 And price_change_percent > Greatest_Increase_percent Then
        Greatest_Increase_percent = (price_change_percent)
        Greatest_Increase_Ticker = ws.Cells(i, 1).Value
    End If
    
     If price_change_percent < 0 And price_change_percent < Greatest_Decrease_percent Then
        Greatest_Decrease_percent = (price_change_percent)
        Greatest_Decrease_Ticker = ws.Cells(i, 1).Value
    End If
    
  
    
    Ticker_volume = Ticker_volume + ws.Cells(i, 7).Value
    
    If Ticker_volume > Greatest_Total_Volume Then
        Greatest_Total_Volume = Ticker_volume
        Greatest_Volume_Ticker = ws.Cells(i, 1).Value
    End If
    ws.Cells(TickerRow, "L").Value = Ticker_volume
    Ticker_volume = 0
    firsRow = True
    
Else
    
    Ticker_volume = Ticker_volume + ws.Cells(i, 7).Value
    If firsRow = True Then
        open_price = ws.Cells(i, 3).Value
        firsRow = False
    End If

End If

Next i

ws.Cells(2, "Q").Value = FormatPercent(Greatest_Increase_percent)
ws.Cells(2, "P").Value = Greatest_Increase_Ticker

ws.Cells(3, "Q").Value = FormatPercent(Greatest_Decrease_percent)
ws.Cells(3, "P").Value = Greatest_Decrease_Ticker

ws.Cells(4, "Q").Value = Greatest_Total_Volume
ws.Cells(4, "P").Value = Greatest_Volume_Ticker




Next ws

End Sub
