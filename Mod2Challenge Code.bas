Attribute VB_Name = "Module1"
Sub Mod2Challenge()

'Create a script that loops through all the stocks for one year and outputs the following information:
'The ticker symbol
'Yearly change from the opening price to the closing price at the end of that year.
'The percentage change from the opening price to the closing price at the end of that year.
'The total stock volume of the stock.
'Return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".

'New Column Headers
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"

'Labels
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"


'Define Variables
Dim Tickername As String
Dim YearlyChange As Double
Dim PercentChange As Double
Dim TotalVolume As Double
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim LastRow As Long
Dim J As Integer
Dim i As Long
Dim NewOpenPrice As Double
Dim Maxinc As Double
Dim Maxdec As Double
Dim MaxVolume As Double
Dim MaxIncTick As String
Dim MaxDecTick As String
Dim MaxVolumeTick As String





TotalVolume = 0


LastRow = Cells(Rows.Count, 1).End(xlUp).Row
J = 2
OpenPrice = Cells(J, 3).Value
ClosePrice = Cells(J, 6).Value
YearlyChange = (ClosePrice - OpenPrice)
PercentChange = (YearlyChange / OpenPrice)
NewOpenPrice = Cells(2, 3).Value
Cells(2, 11).Value = Cells(2, 3).Value

For i = 2 To LastRow
TotalVolume = TotalVolume + Cells(i, 7)

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Cells(J, 9).Value = Cells(i, 1).Value
        Cells(J, 12).Value = TotalVolume
    TotalVolume = 0
    ClosePrice = Cells(i, 6).Value
    OpenPrice = Cells(i, 3).Value
    Cells(J + 1, 11).Value = NewOpenPrice
    Cells(J, 10).Value = ClosePrice
   Cells(J, 13).Value = Cells(J, 10).Value - Cells(J, 11).Value
   Cells(J, 14).Value = (Cells(J, 10).Value - Cells(J, 11).Value) / (Cells(J, 11).Value)
   Cells(J, 10).Value = Cells(J, 13).Value: Cells(J, 13).Value = ""
   Cells(J, 11).Value = Cells(J, 14).Value: Cells(J, 14).Value = ""
  
  'Formating
   Cells(J, 11).NumberFormat = "0.00%"
    If Cells(J, 10).Value > 0 Then
        Cells(J, 10).Interior.ColorIndex = 4
    ElseIf Cells(J, 10).Value < 0 Then
        Cells(J, 10).Interior.ColorIndex = 3
    End If
    
    

    J = J + 1
 
    Else: NewOpenPrice = Cells(i + 2, 3)
    End If
    

Next i
End Sub


