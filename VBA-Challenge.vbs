Sub StockMarket()

'declare variables'
Dim ticker As String
Dim openprice As Double
Dim closingprice As Double
Dim totalstockvolume As Double
Dim summarytableindex As Double
Dim percentincrease, percentdecrease, greatesttotalvolumeticker As String
Dim percentincreasevalue, percentdecreasevalue, greatesttotalvolume As Double


For Each Ws In Worksheets
    'set initial conditions'
    totalstockvolume = 0
    summarytableindex = 2
    percentincrease = 0
    percentdecrease = 0
    greatesttotalvolume = 0
    openprice = Ws.Cells(2, "C").Value
    
Ws.Cells(1, "I").Value = "Ticker"
Ws.Cells(1, "J").Value = "Yearly Change"
Ws.Cells(1, "K").Value = "Percentage Change"
Ws.Cells(1, "L").Value = "Total Stock Volume"
Ws.Cells(1, "P").Value = "Ticker"
Ws.Cells(1, "Q").Value = "Value"
Ws.Cells(2, "O").Value = "Greatest Percent Increase"
Ws.Cells(3, "O").Value = "Greatest Percent Decrease"
Ws.Cells(4, "O").Value = "Greatest Total Volume"
    
For i = 2 To Ws.Cells(Rows.Count, 1).End(xlUp).Row
totalstockvolume = totalstockvolume + Ws.Cells(i, 7).Value
ticker = Ws.Cells(i, 1).Value

If Ws.Cells(i + 1, 1).Value <> Ws.Cells(i, 1).Value Then
Ws.Cells(summarytableindex, "I").Value = ticker
closingprice = Ws.Cells(i, "F").Value
Ws.Cells(summarytableindex, "J").Value = closingprice - openprice
If openprice <> 0 Then
Ws.Cells(summarytableindex, "K").Value = FormatPercent((closingprice - openprice) / openprice, 2)
Else
Ws.Cells(summarytableindex, "K").Value = Null
End If

Ws.Cells(summarytableindex, "L").Value = totalstockvolume

If Ws.Cells(summarytableindex, "J").Value > 0 Then
Ws.Cells(summarytableindex, "J").Interior.ColorIndex = 4
Else
Ws.Cells(summarytableindex, "J").Interior.ColorIndex = 3
End If

If Ws.Cells(summarytableindex, "K").Value > percentincrease Then
percentincrease = Ws.Cells(summarytableindex, "K").Value
percentincreasetotal = Ws.Cells(summarytableindex, "I").Value
End If

If Ws.Cells(summarytableindex, "K").Value < percentdecrease Then
percentdecrease = Ws.Cells(summarytableindex, "K").Value
percentdecreasetotal = Ws.Cells(summarytableindex, "I").Value
End If

If Ws.Cells(summarytableindex, "L").Value > greatesttotalvolume Then
greatesttotalvolume = Ws.Cells(summarytableindex, "L").Value
greatesttotalvolumeticker = Ws.Cells(summarytableindex, "I").Value
End If

openprice = Ws.Cells(i + 1, "C").Value
totalstockvolume = 0
summarytableindex = summarytableindex + 1
End If
Next i
 
Ws.Cells(2, "P").Value = percentincreasetotal
  Ws.Cells(2, "Q").Value = (FormatPercent(percentincrease, 2))
 
  Ws.Cells(3, "P").Value = percentdecreasetotal
    Ws.Cells(3, "Q").Value = (FormatPercent(percentdecrease, 2))
    
     Ws.Cells(4, "P").Value = greatesttotalvolumeticker
    Ws.Cells(4, "Q").Value = greatesttotalvolume

 Next Ws

End Sub
