## VBA_Scripting
# Analysing stock Data for the years 2018, 2019, 2020
## The file contains data regarding the stock of different companies marked by ticker values  over the period of three years (2018, 2019, 2020).

### A VBA script that loops through all the stocks for each year and outputs information including
- ### Yearly change
##### Change from the opening price at the beginning of a given year to the closing price at the end of that year.
- ### Percentage change
##### Change in percentage from the opening price at the beginning of a given year to the closing price at the end of that year.
- ### The total stock volume
##### Sum of all the different stock throughout the year.
### The stock with the following has also been identified
- #### Greatest % increase, 
- #### Greatest % decrease, and 
- #### Greatest total volume 
        .



Sub ticker():
 Dim ws As Integer
  ws = Application.Worksheets.Count
 
 For a = 1 To ws
 
 Worksheets(a).Activate
    
    Dim lastrow As Long
        lastrow = Cells(Rows.Count, 1).End(xlUp).row
        
            'MsgBox (lastrow)
            
    Range("I1").Value = "Ticker"
    Range("j1").Value = "Yearly Change"
    Range("k1").Value = "Percent Change"
    Range("l1").Value = "Total Stock Volume"

Dim row As Integer

    row = 2
   
Dim ticker As String

Dim i As Long


    For i = 2 To lastrow

        ticker = Cells(i, 1).Value
      
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

Range("I" & row).Value = ticker

row = row + 1

Else

Range("I" & row).Value = ticker

End If

Next i

Next a

End Sub

Sub totalvolumecalc():
 Dim ws As Integer
  ws = Application.Worksheets.Count
 
 For a = 1 To ws
 
 Worksheets(a).Activate
 
    Range("I1").Value = "Ticker"
    Range("j1").Value = "Yearly Change"
    Range("k1").Value = "Percent Change"
    Range("l1").Value = "Total Stock Volume"

Dim row As Integer

    row = 2
   
Dim ticker As String
Dim lastrow As Long
Dim yearlychange As Double
Dim percentcahnage As Double
Dim totalstockvol As Double

lastrow = Cells(Rows.Count, 1).End(xlUp).row
'MsgBox (lastrow)
totalstockvol = 0

Dim i As Long

    For i = 2 To lastrow

        ticker = Cells(i, 1).Value
      
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

Range("I" & row).Value = ticker

totalstockvol = totalstockvol + Cells(i, 7).Value

Range("L" & row).Value = totalstockvol

row = row + 1
totalstockvol = 0

Else

totalstockvol = totalstockvol + Cells(i, 7).Value

End If

Next i

Next a

End Sub

Sub yearlychangess():
    Dim ws As Integer
  ws = Application.Worksheets.Count
 
 For a = 1 To ws
 
 Worksheets(a).Activate
 
Dim row As Long
Dim beginrow As Long
Dim lastrow As Long
Dim lastrowb As Long
Dim yearlychange As Double
Dim openval As Double
Dim closeval As Double
Dim initialval As Double
Dim percentchange As Double

row = 2
beginrow = 2
openval = 0
closeval = 0
yearlychange = 0
percentchange = 0
lastrow = Cells(Rows.Count, 1).End(xlUp).row
'MsgBox (lastrow)
lastrowb = Cells(Rows.Count, 9).End(xlUp).row

Dim i As Long

For i = 2 To lastrow
    
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

    openval = Cells(row, 3).Value
    closeval = Cells(i, 6).Value
    
    yearlychange = closeval - openval
    
    Range("j" & row).Value = yearlychange
    
row = row + 1
beginrow = i + 1

openval = 0
closeval = 0
yearlychange = 0
percentchange = 0

End If
    
If openval = 0 And closeval = 0 Then
    percentchange = yearlychange
ElseIf openval = 0 And closeval > 0 Then
    percentchange = 0
Else
    percentchange = yearlychange / openval
        
Range("k" & row).Value = percentchange
    
row = row + 1
beginrow = i + 1

openval = 0
closeval = 0
yearlychange = 0
percentchange = 0

End If
    
For j = 2 To lastrowb
         
If Cells(j, 10).Value > 0 And Cells(j, 11).Value > 0 Then
    Cells(j, 10).Interior.ColorIndex = 4
    Cells(j, 11).Interior.ColorIndex = 4
ElseIf Cells(j, 10).Value < 0 And Cells(j, 11).Value < 0 Then
    Cells(j, 10).Interior.ColorIndex = 3
    Cells(j, 11).Interior.ColorIndex = 3
Else
    Cells(j, 10).Interior.ColorIndex = 0
    Cells(j, 11).Interior.ColorIndex = 0
End If

Next j
    
Next i

Next a
    
End Sub

    


Sub Stock():

Dim ws As Integer
  ws = Application.Worksheets.Count
 
 For a = 1 To ws
 
 Worksheets(a).Activate

Dim maxpercentticker As String
Dim minpercentticker As String
Dim maxvolticker As String
Dim maxpercentvalue As Double
Dim minpercentvalue As Double
Dim maxstockvolume As Double
Dim row As Long

Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Range("O2").Value = "Maximum % Change"
Range("O3").Value = "Minimum % Change"
Range("O4").Value = "Maximum Stock Volume"

maxpercentticker = Cells(2, 11).Value
minpercentticker = Cells(2, 11).Value
maxvolumeticker = Cells(2, 11).Value
maxpercentvalue = Cells(2, 13).Value
minpercentvalue = Cells(2, 13).Value
maxstockvolume = Cells(2, 14).Value
lastrow = Cells(Rows.Count, 11).End(xlUp).row


For k = 2 To lastrow
  
        
If Cells(k, 12) > maxpercentvalue Then
    maxpercentvalue = Cells(k, 13).Value
    maxpercentticker = Cells(k, 11).Value

ElseIf Cells(k, 12) < minpercentvalue Then
    minpercentvalue = Cells(k, 13).Value
    minpercentticker = Cells(k, 11).Value
End If

If Cells(k, 13) > maxstockvolume Then
    maxstockvolume = Cells(k, 14).Value
    maxvolumeticker = Cells(k, 11).Value
        

End If
        
    Range("p2").Value = maxpercentticker
    Range("q2").Value = maxpercentvalue
    Range("q2").Value = FormatPercent(maxpercentvalue, 2)
    Range("p3").Value = minpercentticker
    Range("q3") = minpercentvalue
    Range("q3").Value = FormatPercent(minpercentvalue, 2)
    Range("p4") = maxvolumeticker
    Range("q4") = maxstockvolume
    
Next k

Next a
    
End Sub


