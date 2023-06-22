Attribute VB_Name = "Module1"
Sub Stock()

Dim ws As Worksheet
'run a loop over all sheets at onces
For Each ws In worksheets
 
'Doing some extra job
Dim Ticker As String
Range("K1").Value = "Ticker"
Dim Yearly_Change As String
Range("L1").Value = "Yearly_Change"
Dim Percent_Change As String
Range("M1").Value = "Percent_Change"
Dim Total_volume As String
Range("N1").Value = "Total_volume"

 Range("Q2").Value = "Greatest % Increase"
 Range("Q3").Value = "Greatest % Decrease"
 Range("Q4").Value = "Greatest Total Volume"
 Range("R1").Value = "Ticker"
 Range("S1").Value = "value"

'creating a variable to calculat the last active row for the loop
Dim LR As Double
LR = worksheets("2018").UsedRange.Rows.Count

'create a variable to be able to go down in the new table
Dim table_row As Integer
table_row = 2

'create a variable to be abel to sum all Volume for each ticker
Dim T As Double
T = 0
Dim A As Double
A = 0
'create a Y value for yearly change
'creat a X value to hold the cell for the first open charge
Dim Y As Double
Dim X As Double
X = 2
'creat a percentage value
Dim percent As Double

 Dim Increase As Double
  Dim Decrease As Double
  Dim Total_V As Double
  Increase = Cells(2, 13).Value
  Total_V = Cells(2, 14).Value
  Decrease = Cells(2, 13).Value
  
  For i = 2 To LR
    
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        'place each ticker in  the "K" colume
        Cells(table_row, 11).Value = Cells(i, 1).Value
        
        'calculat the total volume for each ticker and palce volume values in the "N" colume
         T = T + Cells(i, 7).Value
        Range("N" & table_row).Value = T
        
       'calculat the Yearly-change and palce it in the "L" colune
         Y = Cells(i, 6).Value - Cells(X, 3).Value
          Range("L" & table_row).Value = Y
          
          'calculat percent_change and palce it in the "M" colume
          percent = (Y / Cells(X, 3).Value)
          Range("M" & table_row).Value = Format(percent, "Percent")
          
          'add a row to palce the second ticker info
          table_row = table_row + 1
          
        'zero out all values to Re-start with another ticker
        T = 0
        Y = 0
        X = i + 1
    Else
   T = T + Cells(i, 7).Value
   
  End If
  Next i
  
  'find the count of the row
  Dim PR As Double
  PR = Cells(Rows.Count, 13).End(xlUp).Row
  'run a loop on the "M" colume
  For j = 2 To PR
  
  'color each cell in the "M" colume according to its value
    If Cells(j, 13).Value > 0 Then
        Cells(j, 13).Interior.ColorIndex = 4
        Cells(j, 13).Font.ColorIndex = 1
    Else
    Cells(j, 13).Interior.ColorIndex = 3
    Cells(j, 13).Font.ColorIndex = 1
    End If
    
  'condition to find out the highest percentage
  If Cells(j, 13).Value > Increase Then
        Increase = Cells(j, 13).Value
        Cells(2, 18).Value = Cells(j, 11).Value
        Cells(2, 19).Value = Format(Increase, "Percent")
        Else
         Increase = Increase
        End If
   'condition to find out the lowest percentage
    If Cells(j, 13).Value < Decrease Then
        Decrease = Cells(j, 13).Value
        Cells(3, 18).Value = Cells(j, 11).Value
        Cells(3, 19).Value = Format(Decrease, "Percent")
   Else
       Decrease = Decrease
        
    End If
    'condition to find out the total volume
    If Cells(j, 14).Value > Total_V Then
        Total_V = Cells(j, 14).Value
         Cells(4, 18).Value = Cells(j, 11).Value
        Cells(4, 19).Value = Total_V
    Else
        Total_V = Total_V
        
        End If
  
  Next j
  
Next ws

End Sub

