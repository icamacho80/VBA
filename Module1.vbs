Attribute VB_Name = "Module1"
Sub dataanalysis()
    
 ' Set dimensions
    
    Dim total As Double
    Dim i As Long
    Dim change As Double
    Dim j As Integer
    Dim start As Long
    Dim rowCount As Long
    Dim percentChange As Double
    Dim days As Integer
    Dim dailyChange As Double
    Dim averageChange As Double
    Dim ws As Worksheet
    
For Each ws In ThisWorkbook.Worksheets
  
  ws.Activate
  
' Set title row
    
   Range("I1").Value = "Ticker"
   Range("J1").Value = "Quarterly Change"
   Range("K1").Value = "Percent Change"
   Range("L1").Value = "Total Stock Volume"
   
' Set initial values
    j = 0
    total = 0
    change = 0
    start = 2
    
' get the row number of the last row with data
    rowCount = Application.CountA(ActiveSheet.Range("A:A"))
     
    
    For i = 2 To rowCount
    
' If ticker changes then print results
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
' Stores results in variables
            total = total + Cells(i, 7).Value
            
' Handle zero total volume
            If total = 0 Then

' print the results
                Range("I" & 2 + j).Value = Cells(i, 1).Value
                Range("J" & 2 + j).Value = 0
                Range("K" & 2 + j).Value = "%" & 0
                Range("L" & 2 + j).Value = 0
                
                Else
                
' Find First non zero starting value
                If Cells(start, 3) = 0 Then
                    For find_value = start To i
                        If Cells(find_value, 3).Value <> 0 Then
                            start = find_value
                            Exit For
                        End If
                     Next find_value
                End If
                
      
' Calculate Change

               percentChange = Cells(i, 6).Value / Cells(start, 3).Value - 1
               dailyChange = Cells(i, 6).Value - Cells(start, 3).Value
                         
             
' start of the next stock ticker
                start = i + 1
                
' print the results
                
                Range("I" & 2 + j).Value = Cells(i, 1).Value
                Range("J" & 2 + j).Value = dailyChange
                Range("K" & 2 + j).Value = percentChange
                Range("L" & 2 + j).Value = total
                

' colors positives green and negatives red
                
                If Range("J" & 2 + j).Value > 0 Then
                    
                    Range("J" & 2 + j).Interior.ColorIndex = 4
                Else: Range("J" & 2 + j).Interior.ColorIndex = 3
                End If
              
End If
                
' reset variables for new stock ticker
            
            total = 0
            change = 0
            j = j + 1
            days = 0
            
' If ticker is still the same add results

        Else
            total = total + Cells(i, 7).Value
            


End If
Next i

    Columns("J:J").Select
    Selection.Style = "Comma"
    Columns("K:K").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"
    Columns("L:L").Select
    Selection.Style = "Comma"
    Selection.NumberFormat = "_(* #,##0.0_);_(* (#,##0.0);_(* ""-""??_);_(@_)"
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    Columns("L:L").EntireColumn.AutoFit
    Columns("K:K").EntireColumn.AutoFit
    Columns("J:J").EntireColumn.AutoFit


   Range("O2").Value = "Greatest % Increase"
   Range("O3").Value = "Greatest % Decrease"
   Range("O4").Value = "Greatest Total Volume"
   Range("P1").Value = "Ticker"
   Range("Q1").Value = "Value"
   
   Range("Q2").Value = "=MAX(K:K)"
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"
   Range("Q3").Value = "=MIN(K:K)"
   Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"
   Range("Q4").Value = "=MAX(L:L)"
   Selection.Style = "Comma"
    Selection.NumberFormat = "_(* #,##0.0_);_(* (#,##0.0);_(* ""-""??_);_(@_)"
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"

   Range("Q2").Value = "=MAX(K:K)"
   Range("Q3").Value = "=MIN(K:K)"
   Range("Q4").Value = "=MAX(L:L)"


   Range("P2").Value = "=XLOOKUP(Q2,K:K,I:I,0)"
   Range("P3").Value = "=XLOOKUP(Q3,K:K,I:I,0)"
   Range("P4").Value = "=XLOOKUP(Q4,L:L,I:I,0)"


Next ws
   
End Sub


