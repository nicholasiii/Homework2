Attribute VB_Name = "Module1"
Sub Dosomething()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call StocksHomework
    Next
    Application.ScreenUpdating = True
End Sub

Sub StocksHomework()


Dim LastRow, LastRow2 As Long
Dim sht As Worksheet
Dim GreatIncrease, Grestdecrease As Double
Dim Greatvolume As Double


Set sht = ActiveSheet

'Using Find Function (Provided by Bob Ulmas)
  LastRow = sht.Cells.Find("*", searchorder:=xlByRows, searchdirection:=xlPrevious).Row

  ' Set an initial variable for holding the ticker
  Dim Ticker As String

  ' Set an initial variable for holding Pricing Info and Stock Volume
 Dim OrigPrice, FinalPrice As Double
  Dim Stock_Volume As Double
  Stock_Volume = 0

  ' Keep track of the location for each company in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  ' Loop through all stock info
  For i = 2 To LastRow
    
    ' Check if we are still within the same ticker, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Get Ticker, Update price for Final Price
      Ticker = Cells(i, 1).Value
      FinalPrice = Cells(i, 6).Value
      ' Add to the Stock Volume Total
      Stock_Volume = Stock_Volume + Cells(i, 7).Value

      ' Print the Ticker in the Row J table, Store FinalPrice
      Range("j" & Summary_Table_Row).Value = Ticker
      Range("l" & Summary_Table_Row).Value = FinalPrice

      ' Print the Volume Info to the Summary Table
      Range("m" & Summary_Table_Row).Value = Stock_Volume

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Volume Total
      Stock_Volume = 0
    
    ' If the cell immediately following a row is the same brand...
    Else

      ' Add to the Volume Total
      Stock_Volume = Stock_Volume + Cells(i, 7).Value


    
    End If
    'Headers
    Cells(1, 10) = "Ticker"
    Cells(1, 11) = "Yearly Change"
    Cells(1, 12) = "Percent Change"
    Cells(1, 13) = "Stock Volume"
    Next i
    ' Starting here, we are gettting Original price data and putting it in the Row J-M table
    'Reset the Sumary_Table_Row
      Summary_Table_Row = 2
  For i = 2 To LastRow
    
    ' If statement puts OrigPrice in Column K for later math
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        OrigPrice = Cells(i + 1, 3).Value
        'If a stock started on the exchange this year it's initial value will be zero, this is intended to get a proper value
                If OrigPrice = 0 Then
                 For j = 1 To 365
                 If Cells(j + i + 1, 3) <> 0 Then
                    OrigPrice = Cells(j + i + 1, 3).Value
                    Exit For
                    End If
                Next j
                End If
                 
    ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      Range("k" & Summary_Table_Row).Value = OrigPrice

      End If
      Next
      'This algo does not get the first Original Price properly, so just copy it over
      Cells(2, 11) = Cells(2, 3)

'Gets the last row for the Column J-M table
LastRow2 = sht.Cells(sht.Rows.Count, "K").End(xlUp).Row
'Takes the Original Price and the Final Price and calculates gain in both percentage and absolute terms,
'it then does conditional formatting. It also checks for Divide by Zero conditions in case a stock entered the database with no data
  For i = 2 To LastRow2
    If Cells(i, 12) <> 0 Then
        OrigPrice = (Cells(i, 11))
        FinalPrice = Cells(i, 12)
        Cells(i, 11).Value = FinalPrice - OrigPrice
        Cells(i, 12) = (FinalPrice - OrigPrice) / OrigPrice
        Cells(i, 12).NumberFormat = "0.00%"
        If Cells(i, 11) > 0 Then
             Cells(i, 11).Interior.ColorIndex = 4
        Else
             Cells(i, 11).Interior.ColorIndex = 3
       End If
    Else
        Cells(i, 11) = "NULL"
        Cells(i, 12) = "NULL"
    End If
  Next i
  
  
  'Search L for greatest value (i,11), least value, and M (i,12) for greatest total volume, only runs on lines with non-NULL datas
Cells(2, 16) = "Greatest % Increase"
Cells(3, 16) = "Greatest % Decrease"
Cells(4, 16) = "Greatest Volume"
Cells(1, 17) = "Ticker"
Cells(1, 18) = "Value"
GreatIncrease = 0
Greatvolume = 0
Grestdecrease = 0

    For i = 2 To LastRow2
        If (Cells(i, 12) <> "NULL") Then
        If (Cells(i, 12)) > GreatIncrease Then
            GreatIncrease = Cells(i, 12).Value
            Cells(2, 18).Value = GreatIncrease
            Cells(2, 18).NumberFormat = "0.00%"
            Cells(2, 17).Value = Cells(i, 10)
        End If
        If (Cells(i, 12)) < Grestdecrease Then
            Grestdecrease = Cells(i, 12).Value
            Cells(3, 18).Value = Grestdecrease
            Cells(3, 18).NumberFormat = "0.00%"
            Cells(3, 17).Value = Cells(i, 10)
        End If

        If (Cells(i, 13)) > Greatvolume Then
            Greatvolume = Cells(i, 13).Value
            Cells(4, 18).Value = Greatvolume
            Cells(4, 17).Value = Cells(i, 10)
        End If

        End If
        Next i
  
End Sub


