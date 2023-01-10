Attribute VB_Name = "Module1"
Sub AlphabeticalTesting()

    Dim ticker As String

    Dim PriceChange As String
       
    Dim YearOpen As Double
    Dim YearClose As Double
    
    Dim PercentChange As Double

    Dim StockVolume As Double

    J = 0
    StockVolume = 0
    
    ' print header row
    Range("I1").Value = "Ticker"
    Range("J1").Value = "YearlyChange"
    Range("K1").Value = "PercentChange"
    Range("L1").Value = "StockVolume"
    
    YearOpen = Cells(2, 3).Value
  
  
  ' Loop through tickers
    lastrow = Cells(Rows.Count, "A").End(xlUp).Row
  For I = 2 To lastrow
   
    If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
     StockVolume = StockVolume + Cells(I, 7).Value
    
    YearClose = Cells(I, 6).Value
    
       Cells(2 + J, 9).Value = Cells(I, 1).Value
       
     ' Calculation for price change
        Cells(2 + J, 10) = YearClose - YearOpen
       
      'Calculation for Percent change
          Cells(2 + J, 11) = ((YearClose - YearOpen) / YearOpen) - 1
    
        'Calculation for Stock Volume
       Cells(2 + J, 12).Value = StockVolume
       
       
       
    StockVolume = 0
    
     J = J + 1
     
     Else
     StockVolume = StockVolume + Cells(I, 7).Value
     
    End If
    
    If Cells(I - 1, 1).Value <> Cells(I, 1).Value Then
    
        YearOpen = Cells(I, 3).Value
    
    End If
    
  Next I

End Sub

