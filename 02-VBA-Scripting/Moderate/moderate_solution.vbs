Sub Moderate_Stock_Data()
    ' Variable J is used in the For loop for moving to different sheets present on the Excel
    Dim J As Integer
    
    'Variable TotStockVolume it used to keep track of the volume for the Stock ticker
    Dim TotStockVolume As Double
    
    'Variable ReptPos (Report Position) is used to write the output data in the right place
    Dim ReptPos As Integer
    Dim Init_StockPrice As Double
    Dim Close_StockPrice As Double
    Dim Percent_Change As Double
    Dim LastRow As Long
    Dim FirstRow As Boolean
    FirstRow = True
    TotStockVolume = 0
    ReptPos = 2
    
    For J = 1 To Sheets.Count
        Sheets(J).Activate
        TotStockVolume = 0
        Sheets(J).Cells(1, 10) = "Ticker"
        Sheets(J).Cells(1, 11) = "Yearly Change"
        Sheets(J).Cells(1, 12) = "Percent Change"
        Sheets(J).Cells(1, 13) = "Total Stock Volume"
        
        ReptPos = 2
        LastRow = Sheets(J).Cells(Rows.Count, "A").End(xlUp).Row

        Init_Stock_Price = Sheets(J).Cells(2, 6)
        For i = 2 To LastRow
            If Sheets(J).Cells(i, 1) = Sheets(J).Cells(i + 1, 1) Then
                TotStockVolume = TotStockVolume + Sheets(J).Cells(i, 7)
                If FirstRow = True Then
                  Init_Stock_Price = Sheets(J).Cells(i, 6)
                  FirstRow = False
                End If
            Else
                TotStockVolume = TotStockVolume + Sheets(J).Cells(i, 7)
                Sheets(J).Cells(ReptPos, 10) = Sheets(J).Cells(i, 1)
                Sheets(J).Cells(ReptPos, 13) = TotStockVolume
                Close_Stock_Price = Sheets(J).Cells(i, 6)
                Percent_Change = Close_Stock_Price / Init_Stock_Price
                Sheets(J).Cells(ReptPos, 11) = Close_Stock_Price - Init_Stock_Price
                Sheets(J).Cells(ReptPos, 12) = (Close_Stock_Price - Init_Stock_Price) / Init_Stock_Price
                TotStockVolume = 0
                ReptPos = ReptPos + 1
                Init_Stock_Price = 0
                FirstRow = True
            End If
        Next i
    Next
End Sub
