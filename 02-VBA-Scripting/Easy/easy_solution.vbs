Sub Combine()
' Variable J is used in the For loop for moving to different sheets present on the Excel
    Dim J As Integer
    
'Variable TotStockVolume it used to keep track of the volume for the Stock ticker
    Dim TotStockVolume As Double

'Variable ReptPos (Report Position) is used to write the output data in the right place
    Dim ReptPos As Integer
    
'Variable LastRow is used for finding the number of rows in the sheet to avoid processing rows without data
    Dim LastRow As Long
    
    For J = 1 To Sheets.Count
        'Initializing TotStockVolume to 0 and this variable  will be used to 
        'store the total volume for a stock ticker
        TotStockVolume = 0

        
        'Initializing ReptPos (Report Position) to 2 (row 2), as the data is reported from row 2.
        ReptPos = 2

        'Update the Cell1(1,10) and (1,11) with values "Ticker" and "Total Stock Volume" (Header)
        'We will be reporting the Ticker and the corresponding Total Volume below these Values
        Sheets(J).Cells(1, 10) = "Ticker"
        Sheets(J).Cells(1, 11) = "Total Stock Volume"

        'Find the last row in the sheet and assign it to LastRow variable
        LastRow = Sheets(J).Cells(Rows.Count, "A").End(xlUp).Row
        
        'Iterate from Row 2 to LastRow (until the end of the sheet)
        For i = 2 To LastRow
            'If current cell and the next cell are same then add the Stock Volume and
            'go to the next row
            If Sheets(J).Cells(i, 1) = Sheets(J).Cells(i + 1, 1) Then
                TotStockVolume = TotStockVolume + Sheets(J).Cells(i, 7)
            Else
                'If current cell and the next cell are not same then the current cell is the
                'end of the data for the stock ticker, udpate the stock volume and report the data and 
                'go to the next row. 
                'Report the Stock Volume and the Ticker for the Ticker in current cell
                TotStockVolume = TotStockVolume + Sheets(J).Cells(i, 7)
                Sheets(J).Cells(ReptPos, 10) = Sheets(J).Cells(i, 1)
                Sheets(J).Cells(ReptPos, 11) = TotStockVolume
                TotStockVolume = 0
                ReptPos = ReptPos + 1
            End If
        Next i
    Next
End Sub
