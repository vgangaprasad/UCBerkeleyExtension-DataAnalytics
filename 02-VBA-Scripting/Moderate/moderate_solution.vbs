Sub Moderate_Stock_Data()
    ' Variable J is used in the For loop for moving to different sheets present on the Excel
    Dim J As Integer
    
    'Variable TotStockVolume it used to keep track of the volume for the Stock ticker
    Dim TotStockVolume As Double
    
    'Variable ReptPos (Report Position) is used to write the output data in the right place
    Dim ReptPos As Integer

    'Init_StockPrice variable is used to store the initial stock price and this will be used for 
    'calculating Percent Change as well as Yearly change
    Dim Init_StockPrice As Double
    
    
    'Close_StockPrice variable is used to store the closing stock price and this will be used for 
    'calculating Percent Change as well as Yearly change
    Dim Close_StockPrice As Double
    
    'Percent_Change variable will be used for calculating the percentage change and used for reporting also
    Dim Percent_Change As Double
    
    'Variable LastRow is used for finding the number of rows in the sheet to avoid processing rows without data
    Dim LastRow As Long
    
    'FirstRow is a Boolean variable used to find out the first row for a ticker and storing the 
    'init stock price, ticker information
    Dim FirstRow As Boolean
    FirstRow = True
    
    For J = 1 To Sheets.Count

        'Printing the Headers
        Sheets(J).Cells(1, 10) = "Ticker"
        Sheets(J).Cells(1, 11) = "Yearly Change"
        Sheets(J).Cells(1, 12) = "Percent Change"
        Sheets(J).Cells(1, 13) = "Total Stock Volume"
        
        'Initializing ReptPos & Rept2Pos (Report Position) to 2 (row 2), as the data is reported from row 2.
        ReptPos = 2

        'Initializing TotStockVolume to 0 and this variable  will be used to 
        'store the total volume for a stock ticker
        TotStockVolume = 0

        'Find the last row in the sheet and assign it to LastRow variable
        LastRow = Sheets(J).Cells(Rows.Count, "A").End(xlUp).Row

        'Assigining the initial Stock price for the first ticker on the sheet, for the next ones, it will
        'be assigned after reading the first row of the ticker.
        Init_Stock_Price = Sheets(J).Cells(2, 6)
        
        'Iterate from Row 2 to LastRow (until the end of the sheet)
        For i = 2 To LastRow

            'If current cell and the next cell are same then add the Stock Volume and
            'go to the next row
            If Sheets(J).Cells(i, 1) = Sheets(J).Cells(i + 1, 1) Then
                TotStockVolume = TotStockVolume + Sheets(J).Cells(i, 7)
                'This IF statement will be used for assigining the initial Stock price for the ticker, it will assign
                'the Init_Stock_Price value only for the first occurence of the stock ticker.
                If FirstRow = True Then
                  Init_Stock_Price = Sheets(J).Cells(i, 6)
                  FirstRow = False
                End If
            Else
                'If current cell and the next cell are not same then the current cell is the
                'end of the data for the stock ticker, udpate the stock volume and report the data and 
                'go to the next row. 
                'Report the Stock Volume and the Ticker for the Ticker in current cell
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
