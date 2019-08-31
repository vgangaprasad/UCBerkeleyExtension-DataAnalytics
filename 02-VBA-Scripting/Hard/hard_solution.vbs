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
    
    'Variable Rept2Pos (2nd Report Position) is used to write the output data in the right place
    Dim Rept2Pos As Integer

    'Following variables are used for keeping track of 
    'Great % increase/decrease and total volume values
    Dim Great_Per_In_Ticker As String
    Dim Great_Per_In_Value As Double
    Dim Great_Per_De_Ticker As String
    Dim Great_Per_De_Value As Double
    Dim Great_TotVol_Ticker As String
    Dim Great_TotVol_value As Double
    
    
    For J = 1 To Sheets.Count
        
        'Initializing the variables that are used to report the % Increase/Decrease and Total Volume
        Great_Per_In_Ticker = ""
        Great_Per_De_Ticker = ""
        Great_TotVol_Ticker = ""

        Great_Per_In_Value = 0
        Great_Per_De_Value = 0
        Great_TotVol_value = 0
        
        'Printing the Headers
        Sheets(J).Cells(1, 10) = "Ticker"
        Sheets(J).Cells(1, 11) = "Yearly Change"
        Sheets(J).Cells(1, 12) = "Percent Change"
        Sheets(J).Cells(1, 13) = "Total Stock Volume"
        
        Sheets(J).Cells(1, 16) = "Ticker"
        Sheets(J).Cells(1, 17) = "Value"
        
        'Initializing TotStockVolume to 0 and this variable  will be used to 
        'store the total volume for a stock ticker
        TotStockVolume = 0
        'Initializing ReptPos & Rept2Pos (Report Position) to 2 (row 2), as the data is reported from row 2.
        ReptPos = 2
        Rept2Pos = 2

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

                If Sheets(J).Cells(ReptPos, 12) > Great_Per_In_Value Then
                    Great_Per_In_Value = Sheets(J).Cells(ReptPos, 12)
                    Great_Per_In_Ticker = Sheets(J).Cells(i, 1)
                End If

                If Sheets(J).Cells(ReptPos, 12) < Great_Per_De_Value Then
                    Great_Per_De_Value = Sheets(J).Cells(ReptPos, 12)
                    Great_Per_De_Ticker = Sheets(J).Cells(i, 1)
                End If

                If TotStockVolume > Great_TotVol_value Then
                    Great_TotVol_value = TotStockVolume
                    Great_TotVol_Ticker = Sheets(J).Cells(i, 1)
                End If
                
                TotStockVolume = 0
                ReptPos = ReptPos + 1
                Init_Stock_Price = 0
                FirstRow = True
            End If
        Next i
        Sheets(J).Cells(2, 15) = "Greatest % Increase"
        Sheets(J).Cells(2, 16) = Great_Per_In_Ticker
        Sheets(J).Cells(2, 17) = Great_Per_In_Value
        
        Sheets(J).Cells(3, 15) = "Greatest % Decrease"
        Sheets(J).Cells(3, 16) = Great_Per_De_Ticker
        Sheets(J).Cells(3, 17) = Great_Per_De_Value
         
        Sheets(J).Cells(4, 15) = "Greatest Total Volume"
        Sheets(J).Cells(4, 16) = Great_TotVol_Ticker
        Sheets(J).Cells(4, 17) = Great_TotVol_value
    Next
End Sub
