Sub VBA_Homework()

' Name all variables for part I
Dim ticker As String
Dim Summary_Table_Row As Integer
Dim lastRow As Long
Dim date_open As Double
Dim date_close As Double
Dim year_change As Double
Dim percent_change As Double
Dim total_volume As Double

' loop worksheets
For Each ws In Worksheets

    ' Find the last row
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ' Add headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"


    ' Start variables
    Summary_Table_Row = 0
    year_change = 0
    date_open = 0
    percent_change = 0
    total_volume = 0

    ' Start loop
    For i = 2 To lastRow

        ' Get Ticker value
        ticker = ws.Cells(i, 1).Value
        
        ' Get the start opening price
        If date_open = 0 Then
            date_open = ws.Cells(i, 3).Value
        
        End If
        
        ' Calculate total stock volume
        total_volume = total_volume + ws.Cells(i, 7).Value

        
        ' Check to look for different ticker
        If ws.Cells(i + 1, 1).Value <> ticker Then
            
            ' Add 1 when a different ticker is found
            Summary_Table_Row = Summary_Table_Row + 1
                
            ' Print ticker in column 9
            ws.Cells(Summary_Table_Row + 1, 9) = ticker
            
            ' Get closing price
            date_close = ws.Cells(i, 6)
            
            ' Calculate yearly change
            year_change = date_close - date_open
            
            ' Print year_change in column 10
            ws.Cells(Summary_Table_Row + 1, 10).Value = year_change
            
                ' Color Yearly Change column green > 0 or red < 0
                ' If yearly change value is greater than 0, color cell green.
                If year_change > 0 Then
                    ws.Cells(Summary_Table_Row + 1, 10).Interior.ColorIndex = 4
                
                    ' If yearly change value is less than 0, color cell red.
                    Else: ws.Cells(Summary_Table_Row + 1, 10).Interior.ColorIndex = 3
           
                End If
            
            
                ' Calculate percent change
                If date_open = 0 Then
                    percent_change = 0

                Else
                    percent_change = (year_change / date_open)

                End If
            
            ' Print percent_change in column 11
            ws.Cells(Summary_Table_Row + 1, 11).Value = percent_change

            ' Format the column 11 to percent
            ws.Cells(Summary_Table_Row + 1, 11).Value = Format(percent_change, "0.00%")
            
            ' Reset open
            date_open = 0

            ' Print total_volume in column 12
            ws.Cells(Summary_Table_Row + 1, 12).Value = total_volume
            
            ' Reset total volume
            total_volume = 0

        End If
        
    Next i
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Start Challenges

    ' Name variables
    Dim gpercent_inc As Double
    Dim gpercent_inc_ticker As String
    Dim gpercent_dec As Double
    Dim gpercent_dec_ticker As String
    Dim gtotal_vol As Double
    Dim gtotal_vol_ticker As String

    ' Get the last row
    lastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row

    ' Add new headers/sections for the challenge
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    ' Start variables
    gpercent_inc = ws.Cells(i, 11).Value
    gpercent_inc_ticker = ws.Cells(i, 9).Value
    gpercent_dec = ws.Cells(i, 11).Value
    gpercent_dec_ticker = ws.Cells(i, 9).Value
    gtotal_vol = ws.Cells(i, 12).Value
    gtotal_vol_ticker = ws.Cells(i, 9).Value
     

    ' Start the loop
    For i = 2 To lastRow
    
        ' Get greatest percent increase ticker
        If ws.Cells(i, 11).Value > gpercent_inc Then
            gpercent_inc = ws.Cells(i, 11).Value
            gpercent_inc_ticker = ws.Cells(i, 9).Value

        End If
        
        ' Get greatest percent decrease ticker
        If ws.Cells(i, 11).Value < gpercent_dec Then
            gpercent_dec = ws.Cells(i, 11).Value
            gpercent_dec_ticker = ws.Cells(i, 9).Value

        End If
        
        ' Get greatest total volume ticker
        If ws.Cells(i, 12).Value > gtotal_vol Then
            gtotal_vol = ws.Cells(i, 12).Value
            gtotal_vol_ticker = ws.Cells(i, 9).Value

        End If
        
        ' Print challenge values
        ws.Cells(2, 16).Value = gpercent_inc_ticker
        ws.Cells(2, 17).Value = Format(gpercent_inc, "0.00%")
        ws.Cells(3, 16).Value = gpercent_dec_ticker
        ws.Cells(3, 17).Value = Format(gpercent_dec, "0.00%")
        ws.Cells(4, 16).Value = gtotal_vol_ticker
        ws.Cells(4, 17).Value = gtotal_vol
    
    Next i
    
Next ws


End Sub
