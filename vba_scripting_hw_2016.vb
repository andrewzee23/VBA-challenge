sub stockticker()


    Dim ticker_name As String
    Dim yclose As Double
    yclose = 0
    Dim yopen As Double
    yopen = 0
    Dim ychange As Double
    ychange = 0
    Dim pchange As Double
    pchange = 0
    Dim ticker_totvol As Double
    ticker_totvol = 0
    Dim summary_table_row As Integer
    summary_table_row = 2
    Dim max_ticker As String
    max_ticker = " "
    Dim min_ticker As String
    min_ticker = " "
    Dim maxpercent As Double
    maxpercent = 0
    Dim minpercent As Double
    minpercent = 0
    Dim maxvolume As Double
    maxvolume = 0
    Dim ticker_maxvol As String
    ticker_maxvol = " "

    On Error Resume Next



    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Vol"


    yopen = cells(2, 3).value

    lastrow = cells(rows.count, 1).End(xlUp).Row


    For i = 2 to lastrow



    'Rows.Count maybe
        'when rows.count get range error in VBA here
        if cells(i + 1, 1).Value <> cells(i, 1).Value then

            ticker_name = cells(i, 1).Value
            yclose = cells(i, 6).Value
            ychange = yclose - yopen
            pchange = (ychange / yopen)
            ticker_totvol = ticker_totvol + cells(i, 7).value


            range("I" & summary_table_row).value = ticker_name
            range("J" & summary_table_row).value = ychange
            range("K" & summary_table_row).value = pchange
            range("K" & summary_table_row).NumberFormat = "0.00%"
            range("L" & summary_table_row).value = ticker_totvol

            summary_table_row = summary_table_row + 1
            ychange = 0
            yclose = 0
            yopen = cells(i + 1, 3).value
            ticker_totvol = 0
        
        Else

            ticker_totvol = ticker_totvol + cells(i, 7).value

        End if

    

    next i


    lastrow_summary_table = Cells(Rows.Count, 11).End(xlUp).Row

    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1, 16).value = "Ticker"
    Cells(1, 17).value = "Value"

    For i = 2 to lastrow_summary_table
        If cells(i, 10).value > 0 then
            cells(i, 10).Interior.colorindex = 4
        Else
            cells(i, 10).interior.colorindex = 3
        end If
    next i



    For i = 2 to lastrow_summary_table

        If Cells(i, 11).Value = Application.WorksheetFunction.Max(Range("K2:K" & lastrow_summary_table)) Then
            Cells(2, 16).Value = cells(i, 9).Value
            cells(2, 17).value = Cells(i, 11).value
            cells(2, 17).NumberFormat = "0.00%"
        
        ElseIf Cells(i, 11).Value = Application.WorksheetFunction.Min(Range("K2:K" & lastrow_summary_table)) Then
            cells(3, 16).value = Cells(i, 9).value
            cells(3, 17).value = cells(i, 11).value
            cells(3, 17).NumberFormat = "0.00%"
        
        ElseIf Cells(i, 12).Value = Application.WorksheetFunction.Max(Range("L2:L" & lastrow_summary_table)) Then
            cells(4, 16).Value = cells(i, 9).Value
            cells(4, 17).value = cells(i, 12).Value

        end If


    next i

end sub
        