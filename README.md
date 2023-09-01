# VBA-Files

Sub ticker_03()
For Each ws In Worksheets

     Dim ticker As String
     ticker = ""

     Dim startdate As Long

     Dim openvalue As Double

     Dim closevalue As Double

     Dim yearlychange As Double

     Dim percentchange As Double

     Dim volume As LongLong
     volume = 0

     Dim tickercount As Integer
     tickercount = 1

     Dim i As Long

     LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

     For i = 2 To LastRow

            If ws.Cells(i, 1).Value = ticker Then
            volume = volume + ws.Cells(i, 7).Value

            Else: tickercount = tickercount + 1
            ticker = ws.Cells(i, 1).Value
            ws.Cells(tickercount, 9).Value = ticker
            volume = 0 + ws.Cells(i, 7).Value
            startdate = ws.Cells(i, 2).Value
            openvalue = ws.Cells(i, 3).Value
            End If

            If ws.Cells(i + 1, 1).Value <> ticker Then
            closevalue = ws.Cells(i, 6).Value
            yearlychange = closevalue - openvalue
            percentchange = yearlychange / openvalue
            percentchange = Application.WorksheetFunction.Round(percentchange, 4)
            ws.Cells(tickercount, 10).Value = yearlychange
            ws.Cells(tickercount, 11).Value = percentchange
            ws.Cells(tickercount, 12).Value = volume
                If yearlychange < 0 Then
                ws.Cells(tickercount, 10).Interior.ColorIndex = 3
                ElseIf yearlychange > 0 Then
                ws.Cells(tickercount, 10).Interior.ColorIndex = 4
                End If
            End If
        Next i

        ws.Columns("K").NumberFormat = "0.00%"
        ws.Columns("L").NumberFormat = "0"

        With ws.Range("I1:L1")
        .NumberFormat = "Text"
        .Font.Bold = True
        End With

        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Columns("I:L").AutoFit


        Dim inc As Double
        inc = 0
        Dim dec As Double
        dec = 0
        Dim maxVol As LongLong
        maxVol = 0
        Dim incTic As String
        Dim decTic As String
        Dim maxVolTic As String

        dataEnd = ws.Cells(Rows.Count, 9).End(xlUp).Row

        For i = 2 To dataEnd

            If inc < ws.Cells(i, 11).Value Then
            inc = ws.Cells(i, 11).Value
            incTic = ws.Cells(i, 9).Value
            End If
            If dec > ws.Cells(i, 11).Value Then
            dec = ws.Cells(i, 11).Value
            decTic = ws.Cells(i, 9).Value
            End If
            If maxVol < ws.Cells(i, 12).Value Then
            maxVol = ws.Cells(i, 12).Value
            maxVolTic = ws.Cells(i, 9).Value
            End If
        Next i

        ws.Range("P2").Value = incTic
        ws.Range("Q2").Value = inc
        ws.Range("P3").Value = decTic
        ws.Range("Q3").Value = dec
        ws.Range("P4").Value = maxVolTic
        ws.Range("Q4").Value = maxVol

        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        ws.Range("Q4").NumberFormat = "0,000"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O1:Q4").Font.Bold = True
        ws.Columns("O:Q").AutoFit
    Next
End Sub
