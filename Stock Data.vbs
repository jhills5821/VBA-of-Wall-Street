Sub stockdata():

Dim ws As Worksheet

For Each ws In Worksheets
    
    Dim rowcounter As Long
    Dim iftrue As Long
    Dim n As Integer
    Dim linecount As Long

    rowcounter = 2
    iftrue = 1
    n = 1

    linecount = WorksheetFunction.CountA(Columns(1))

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Volume"

    For i = 1 To linecount

        Dim row1 As String
        Dim row2 As String
        Dim ticker As String
        Dim totalvolume As Double
        Dim volume As Long
        Dim openprice As Double
        Dim closeprice As Double
        Dim openrow As Long
        Dim closerow As Long
        Dim yearlychange As Double
        Dim percentchange As Double
           
          
        
        row1 = ws.Cells(i + 1, 1).Value
        row2 = ws.Cells(i + 2, 1).Value
        ticker = ws.Cells(i, 1).Value
        volume = ws.Cells(i + 1, 7).Value
        rowcounter = rowcounter + 1
    
            If row1 = row2 Then
                iftrue = iftrue + 1
                totalvolume = totalvolume + volume
            Else
                openrow = rowcounter - iftrue
                closerow = iftrue + openrow - 1
                openprice = ws.Cells(openrow, 3).Value
                closeprice = ws.Cells(closerow, 6).Value
                yearlychange = closeprice - openprice
                If openprice = 0 Then
                    percentchange = 0
                Else
                    percentchange = yearlychange / openprice
                        If percentchange < 0 Then
                            ws.Cells(n + 1, 11).Interior.ColorIndex = 3
                        Else
                            ws.Cells(n + 1, 11).Interior.ColorIndex = 4
                        End If
                End If
                ws.Cells(n + 1, 9).Value = ticker
                ws.Cells(n + 1, 10).Value = yearlychange
                ws.Cells(n + 1, 11).Value = percentchange
                ws.Cells(n + 1, 12).Value = totalvolume + volume
                iftrue = 1
                n = n + 1
                totalvolume = 0
            End If
    
    Next i

    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"

    Dim linecount2 As Long
    Dim maxvalue As Double
    Dim minvalue As Double
    Dim volvalue As Double

    linecount2 = WorksheetFunction.CountA(Columns(9))
    maxvalue = 0
    minvalue = 0
    volvalue = 0

    For i = 1 To linecount2

    Dim testval As Double
    Dim voltestval As Double
    Dim maxrowvalue As Integer
    Dim mainrowvalue As Integer
    Dim volrowvalue As Integer

    testval = ws.Cells(i + 1, 11).Value
    voltestval = ws.Cells(i + 1, 12).Value

        If testval > maxvalue Then
            maxvalue = testval
            maxrowvalue = i + 1
        End If
        If testval < minvalue Then
            minvalue = testval
            minrowvalue = i + 1
        End If
        If voltestval > volvalue Then
            volvalue = voltestval
            volrowvalue = i + 1
        End If
    Next i

    Dim maxticker2 As String
    Dim minticker2 As String
    Dim volticker2 As String

    maxticker2 = ws.Cells(maxrowvalue, 9).Value
    minticker2 = ws.Cells(minrowvalue, 9).Value
    volticker2 = ws.Cells(volrowvalue, 9).Value

    ws.Cells(2, 16).Value = maxticker2
    ws.Cells(2, 17).Value = maxvalue

    ws.Cells(3, 16).Value = minticker2
    ws.Cells(3, 17).Value = minvalue

    ws.Cells(4, 16).Value = volticker2
    ws.Cells(4, 17).Value = volvalue
    
Next ws

End Sub