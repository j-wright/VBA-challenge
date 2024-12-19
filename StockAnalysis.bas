Attribute VB_Name = "Module1"
Sub StockAnalysis()
'MsgBox ("Test message")
Dim ticker As String
Dim tickerArray(), currentPage() As String
Dim quarterChange, percentChange As Double
Dim openQuarter, closeQuarter, smallestPerc, largestPerc As Double
Dim totalVolume, newVolume, largestVolume As LongLong
Dim rowIndx, rowCount, i, numPages As Integer
Dim tickerCounter As Integer

Dim dev As Boolean
dev = False

If dev = False Then
    currentPage = Split("Q1 Q2 Q3 Q4", " ")
    ' find row count for a given page, assumes stock data starting at row 2 with no gaps
    numPages = 3
Else
    currentPage = Split("A B C D E F", " ")
    ' find row count for a given page, assumes stock data starting at row 2 with no gaps
    numPages = 5
End If
For pageIndx = 0 To numPages
ticker = ""
oldTicker = ""
tickerCounter = 0

' setup headers and table
Worksheets(currentPage(pageIndx)).Range("I1").Value = "Ticker"
Worksheets(currentPage(pageIndx)).Range("J1").Value = "Quarterly Change"
Worksheets(currentPage(pageIndx)).Range("K1").Value = "Percent Change"
Worksheets(currentPage(pageIndx)).Range("L1").Value = "Total Stock Volume"

Worksheets(currentPage(pageIndx)).Range("O2").Value = "Greatest % Increase"
Worksheets(currentPage(pageIndx)).Range("O3").Value = "Greatest % Decrease"
Worksheets(currentPage(pageIndx)).Range("O4").Value = "Greatest Total Volume"
Worksheets(currentPage(pageIndx)).Range("P1").Value = "Ticker"
Worksheets(currentPage(pageIndx)).Range("Q1").Value = "Value"


rowCount = findRowCount(currentPage(pageIndx))
'MsgBox ("Page: " & currentPage(pageIndx) & " Last Row: " & rowCount)
    For rowIndx = 2 To rowCount
        ticker = Worksheets(currentPage(pageIndx)).Cells(rowIndx, 1)
        

        If ticker <> "" And ticker <> oldTicker Then ' we have a new ticker
            tickerCounter = tickerCounter + 1
            Worksheets(currentPage(pageIndx)).Cells(tickerCounter + 1, 9).Value = ticker
            oldTicker = ticker
            totalVolume = Worksheets(currentPage(pageIndx)).Cells(rowIndx, 7)
            openQuarter = Worksheets(currentPage(pageIndx)).Cells(rowIndx, 3)
'             MsgBox ("New ticker: " & ticker & " Open : " & openQuarter)
        ElseIf Worksheets(currentPage(pageIndx)).Cells(rowIndx + 1, 1) <> "" Then ' not at last row
            newVolume = Worksheets(currentPage(pageIndx)).Cells(rowIndx, 7)
            totalVolume = totalVolume + newVolume
            Worksheets(currentPage(pageIndx)).Cells(tickerCounter + 1, 12) = totalVolume
                If ticker <> Worksheets(currentPage(pageIndx)).Cells(rowIndx + 1, 1).Value Then
                    closeQuarter = Worksheets(currentPage(pageIndx)).Cells(rowIndx, 6).Value
                    quarterChange = closeQuarter - openQuarter
                    percentChanged = quarterChange / openQuarter
                    Worksheets(currentPage(pageIndx)).Cells(tickerCounter + 1, 10).Value = quarterChange
                    Worksheets(currentPage(pageIndx)).Cells(tickerCounter + 1, 11).Value = percentChanged
                    Worksheets(currentPage(pageIndx)).Cells(tickerCounter + 1, 11).NumberFormat = "0.00%"
                End If
        Else ' at last row
            closeQuarter = Worksheets(currentPage(pageIndx)).Cells(rowIndx, 6).Value
            quarterChange = closeQuarter - openQuarter
            percentChanged = quarterChange / openQuarter
            Worksheets(currentPage(pageIndx)).Cells(tickerCounter + 1, 10).Value = quarterChange
            Worksheets(currentPage(pageIndx)).Cells(tickerCounter + 1, 11).Value = percentChanged
            Worksheets(currentPage(pageIndx)).Cells(tickerCounter + 1, 11).NumberFormat = "0.00%"
        End If

    Next rowIndx
        smallestPerc = WorksheetFunction.Min(Worksheets(currentPage(pageIndx)).Range("K2:K" & tickerCounter).Value)
        largestPerc = WorksheetFunction.Max(Worksheets(currentPage(pageIndx)).Range("K2:K" & tickerCounter).Value)
        largestVolume = WorksheetFunction.Max(Worksheets(currentPage(pageIndx)).Range("L2:L" & tickerCounter).Value)
        
        Worksheets(currentPage(pageIndx)).Range("Q2").Value = largestPerc
        Worksheets(currentPage(pageIndx)).Range("Q3").Value = smallestPerc
        Worksheets(currentPage(pageIndx)).Range("Q4").Value = largestVolume
        Worksheets(currentPage(pageIndx)).Cells(2, 17).NumberFormat = "0.00%"
        Worksheets(currentPage(pageIndx)).Cells(3, 17).NumberFormat = "0.00%"

        i = 2
        While Worksheets(currentPage(pageIndx)).Cells(i, 12).Value <> largestVolume
         i = i + 1
        Wend
        Worksheets(currentPage(pageIndx)).Range("P4").Value = Worksheets(currentPage(pageIndx)).Cells(i, 9).Value
        
         i = 2
        While Worksheets(currentPage(pageIndx)).Cells(i, 11).Value <> largestPerc
         i = i + 1
        Wend
        Worksheets(currentPage(pageIndx)).Range("P2").Value = Worksheets(currentPage(pageIndx)).Cells(i, 9).Value

        i = 2
        While Worksheets(currentPage(pageIndx)).Cells(i, 11).Value <> smallestPerc
         i = i + 1
        Wend
        Worksheets(currentPage(pageIndx)).Range("P3").Value = Worksheets(currentPage(pageIndx)).Cells(i, 9).Value
        
        
        For i = 2 To tickerCounter + 1
            quarterChange = Worksheets(currentPage(pageIndx)).Cells(i, 10).Value
            If quarterChange < 0 Then
                Worksheets(currentPage(pageIndx)).Cells(i, 10).Interior.Color = RGB(255, 0, 0)
            ElseIf quarterChange > 0 Then
                Worksheets(currentPage(pageIndx)).Cells(i, 10).Interior.Color = RGB(0, 255, 0)
            End If
        Next i
        
        Worksheets(currentPage(pageIndx)).Columns("A:Q").AutoFit
        
Next pageIndx

'' just grab all stock tickers
'Dim rowIndx As Integer
'rowIndx = 2
'While Worksheets("A").Cells(rowIndx, 1).Value <> ""
''   MsgBox ("Row index: " & rowIndx & " ticker: " & Worksheets("A").Cells(rowIndx, 1).Value)
'
'   rowIndx = rowIndx + 1
'Wend
End Sub

Function findRowCount(page)
' find row count
rowIndx = 2
'rowCount = 0
While Worksheets(page).Cells(rowIndx, 1).Value <> ""
   rowIndx = rowIndx + 1
Wend
    findRowCount = rowIndx - 1
End Function

