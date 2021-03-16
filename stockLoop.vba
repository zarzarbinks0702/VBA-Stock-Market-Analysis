'used (Macro to Loop Through All Worksheets in a Workbook. (n.d.). Retrieved March 14, 2021, from
'https://support.microsoft.com/en-us/topic/macro-to-loop-through-all-worksheets-in-a-workbook-feef14e3-97cf-00e2-538b-5da40186e2b0)
'to loop through all the worksheets
'used in class activities from lessons 2.1, 2.2, and 2.3 as guides


'Function to go through all the data
Sub stockLoop():
    'declaring variables
    Dim wsCount As Integer '<- used to loop through the whole workbook
    Dim ws As Worksheet '<- variable used to loop through whole workbook
    Dim lastRow As Long '<- the last filled row in the sheet
    Dim tickerID As String '<- holds the ticker id to put in the table
    Dim tickerTable As Integer '<- used to put ticker id in the table
    Dim yearChange As Double '<- yearly change from first opening to last closing
    Dim yearTable As Double '<- used to put year change in the table
    Dim percentChange As Double '<- percent change from first opening to last closing
    Dim percentTable As Double '<- used to put percent change in the table
    Dim stockVolume As Single '<- total volume of the stock
    Dim volumeTable As Single '<- used to put stock volume total in table
    Dim column As Integer '<- used to grab the ticker ids

    'defining wide variables
    wsCount = ActiveWorkbook.Worksheets.Count
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    tickerTable = 2
    column = 1
    yearTable = 2
    percentTable = 2
    stockVolume = 0
    volumeTable = 1

    'loop through whole workbook
    For Each ws In Worksheets
        For i = 1 To lastRow
            If Cells(i + 1, column).Value <> Cells(i, column).Value Then '<- if the next cell is a different ticker id
                tickerID = Cells(i + 1, column).Value
                ws.Range("I" & tickerTable).Value = tickerID
                tickerTable = tickerTable + 1
                volumeTable = volumeTable + 1
                stockVolume = 0
            Else '<- if the next cell is the same ticker id
                stockVolume = stockVolume + Cells(i, 7).Value
                ws.Range("L" & volumeTable) = stockVolume
            End If
        Next i
    Next ws

End Sub
