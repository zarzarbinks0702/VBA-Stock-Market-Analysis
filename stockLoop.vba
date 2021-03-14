'used (Macro to Loop Through All Worksheets in a Workbook. (n.d.). Retrieved March 14, 2021, from
'https://support.microsoft.com/en-us/topic/macro-to-loop-through-all-worksheets-in-a-workbook-feef14e3-97cf-00e2-538b-5da40186e2b0)
'to loop through all the worksheets
'used in class activities from lessons 2.1, 2.2, and 2.3 as guides


'Function to go through all the data
Sub stockLoop():
    'declaring variables
    Dim wsCount As Integer '<- used to loop through the whole workbook
    Dim w As Integer '<- variable used to loop through whole workbook
    Dim lastRow As Long '<- the last filled row in the sheet
    Dim tickerID As String '<- holds the ticker id to put in the table
    Dim tickerTable As Integer '<- used to put ticker id in the table
    Dim yearChange As Double '<- yearly change from first opening to last closing
    Dim percentChange As Double '<- percent change from first opening to last closing
    Dim stockVolume As Long '<- total volume of the stock
    Dim column As Integer '<- used to grab the ticker ids

    'defining wide variables
    wsCount = ActiveWorkbook.Worksheets.Count
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    tickerTable = 2
    column = 1
    'loop through whole workbook
    For w = 1 To wsCount
        For i = 1 To lastRow
            If Cells(i + 1, column).Value <> Cells(i, column).Value Then
                tickerID = Cells(i + 1, column).Value
                Range("I" & tickerTable).Value = tickerID
                tickerTable = tickerTable + 1
            End If
        Next i
    Next w

End Sub
