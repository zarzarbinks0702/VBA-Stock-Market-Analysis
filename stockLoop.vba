'used (Macro to Loop Through All Worksheets in a Workbook. (n.d.). Retrieved March 14, 2021, from
'https://support.microsoft.com/en-us/topic/macro-to-loop-through-all-worksheets-in-a-workbook-feef14e3-97cf-00e2-538b-5da40186e2b0)
'to loop through all the worksheets
'used in class activities from lessons 2.1, 2.2, and 2.3 as guides


'Function to go through all the data
Sub stockLoop():
    'declaring variables
    'Dim wsCount As Integer '<- used to loop through the whole workbook
    Dim ws As Worksheet '<- variable used to loop through whole workbook
    Dim lastRow As Long '<- the last filled row in the sheet
    Dim tickerID As String '<- holds the ticker id to put in the table
    Dim yearChange As Double '<- yearly change from first opening to last closing
    Dim percentChange As Double '<- percent change from first opening to last closing
    Dim stockVolume As Single '<- total volume of the stock
    Dim tableSpot As Single '<- used to put values in table
    Dim openValue As Double '<- used in change calculations
    Dim closeValue As Double '<- used in change calculations

    'defining wide variables
    'wsCount = ActiveWorkbook.Worksheets.Count
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    tableSpot = 2
    stockVolume = 0

    'loop through whole workbook
    For Each ws In ThisWorkbook.Worksheets
        For i = 2 To lastRow
            'get values to calculate yearly change and percent change
            If Cells(i, 2).Value = "20160101" Then
                openValue = Cells(i, 3).Value
            ElseIf Cells(i, 2).Value = "20161230" Then
                closeValue = Cells(i, 6).Value
            End If

            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then '<- if the next cell is a different ticker id
                tickerID = Cells(i, 1).Value
                yearChange = closeValue - openValue
                percentChange = yearChange / openValue * 100
                ws.Range("I" & tableSpot).Value = tickerID
                ws.Range("J" & tableSpot).Value = yearChange
                ws.Range("K" & tableSpot).Value = percentChange
                tableSpot = tableSpot + 1
                stockVolume = 0
            Else '<- if the next cell is the same ticker id
                stockVolume = stockVolume + Cells(i, 7).Value
                ws.Range("L" & tableSpot) = stockVolume
            End If

        Next i
    Next ws

End Sub
