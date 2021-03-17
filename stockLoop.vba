'Function to go through all the data
Sub stockLoop():
    'declaring variables (Tom (AnalystCave), 2021), (Vba variable types: Declare different types of variable in excel vba 2020)
    Dim wkst() As Worksheet '<- variable used to loop through whole workbook
    Dim ws As Variant '<- used to loop through whole workbook
    Dim wsCount As Integer '<- used to loop through whole workbook
    Dim lastRow As Long '<- the last filled row in the sheet
    Dim tickerID As String '<- holds the ticker id to put in the table
    Dim yearChange As Double '<- yearly change from first opening to last closing
    Dim percentChange As Double '<- percent change from first opening to last closing
    Dim stockVolume As Single '<- total volume of the stock
    Dim tableSpot As Single '<- used to put values in table
    Dim openValue As Double '<- used in change calculations
    Dim closeValue As Double '<- used in change calculations

    'creating list for looping through sheets (Loop through all worksheets with for each - vba code examples 2019)
    wsCount = ThisWorkbook.Worksheets.Count - 1
    ReDim wkst(wsCount)
    For i = LBound(wkst) To UBound(wkst)
        Set wkst(i) = ThisWorkbook.Sheets(i + 1)
    Next i

    'defining wide variables
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row

    'loop through whole workbook
    For Each ws In wkst
        'defining beginning variables to reset for each sheet
        tableSpot = 2
        stockVolume = 0
        For i = 2 To lastRow
            openValue = Cells(2, 3).Value '<- gives first open value from table
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then '<- if the next cell is a different ticker id
                tickerID = Cells(i, 1).Value
                closeValue = Cells(i, 6).Value
                yearChange = closeValue - openValue
                percentChange = yearChange / openValue * 100
                ws.Range("I" & tableSpot).Value = tickerID
                ws.Range("J" & tableSpot).Value = Round(yearChange, 2)

                'change color of cell based on positive or negative yearly change
                If yearChange.Value >= 0 Then
                    ws.Range("J" & tableSpot).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & tableSpot).Interior.ColorIndex = 3
                End If

                ws.Range("K" & tableSpot).Value = Format(percentChange, "#.##""%") '(Y, 2021)
                tableSpot = tableSpot + 1
                stockVolume = 0
                openValue = Cells(i + 1, 3).Value
            Else '<- if the next cell is the same ticker id
                stockVolume = stockVolume + Cells(i, 7).Value
                ws.Range("L" & tableSpot) = stockVolume
            End If

        Next i
    Next ws

End Sub

'References in code:
'Loop through all worksheets with for each - vba code examples.(2019, April 05). Retrieved March 16, 2021, from https://www.automateexcel.com/vba/cycle-and-update-all-worksheets/
'Tom (AnalystCave). (2021, January 04). Excel worksheets TUTORIAL: VBA Activesheet vs WORKSHEETS. Retrieved March 17, 2021, from https://analystcave.com/excel-vba-worksheets-tutorial-vba-activesheet-vs-worksheets/
'Vba variable types: Declare different types of variable in excel vba. (2020, August 21). Retrieved March 17, 2021, from https://www.educba.com/vba-variable-types/
'Y, J. A. (2021, February 22). Vba format: How to use vba format function? (examples). Retrieved March 17, 2021, from https://www.wallstreetmojo.com/vba-format/
'used in class activities from lessons 2.1, 2.2, and 2.3 as guides
