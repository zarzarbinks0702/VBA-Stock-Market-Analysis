'Function to go through all the data
Sub stockLoop():
    'declaring variables (Tom (AnalystCave), 2021), (Vba variable types: Declare different types of variable in excel vba 2020)
    Dim lastRow As Long '<- the last filled row in the sheet
    Dim tickerID As String '<- holds the ticker id to put in the table
    Dim yearChange As Double '<- yearly change from first opening to last closing
    Dim percentChange As Double '<- percent change from first opening to last closing
    Dim stockVolume As Single '<- total volume of the stock
    Dim tableSpot As Single '<- used to put values in table
    Dim openValue As Double '<- used in change calculations
    Dim closeValue As Double '<- used in change calculations

        'defining beginning variables to reset for each sheet
        lastRow = Cells(rows.Count, 1).End(xlUp).Row
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        tableSpot = 2
        stockVolume = 0
        openValue = Cells(2, 3).Value '<- gives first open value from table

        For i = 2 To lastRow

                        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then '<- if the next cell is a different ticker id
                tickerID = Cells(i, 1).Value
                closeValue = Cells(i, 6).Value
                yearChange = closeValue - openValue
                percentChange = yearChange / openValue * 100
                Range("I" & tableSpot).Value = tickerID
                Range("J" & tableSpot).Value = Round(yearChange, 2)

                'change color of cell based on positive or negative yearly change
                If yearChange >= 0 Then
                    Range("J" & tableSpot).Interior.ColorIndex = 4
                Else
                    Range("J" & tableSpot).Interior.ColorIndex = 3
                End If

                Range("K" & tableSpot).Value = Format(percentChange, "#.##""%") '(Y, 2021 VBA format)
                tableSpot = tableSpot + 1
                stockVolume = 0
                openValue = Cells(i + 1, 3).Value
            Else '<- if the next cell is the same ticker id
                stockVolume = stockVolume + Cells(i, 7).Value
                Range("L" & tableSpot) = stockVolume
            End If

        Next i

End Sub



'References in code:
'Tom (AnalystCave). (2021, January 04). Excel worksheets TUTORIAL: VBA Activesheet vs WORKSHEETS. Retrieved March 17, 2021, from https://analystcave.com/excel-vba-worksheets-tutorial-vba-activesheet-vs-worksheets/
'Vba variable types: Declare different types of variable in excel vba. (2020, August 21). Retrieved March 17, 2021, from https://www.educba.com/vba-variable-types/
'Y, J. A. (2021, February 22). Vba format: How to use vba format function? (examples). Retrieved March 17, 2021, from https://www.wallstreetmojo.com/vba-format/
'used in class activities from lessons 2.1, 2.2, and 2.3 as guides
