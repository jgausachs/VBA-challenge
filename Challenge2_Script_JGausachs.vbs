Attribute VB_Name = "Module1"
Sub priceMove():
'Caution - this subroutine ONLY works with a dataset sorted by field: <ticker>

    Dim activeTicker As String          ' active ticker code; for use in loop
    Dim firstDate As Long               ' earliest date for a given ticker; initial date anchor
    Dim lastDate As Long                ' latest date for a given ticker; final date anchor
    Dim lastRow As Double               ' the last populated row number
    Dim openPrice As Currency           ' share price at open of period
    Dim closePrice As Currency          ' share price at end of period
    Dim sumVolume As Double             ' cummulative trade volume over period
    Dim sryTableRow As Integer          ' active row for displaying ticker in summary table
    Dim greatPctIncrTick As String      ' greatest percent value increase ticker
    Dim greatPctIncrVal As Double       ' greatest percent value increase value
    Dim greatPctDecrTick As String      ' greatest percent value decrese ticker
    Dim greatPctDecrVal As Double       ' greatest percent value decrease value
    Dim greatVolTick As String          ' greatest trade volume ticker
    Dim greatVolVal As Double           ' greatest trade volume value


'Loop through all sheets
For Each ws In Worksheets

'Initial settings
sryTableRow = 2
sumVolume = 0
greatPctIncrTick = ""
greatPctIncrVal = 0
greatPctDecrTick = ""
greatPctDecrVal = 0
greatVolTick = ""
greatVolVal = 0

firstDate = ws.Cells(2, 2).Value
lastDate = ws.Cells(3, 2).Value
openPrice = ws.Cells(2, 3).Value
closePice = ws.Cells(2, 6).Value
ws.Columns("J").NumberFormat = "#,##0.00"
ws.Columns("K").NumberFormat = "0.00%"
ws.Columns("Q").NumberFormat = "0.00%"
ws.Range("Q4").NumberFormat = "##0"

'Set column headers
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

'Find last row in worksheet
lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Loop through each row in worksheet
For i = 2 To lastRow

    'Check for a new ticker on next row - run instructions prior to moving to new ticker row
    If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
      
        'Record open/close share price - if warranted - by reading last row in looped ticker
        If (firstDate > ws.Cells(i, 2).Value) Then
            firstDate = ws.Cells(i, 2).Value
            openPrice = ws.Cells(i, 3).Value
        Else:
            lastDate = ws.Cells(i, 2).Value
            closePice = ws.Cells(i, 6).Value
        End If

        'Write active ticker code, price change and cummulative trade volume in summary table
        activeTicker = Cells(i, 1).Value
        ws.Range("I" & sryTableRow).Value = activeTicker
        ws.Range("J" & sryTableRow).Value = closePice - openPrice
        ws.Range("K" & sryTableRow).Value = (closePice - openPrice) / openPrice
        ws.Range("L" & sryTableRow).Value = sumVolume + ws.Cells(i, 7).Value

        'If current price increase exceeds stored value then replace stored ticker/value with current one
        If (((closePice - openPrice) / openPrice) > greatPctIncrVal) Then
            greatPctIncrVal = (closePice - openPrice) / openPrice
            greatPctIncrTick = ws.Cells(i, 1).Value
        End If

        'If current price decrease exceeds stored value then replace stored ticker/value with current one
        If (((closePice - openPrice) / openPrice) < greatPctDecrVal) Then
            greatPctDecrVal = (closePice - openPrice) / openPrice
            greatPctDecrTick = ws.Cells(i, 1).Value
        End If

        'If current trade volume exceeds stored value then replace stored ticker/value with current one
        If (sumVolume + ws.Cells(i, 7).Value > greatVolVal) Then
            greatVolVal = sumVolume + ws.Cells(i, 7).Value
            greatVolTick = ws.Cells(i, 1).Value
        End If

        'Conditional colour formatting of price change cell
        If ((closePice - openPrice) < 0) Then
            ws.Range("J" & sryTableRow).Interior.ColorIndex = 3
        Else:
            ws.Range("J" & sryTableRow).Interior.ColorIndex = 4
        End If

        'Increase counter for upcoming ticker write-up in summary table
        sryTableRow = sryTableRow + 1
        sumVolume = 0

        'Loop through rows with same ticker and runs instruction set
        'Store earliest date if exists in current row
        ElseIf (firstDate > ws.Cells(i, 2).Value) Then
            firstDate = ws.Cells(i, 2).Value
            openPrice = ws.Cells(i, 3).Value
    
        'Store latest date - if existing - in current row
        Else:
            lastDate = ws.Cells(i, 2).Value
            closePice = ws.Cells(i, 6).Value
    End If
    
    sumVolume = sumVolume + ws.Cells(i, 7).Value

Next i

'Display greatest increase / greatest decrease / greatest trade volume
ws.Range("P2").Value = greatPctIncrTick
ws.Range("Q2").Value = greatPctIncrVal
ws.Range("P3").Value = greatPctDecrTick
ws.Range("Q3").Value = greatPctDecrVal
ws.Range("P4").Value = greatVolTick
ws.Range("Q4").Value = greatVolVal

'Autofit width of display columns
ws.Columns("I:L").AutoFit
ws.Columns("O:Q").AutoFit

Next ws

End Sub
