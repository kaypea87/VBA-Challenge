Attribute VB_Name = "Module1"
Sub RunMacro()
'source:https://www.extendoffice.com/documents/excel/5333-excel-run-macro-multiple-sheets.html
    
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call multi
    Next
    Application.ScreenUpdating = True
End Sub

Sub multi()

'---------- Unique list of tickers------------
'source: https://www.statology.org/vba-get-unique-values-from-column/#:~:text=You%20can%20use%20the%20AdvancedFilter,from%20a%20column%20in%20Excel.&text=This%20particular%20example%20extracts%20a,them%20starting%20in%20cell%20E1.
Range("A:A").AdvancedFilter Action:=xlFilterCopy, CopytoRange:=Range("J1"), Unique:=True
Range("J1").Value = "Ticker"

'----------- Volume Total ------------------
'Assign variable for column for sumif formula
Dim v As Integer
lastrow = Cells(Rows.Count, 10).End(xlUp).Row

'column J for tickers is 10 and column for output is 13
For v = 2 To lastrow
    Cells(v, 13).Value = WorksheetFunction.sumif(Range("A:A"), Cells(v, 10), Range("G:G"))
Next v
Range("M1").Value = "Total Stock Volume"

'-----------Yearly and Percentage Change and conditional formatting---------
Dim openprice As Double
Dim closeprice As Double
'last row for column A
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
'last row for column J that contains tickers from the filter in unique tickers
lastrow_tickers = Cells(Rows.Count, 10).End(xlUp).Row

Dim openprice_found As Boolean

For t = 2 To lastrow_tickers
'set open price to false because we have not found it yet
    openprice_found = False
 'loop to check if tickers in column A match those in column J and find open/close prices
    For i = 2 To lastrow
        If Cells(i, 1).Value = Cells(t, 10).Value Then
            closeprice = Cells(i, 6).Value
            If openprice_found = False Then
                openprice_found = True
                openprice = Cells(i, 3).Value
            End If
        End If
    Next i
    
'calculate the yearly change and percent change
    Cells(t, 11).Value = closeprice - openprice
    Cells(t, 12).Value = Cells(t, 11).Value / openprice
'conditional formatting of the cells in yearly change
     If Cells(t, 11).Value >= 0 Then
        Cells(t, 11).Interior.ColorIndex = 4
    Else
        Cells(t, 11).Interior.ColorIndex = 3
    End If
     If Cells(t, 12).Value >= 0 Then
        Cells(t, 12).Interior.ColorIndex = 4
    Else
        Cells(t, 12).Interior.ColorIndex = 3
    End If
Next t
Range("K1").Value = "Yearly Change"
Range("L1").Value = "Percent Change"
Range("L:L").NumberFormat = "0.00%"

'----------Summary Table-----------

Cells(2, 16).Value = ("Greatest % Increase")
Cells(3, 16).Value = ("Greatest % Decrease")
Cells(4, 16).Value = ("Greatest Total Volume")
Cells(1, 17).Value = ("Ticker")
Cells(1, 18).Value = ("Value")

Dim maxval As Double
Dim minval As Double
Dim maxvol As Double
Dim tickermax As String
Dim tickermin As String
Dim tickervol As String

lastrow_tickers = Cells(Rows.Count, 10).End(xlUp).Row

maxval = Cells(2, 12).Value
minval = Cells(2, 12).Value
maxvol = Cells(2, 13).Value
For i = 2 To lastrow_tickers
    If Cells(i, 12).Value > maxval Then
        maxval = Cells(i, 12).Value
        tickermax = Cells(i, 10).Value
    End If
    If Cells(i, 12).Value < minval Then
        minval = Cells(i, 12).Value
        tickermin = Cells(i, 10).Value
    End If
    If Cells(i, 13).Value > maxvol Then
    maxvol = Cells(i, 13).Value
    tickervol = Cells(i, 10).Value
    End If

Next i

Cells(2, 18).Value = maxval
Cells(2, 18).NumberFormat = "0.00%"
Cells(2, 17).Value = tickermax
Cells(3, 18).Value = minval
Cells(3, 18).NumberFormat = "0.00%"
Cells(3, 17).Value = tickermin
Cells(4, 18).Value = maxvol
Cells(4, 17).Value = tickervol

End Sub

