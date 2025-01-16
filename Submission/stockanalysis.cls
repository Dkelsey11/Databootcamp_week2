VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Name = "Module1"
Sub stockanalysis()

'Set Dimensions

Dim total As Double
Dim quarter As Double
Dim percent As Double
Dim i As Long
Dim j As Integer
Dim lastRow As Long
Dim start As Long


'Set Column Headers
Range("I1").Value = "Ticker"
Range("J1").Value = "Quarterly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

'Initialize Dimensions

j = 0
total = 0
quarter = 0
start = 2

'Find last row number

lastRow = Cells(Rows.Count, "A").End(xlUp).Row


'Create a for loop (First thing we need to do is set up conditional to switch ticker, the next thing we need to do is set up calculations - create formulas, then print results, and then coloration for quarterly change

For i = 2 To lastRow

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

total = total + Cells(i, 7).Value

quarter = Cells(i, 6) - Cells(start, 3)

percent = quarter / Cells(start, 3)

start = i + 1

Range("I" & 2 + j).Value = Cells(i, 1).Value
Range("J" & 2 + j).Value = quarter
Range("K" & 2 + j).Value = percent
Range("K" & 2 + j).NumberFormat = "0.00%"

Range("L" & 2 + j).Value = total

Select Case quarter
                    Case Is > 0
                        Range("J" & 2 + j).Interior.ColorIndex = 4
                    Case Is < 0
                        Range("J" & 2 + j).Interior.ColorIndex = 3
                    Case Else
                        Range("J" & 2 + j).Interior.ColorIndex = 0
                        End Select

'Reset the variable (end loop)
total = 0
quarter = 0
j = j + 1


'Then find the min and max

End If

Next i

    'Second loop for Second Leaderboard
    Dim max_price As Double
    Dim min_price As Double
    Dim max_volume As LongLong
    Dim max_price_stock As String
    Dim min_price_stock As String
    Dim max_volume_stock As String

    
 

    'initialization of first row
    max_price = Cells(2, 11).Value
    min_price = Cells(2, 11).Value
    max_volume = Cells(2, 12).Value
    max_price_stock = Cells(2, 9).Value
    max_price_stock = Cells(2, 9).Value
    max_price_stock = Cells(2, 9).Value
    
    For j = 2 To leaderboard_row
        If (Cells(j, 11).Value > max_price) Then
        max_price = Cells(j, 11).Value
        max_price_stock = Cells(j, 9).Value
               
        End If
        If (Cells(j, 12).Value > max_volume) Then
        max_price = Cells(j, 12).Value
        max_price_stock = Cells(j, 9).Value
        End If
    

    Next j
    
    'Write out to Excel Workbook
    Range("O2").Value = max_price_stock
    Range("O3").Value = min_price_stock
    Range("O4").Value = max_volume_stock
    
    Range("P2").Value = max_price_stock
    Range("P3").Value = min_price_stock
    Range("P4").Value = max_volume_stock
    


End Sub
