Attribute VB_Name = "Module1"

Sub SortStock()
Dim row As Double
Dim tablerow As Integer
Dim stock As Double
Dim opening As Double
Dim closing As Double
Dim compare As Double
Dim big As Double ' carry the greatest %
Dim small As Double  'carry the greatest -%
Dim total As Double
Dim big_ticker As String
Dim small_ticker As String
Dim total_ticker As String

'-----------------
Dim i As Integer
Dim ws_num As Integer

Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet
ws_num = ThisWorkbook.Worksheets.Count

For i = 1 To ws_num
    ThisWorkbook.Worksheets(i).Activate
'-----------------

                row = Cells(Rows.Count, 1).End(xlUp).row

                tablerow = 1

                Range("I1,P1").Value = "Ticker"
                Range("J1").Value = "Yearly Change"
                Range("K1").Value = "Percentage Change"
                Range("L1").Value = "Total Stock Volumne"
                Range("O2").Value = "Greatest % Increase"
                Range("O3").Value = "Greatest % Decrease"
                Range("O4").Value = "Greatest Total Volume"
                Range("Q1").Value = "Value"



                    For x = 2 To row

                        ' Compare Column A if the below cell if the same or not
                                    If Cells(x + 1, 1) = Cells(x, 1) Then

                                        ' if its the same, it sums up the stock with the previous amount
                                        stock = stock + Cells(x, 7)
                                                    If Cells(x, 1) <> Cells(x - 1, 1) Then ' comparing the ticker with ticker on top, if different, grab the opening
                                                    opening = Cells(x, 3)
                                                    End If

                                    'when the 2 cells are not equal (does not meet the = requirement)
                                    Else:
                                    ' if 2 cells are different, move to next row in the table
                                        tablerow = tablerow + 1

                                        ' print last pick of x
                                        Cells(tablerow, 9) = Cells(x, 1) ' Print the ticker name before it does not meet = requirement
                                        Cells(tablerow, 12) = stock + Cells(x, 7) ' the stock vol including the last stock vol before the unequal condition
                                        stock = 0
                                        Cells(tablerow, 10) = Cells(x, 6) - opening '"opening:" + Str(opening) + "closing:" + Str(Cells(x, 6)) 'closing =cells(x,6)

                                        'Cells(tablerow, 13) = "opening:" + Str(opening) + " closing:" + Str(Cells(x, 6)) 'closing =cells(x,6)
                                         If opening <> 0 Then Cells(tablerow, 11) = (Cells(x, 6) - opening) / opening

                                            If Cells(tablerow, 10) < 0 Then Cells(tablerow, 10).Interior.ColorIndex = 3 'check changes on the fly

                                            If Cells(tablerow, 10) > 0 Then Cells(tablerow, 10).Interior.ColorIndex = 4

                                    End If

                    Next x

                big = Range("K2")
                small = Range("k2")
                total = Range("L2")

                For g = 2 To tablerow
                    If Cells(g + 1, 11) > big Then big = Cells(g + 1, 11): big_ticker = Cells(g + 1, 9)
                    If Cells(g + 1, 11) < small Then small = Cells(g + 1, 11): small_ticker = Cells(g + 1, 9)
                    If Cells(g + 1, 12) > total Then total = Cells(g + 1, 12): total_ticker = Cells(g + 1, 9)
                Next g

                Range("P2") = big_ticker
                Range("p3") = small_ticker
                Range("P4") = total_ticker
                Range("Q2") = big
                Range("Q3") = small
                Range("Q4") = total

                Range("K:K").NumberFormat = "0.00%"
                Range("Q2,Q3").NumberFormat = "0.00%"



 '-------------
Next


starting_ws.Activate
'-------------

End Sub


