Attribute VB_Name = "Module2"
Private Sub display()
k = 0

flex2.Cols = 7
flex2.Rows = 1
flex2.Row = k
flex2.Col = 0

flex2.Text = "DATE"
flex2.Col = 1
flex2.Text = "ENO"
flex2.Col = 2
flex2.Text = "NAME"
flex2.Col = 3
flex2.Text = "2000"
flex2.Col = 4
flex2.Text = "1000"
flex2.Col = 5
flex2.Text = "500"
flex2.Col = 6
flex2.Text = "100"
flex2.Col = 7
flex2.Text = "50"
flex2.Col = 8
flex2.Text = "20"
flex2.Col = 9
flex2.Text = "10"
flex2.Col = 10
flex2.Text = "5"
flex2.Col = 11
flex2.Text = "COINS"
flex2.Col = 12
flex2.Text = "TOTAL"

j = 0
k = k + 1
'Dim da1, da2, mo1, mo2, ye1, ye2 As Integer
'da1 = Day(D1)
'mo1 = Month(D1)
'ye1 = Year(D1)
'da2 = Day(d2)
'mo2 = Month(d2)
'ye2 = Year(d2)

rs2.Open "Select * from ventab where day(date) >=" & da1 _
                                    & " and day(date) <=" & da2 _
                                    & " and month(date) >=" & mo1 _
                                    & " and month(date) <=" & mo2 _
                                    & " and year(date) >=" & ye1 _
                                    & " and year(date) <=" & ye2 _
                                    , con, adOpenDynamic, adLockOptimistic


While rs2.EOF <> True

flex2.Rows = flex2.Rows + 1
flex2.Row = flex2.Row + 1
flex2.Col = j

flex2.Text = rs2.Fields(0)
flex2.Col = flex2.Col + 1
flex2.Text = rs2.Fields(1)
flex2.Col = flex2.Col + 1
flex2.Text = rs2.Fields(2)
flex2.Col = flex2.Col + 1
flex2.Text = rs2.Fields(3)
flex2.Col = flex2.Col + 1
flex2.Text = rs2.Fields(4)
flex2.Col = flex2.Col + 1
flex2.Text = rs2.Fields(5)
flex2.Col = flex2.Col + 1
flex2.Text = rs2.Fields(6)
flex2.Col = flex2.Col + 1
flex2.Text = rs2.Fields(7)
flex2.Col = flex2.Col + 1
flex2.Text = rs2.Fields(8)
flex2.Col = flex2.Col + 1
flex2.Text = rs2.Fields(9)
flex2.Col = flex2.Col + 1
flex2.Text = rs2.Fields(10)
flex2.Col = flex2.Col + 1
flex2.Text = rs2.Fields(11)
flex2.Col = flex2.Col + 1
flex2.Text = rs2.Fields(12)
k = k + 1
rs2.MoveNext

Wend


End Sub

