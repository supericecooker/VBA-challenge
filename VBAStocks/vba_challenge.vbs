Sub alphatest()


    ' LOOP THROUGH ALL SHEETS


Dim WS As Worksheet

    For Each WS In ActiveWorkbook.Worksheets

    WS.Activate

        ' set up last row code

        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row



        ' creating column name

        Cells(1, "k").Value = "Ticker"
Range("K1").Font.Bold = True
Range("l1").Font.Bold = True
Range("m1").Font.Bold = True
Range("n1").Font.Bold = True


        Cells(1, "l").Value = "Yearly Change"

        Cells(1, "m").Value = "Percent Change"

        Cells(1, "n").Value = "Total Stock Volume"

        

        'Set up column name as variables

        Dim OpenP As Double

        Dim CloseP As Double

        Dim YearC As Double

        Dim TickerN As String

        Dim PercentC As Double

        Dim Total_StockVolume As Single

        StockVolume = 0
       
 'starting setting up row and column for nested condition. double for large number with decimal, integer for round figure, single for decimal.
 

        Dim Row As Double

        'row is 2 because 1st row is title
        Row = 2
        
        Dim Column As Integer

        Column = 1
'long holds and store integer.


        Dim i As Long


        'set up Open price navigate @C2 cell.


        OpenP = Cells(2, Column + 2).Value


         'start Loop through ticker


        For i = 2 To LastRow

         'Check if we are still within the same ticker symbol, if it is not...

        If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then


         'Set up new Ticker column K


        TickerN = Cells(i, Column).Value

        Cells(Row, Column + 10).Value = TickerN

         'call Close Price a name column F

         CloseP = Cells(i, Column + 5).Value

          'calculate Yearly Change from OP to CP

         YearC = CloseP - OpenP

           'assign calculated value to yearly change column L

         Cells(Row, Column + 11).Value = YearC


            'conidtioning percent change column M

         If (OpenP = 0 And CloseP = 0) Then

         PercentC = 0

         ElseIf (OpenP = 0 And CloseP <> 0) Then

         PercentC = 1

         Else

                    PercentC = YearC / OpenP

                    Cells(Row, Column + 12).Value = PercentC

                    Cells(Row, Column + 12).NumberFormat = "0.00%"

             End If


                'calculate Total stock Volumn

                StockVolume = Volume + Cells(i, Column + 6).Value

                'assigned column N

                Cells(Row, Column + 13).Value = StockVolume

                ' Add one to the summary table row

                Row = Row + 1

                ' reset the Open Price

                OpenP = Cells(i + 1, Column + 2)

                ' reset the Volumn Total

                StockVolume = 0

            'if cells are the same ticker

            Else

                StockVolume = Volume + Cells(i, Column + 6).Value

            End If
 'end loop
        Next i

        

        ' Determine the Last Row of Yearly Change for each ws column K

        YearCLastRow = WS.Cells(Rows.Count, Column + 11).End(xlUp).Row

        'condition the Cell Colors

        For j = 2 To YearCLastRow

            If (Cells(j, Column + 11).Value > 0 Or Cells(j, Column + 11).Value = 0) Then

                Cells(j, Column + 11).Interior.ColorIndex = 4

            ElseIf Cells(j, Column + 11).Value < 0 Then

                Cells(j, Column + 11).Interior.ColorIndex = 3

            End If

        Next j

        

        
        

    Next WS

        

End Sub
