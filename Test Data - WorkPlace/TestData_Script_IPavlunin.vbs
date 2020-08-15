Sub stockdata()

    ' define variables
    Dim total As Double
    Dim ticker As String
    Dim rownumber As Integer
    Dim yearlychange As Double
    Dim percentchange As Double
    Dim openprice As Double
    Dim closeprice As Double
    Dim maxtotal As Double
    Dim maxpercent As Double
    Dim minpercent As Double
    Dim totalincrease As String
    Dim maxincrease As String
    Dim minincrease As String
    Dim ws As Worksheet
    
    
    ' loop for all worksheets
    For Each ws In Worksheets
    
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        total = 0
        ' 1st row to calculate
        rownumber = 2

        ' define openprice as open year value
        openprice = ws.Cells(2, 3).Value

        ' column headings
        ws.Range("I1,P1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ' loop from 2nd to the last row
            For i = 2 To LastRow
                
                ' stock change checker
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    
                    ' define closeprice as close year value
                    closeprice = ws.Cells(i, 6).Value
                    
                    ' calculate yearchange
                    yearlychange = closeprice - openprice
                    ws.Range("J" & rownumber).Value = yearlychange
                    
                    ' define stockname as a ticker
                    ticker = ws.Cells(i, 1).Value

                    ' calculate total stock volume
                    total = total + Cells(i, 7).Value

                    ' printing out total stock values
                    ws.Range("L" & rownumber).Value = total
                    ws.Range("L" & rownumber).NumberFormat = "0"
                            
                            ' conditional formatting
                            If yearlychange < 0 Then
                            
                            ws.Cells(rownumber, 10).Interior.ColorIndex = 3
                            
                            Else
                            
                            ws.Cells(rownumber, 10).Interior.ColorIndex = 4
                            
                            End If
                            
                            ' calculate oercentage change and fixing /0 error
                            If openprice <> 0 Then
                            
                            percentchange = yearlychange / openprice
                            
                            Else
                            
                            percentchange = 0
                            
                            End If

                    ' printing out column values and format       
                    ws.Range("K" & rownumber).Value = percentchange
                    ws.Range("K" & rownumber).NumberFormat = "0.00%"
                    ws.Range("I" & rownumber).Value = ticker
                    
                    ' nest loop
                    rownumber = rownumber + 1
                    
                    ' zeroing total for next loop
                    total = 0
                    
                    openprice = ws.Cells(i + 1, 3).Value
                    
                Else
                
                    total = total + ws.Cells(i, 7).Value
                    
                End If
                
            Next i
            
    Next ws
        
        ' loop for summary table
        For Each ws In Worksheets
        
            ' headings and formating
            ws.Range("O2").Value = "Greatest % Increase"
            ws.Range("O3").Value = "Greatest % Decrease"
            ws.Range("O4").Value = "Greatest Total Volume"
            ws.Range("Q2").NumberFormat = "0.00%"
            ws.Range("Q3").NumberFormat = "0.00%"
            ws.Range("Q4").NumberFormat = "0.0000E+00"
    
            maxtotal = WorksheetFunction.Max(ws.Range("L:L"))
            
            maxpercent = WorksheetFunction.Max(ws.Range("K:K"))
            
            minpercent = WorksheetFunction.Min(ws.Range("K:K"))
            
            For i = 2 To 3200
            
                If ws.Cells(i, 12).Value = maxtotal Then
            
                totalincrease = ws.Cells(i, 12).Offset(0, -3).Value
            
                ElseIf ws.Cells(i, 11).Value = maxpercent Then
            
                maxincrease = ws.Cells(i, 11).Offset(0, -2).Value
            
                ElseIf ws.Cells(i, 11).Value = minpercent Then
            
                minincrease = ws.Cells(i, 11).Offset(0, -2).Value
            
                End If
            
            Next i
            
            ' printing out summary table values
            ws.Range("Q2").Value = maxpercent
            ws.Range("Q3").Value = minpercent
            ws.Range("Q4").Value = maxtotal
    
            ws.Range("P2").Value = maxincrease
            ws.Range("P3").Value = minincrease
            ws.Range("P4").Value = totalincrease
    
    Next ws

End Sub