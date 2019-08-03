Attribute VB_Name = "Module1"
Sub stockdata_easy()
    
    ' Specify column of interest and start at row 2. Start total volume at 0.
    
    Dim column As Integer
    column = 1
    
    Dim row As Integer
    row = 2
    
    Dim totalvolume As Double
    totalvolume = 0
    
    
    ' Loop through all rows in sheet
    
    For i = 2 To 760192
    
    
    ' Add up total volume for ticker and print when loop reaches next ticker
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            totalvolume = totalvolume + Cells(i, 7).Value
    
            Cells(row, 10).Value = totalvolume
            Cells(row, 9).Value = Cells(i, 1).Value
            
            
    ' Reset total volume and move to next row
    
            row = row + 1
            totalvolume = 0
    
    
    ' Keep adding to total volume if ticker is the same
    
        Else
            totalvolume = totalvolume + Cells(i, 7).Value
    
    End If
    Next i


End Sub



Sub stockdata_moderate()

    ' Specify column of interest and start at row 2. Start total volume at 0.
    
    Dim column As Integer
    column = 1
    
    Dim row As Integer
    row = 2
    
    
    ' Create variables for opening price, closing price, difference, and percent difference
    Dim openingprice As Double
    Dim closingprice As Double
    Dim difference As Double
    Dim percentdifference As Double
    
    
    ' Start opening price at first ticker's first row and rest of variables at 0
    openingprice = Cells(2, 3).Value
    closingprice = 0
    difference = 0
    percentdifference = 0
    
    
    ' Loop through all rows in sheet
    
    For i = 2 To 760192
    
    
    ' Calculate difference and percent difference between closing and opening and print when loop reaches next ticker
    
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    closingprice = Cells(i, 6).Value
    difference = closingprice - openingprice
    percentdifference = (closingprice - openingprice) / openingprice
    Cells(row, 11).Value = difference
    Cells(row, 12).Value = percentdifference
    
    
    ' Reset value of variables, specify first opening price for next ticker, and move to next row
    
    closingprice = 0
    difference = 0
    percentdifference = 0
    openingprice = Cells(i + 1, 3).Value
    row = row + 1
    End If
    
    
    ' Conditional formatting for yearly change (highlight cell green if positive, red if negative)
    
    If Cells(row, 11).Value < 0 Then
    Cells(row, 11).Interior.ColorIndex = 3
    Else
    Cells(row, 11).Interior.ColorIndex = 4
    
    End If
    Next i
    
End Sub

