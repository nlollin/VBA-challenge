Attribute VB_Name = "Module1"
Sub ticker()

Dim ws As Worksheet

For Each ws In Worksheets

    Dim i As Long
    Dim x As Long
    x = 2

    For i = 2 To 759002

        ws.Cells(x, 9).Value = ws.Cells(i, 1).Value
        
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then x = x + 1
        
    
    Next i
    
    Dim yearlychange As Double
    Dim start As Double
    Dim finish As Double
    Dim percentchange As Double
    Dim y As Long
    Dim z As Long

    y = 2
    z = 1
    start = ws.Cells(y, 3).Value

    For i = 3 To 759002

        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then finish = ws.Cells(i - 1, 6).Value
        yearlychange = finish - start
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then z = z + 1
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then ws.Cells(z, 10) = yearlychange
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then percentchange = yearlychange / start
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then ws.Cells(z, 11) = percentchange
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then start = ws.Cells(i, 3).Value

    Next i

    Dim total As LongLong
    Dim w As Long
    w = 2
    total = 0

    For i = 2 To 759002

        total = total + ws.Cells(i, 7).Value

        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then ws.Cells(w, 12).Value = total
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then w = w + 1
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then total = 0

    Next i

    Dim maxpercent As Double
    Dim minpercent As Double
    Dim maxvolume As LongLong
    Dim a As String
    Dim b As String
    Dim c As String

    maxpercent = ws.Cells(2, 11).Value
    minpercent = ws.Cells(2, 11).Value
    maxvolume = ws.Cells(2, 12).Value
    a = ws.Cells(2, 9).Value
    b = a
    c = a

    For i = 3 To 759002

        If ws.Cells(i, 11).Value > maxpercent Then a = ws.Cells(i, 9).Value
        If ws.Cells(i, 11).Value > maxpercent Then maxpercent = ws.Cells(i, 11).Value
        If ws.Cells(i, 11).Value < minpercent Then b = ws.Cells(i, 9).Value
        If ws.Cells(i, 11).Value < minpercent Then minpercent = ws.Cells(i, 11).Value
        If ws.Cells(i, 12).Value > maxvolume Then c = ws.Cells(i, 9).Value
        If ws.Cells(i, 12).Value > maxvolume Then maxvolume = ws.Cells(i, 12).Value

    Next i

    ws.Cells(2, 16).Value = a
    ws.Cells(2, 17).Value = maxpercent
    ws.Cells(3, 16).Value = b
    ws.Cells(3, 17).Value = minpercent
    ws.Cells(4, 16).Value = c
    ws.Cells(4, 17).Value = maxvolume

Next ws

End Sub



