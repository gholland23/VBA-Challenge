Sub answersyo()

Dim ws As Worksheet
'Set Workbook = ActiveWorkbook'
For Each ws In Worksheets
'ws.Activate'

    'first things first, I'm the realist'
    ws.Range("K1").Value = "ticker"
    ws.Range("r1").Value = "ticker"
    ws.Range("s1").Value = "value"
    ws.Range("L1").Value = "yearly change"
    ws.Range("M1").Value = "percent change"
    ws.Range("n1").Value = "total volume"
    

    'beginning code...im breaking it into parts'
    
    Dim ticker As String
    Dim lastrow As Long
    Dim summarytablerow As Integer
    summarytablerow = 2
    Dim volumetotal As Double
    volumetotal = 0
    
    'secondpart variables'
    Dim yearlyopen As Double
    Dim yearlyclose As Double
    Dim yearlychange As Double
    Dim previousprice As Long
    previousprice = 2
    Dim percentchange As Double

    'third part'
    Dim greatestincrease As Double
    ws.Range("q2").Value = "greatest % increase"
    Dim greatesttotalvolume As Double
    ws.Range("q3").Value = "greatest total volume"
    Dim greatestdecrease As Double
    ws.Range("q4").Value = "greatest % decrease"
    'ticker for greatest %"
    Dim arbitrary As Integer
    'ticker for greatest total volume'
    Dim arbitrary2 As Integer
    'ticker for greatest % decrease'
    Dim arbitrary3 As Integer
    '2ndsheet'
    Dim arbitrary4 As Integer
    Dim arbitrary5 As Integer
    Dim arbitrary6 As Integer
    '3rdsheet'
    Dim arbitrary7 As Integer
    Dim arbitrary8 As Integer
    Dim arbitrary9 As Integer
    
    




    'main code'
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastrow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ticker = ws.Cells(i, 1).Value
            volumetotal = volumetotal + ws.Cells(i, 7).Value
            ws.Range("k" & summarytablerow).Value = ticker
            ws.Range("n" & summarytablerow).Value = volumetotal
            summarytablerow = summarytablerow + 1
            volumetotal = 0
        
        Else
            volumetotal = volumetotal + ws.Cells(i, 7).Value
        End If
        
            
        
        'open and closing stuff'
        
        yearlyopen = ws.Range("c" & previousprice)
        yearlyclose = ws.Range("f" & i)
        yearlychange = yearlyopen - yearlyclose
        ws.Range("l" & summarytablerow).Value = yearlychange
        
        If yearlyopen = 0 Then
            percentchange = 0
        Else
            yearlyopen = ws.Range("c" & previousprice)
            percentchange = yearlychange / yearlyopen
            ws.Range("m" & summarytablerow).Value = percentchange
            ws.Range("m" & summarytablerow).NumberFormat = "0.00%"
        End If
        
        If ws.Range("l2" & summarytablerow).Value >= 0 Then
            ws.Range("l2" & summarytablerow).Interior.ColorIndex = 4
        Else
            ws.Range("l2" & summarytablerow).Interior.ColorIndex = 3
        End If
        
        
    Next i
     
Next ws


    Range("s2") = Application.WorksheetFunction.Max(Range("sheet1!M1:M" & lastrow))
    Range("s2").NumberFormat = "0.00%"
    Range("s3") = Application.WorksheetFunction.Max(Range("sheet1!N1:N" & lastrow))
    Range("s4") = Application.WorksheetFunction.Min(Range("sheet1!M1:M" & lastrow))
    Range("s4").NumberFormat = "0.00%"
    
    arbitrary = Application.Match(Range("s2").Value, Range("sheet1!M1:M" & lastrow), 0)
    Range("r2").Value = Cells(arbitrary, 11)
    arbitrary2 = Application.Match(Range("s3").Value, Range("sheet1!N1:N" & lastrow), 0)
    Range("r3").Value = Cells(arbitrary2, 11)
    arbitrary3 = Application.Match(Range("s4").Value, Range("sheet1!M1:M" & lastrow), 0)
    Range("r4").Value = Cells(arbitrary3, 11)
    
    Worksheets("sheet2").Range("s2") = Application.WorksheetFunction.Max(Worksheets("sheet2").Range("sheet2!M1:M" & lastrow))
    Worksheets("sheet2").Range("s2").NumberFormat = "0.00%"
    Worksheets("sheet2").Range("s3") = Application.WorksheetFunction.Max(Worksheets("sheet2").Range("sheet2!N1:N" & lastrow))
    Worksheets("sheet2").Range("s4") = Application.WorksheetFunction.Min(Worksheets("sheet2").Range("sheet2!M1:M" & lastrow))
    Worksheets("sheet2").Range("s4").NumberFormat = "0.00%"
    
    arbitrary4 = Application.Match(Worksheets("sheet2").Range("s2").Value, Worksheets("sheet2").Range("sheet2!M1:M" & lastrow), 0)
    Worksheets("sheet2").Range("r2").Value = Worksheets("sheet2").Cells(arbitrary4, 11)
    arbitrary5 = Application.Match(Worksheets("sheet2").Range("s3").Value, Worksheets("sheet2").Range("sheet2!N1:N" & lastrow), 0)
    Worksheets("sheet2").Range("r3").Value = Worksheets("sheet2").Cells(arbitrary5, 11)
    arbitrary6 = Application.Match(Worksheets("sheet2").Range("s4").Value, Worksheets("sheet2").Range("sheet2!M1:M" & lastrow), 0)
    Worksheets("sheet2").Range("r4").Value = Worksheets("sheet2").Cells(arbitrary6, 11)

    Worksheets("sheet3").Range("s2") = Application.WorksheetFunction.Max(Worksheets("sheet3").Range("sheet3!M1:M" & lastrow))
    Worksheets("sheet3").Range("s2").NumberFormat = "0.00%"
    Worksheets("sheet3").Range("s3") = Application.WorksheetFunction.Max(Worksheets("sheet3").Range("sheet3!N1:N" & lastrow))
    Worksheets("sheet3").Range("s4") = Application.WorksheetFunction.Min(Worksheets("sheet3").Range("sheet3!M1:M" & lastrow))
    Worksheets("sheet3").Range("s4").NumberFormat = "0.00%"

    arbitrary7 = Application.Match(Worksheets("sheet3").Range("s2").Value, Worksheets("sheet3").Range("sheet3!M1:M" & lastrow), 0)
    Worksheets("sheet3").Range("r2").Value = Worksheets("sheet3").Cells(arbitrary7, 11)
    arbitrary8 = Application.Match(Worksheets("sheet3").Range("s3").Value, Worksheets("sheet3").Range("sheet3!N1:N" & lastrow), 0)
    Worksheets("sheet3").Range("r3").Value = Worksheets("sheet3").Cells(arbitrary8, 11)
    arbitrary9 = Application.Match(Worksheets("sheet3").Range("s4").Value, Worksheets("sheet3").Range("sheet3!M1:M" & lastrow), 0)
    Worksheets("sheet3").Range("r4").Value = Worksheets("sheet3").Cells(arbitrary9, 11)

End Sub


