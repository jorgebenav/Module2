Sub ModelTwo()
    Dim Company_name As String, LastCompany As String
    Dim Vol As Double, Company_total As Double
    Dim ws As Worksheet
    Dim LastRow As Long, outputRow As Long, secondoutputRow As Long
    Dim i As Long
    Dim sheetNames As Variant, name As Variant
    Dim OpenPrice As Double, ClosePrice As Double
    Dim YCH As Double, PercentCH As Double

    'The sheets in the book are referred by their name
    'Then For Each statements is made
    sheetNames = Array("A", "B", "C", "D", "E", "F")
    For Each name In sheetNames
        Set ws = ThisWorkbook.Sheets(name)
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        outputRow = 2
        secondoutputRow = 2
        Company_total = 0
        Vol = 0
        LastCompany = ""
          
        For i = 2 To LastRow
            Company_name = ws.Cells(i, 1).Value
            If Company_name <> LastCompany And i <> 2 Then
                ' Output totals from previous company
                ws.Range("K" & outputRow).Value = LastCompany
                ws.Range("N" & outputRow).Value = Company_total
                YCH = ClosePrice - OpenPrice
                If OpenPrice <> 0 Then
                    PercentCH = (YCH / OpenPrice) * 100
                    ws.Range("L" & outputRow).Value = YCH
                    ws.Range("M" & outputRow).Value = PercentCH
                End If
                outputRow = outputRow + 1
                
                ' Reset variables for new company
                Company_total = 0
                OpenPrice = 0
                ClosePrice = 0
            End If
            
            ' Accumulate volume
            Company_total = Company_total + ws.Cells(i, 7).Value
            
            ' Set first and last prices
            If OpenPrice = 0 Then
                OpenPrice = ws.Cells(i, 3).Value
            End If
            ClosePrice = ws.Cells(i, 6).Value
            
            LastCompany = Company_name
            
            ' Ensure to output last company's data after loop
            If i = LastRow Then
                ws.Range("K" & outputRow).Value = Company_name
                ws.Range("N" & outputRow).Value = Company_total
                YCH = ClosePrice - OpenPrice
                If OpenPrice <> 0 Then
                    PercentCH = (YCH / OpenPrice) * 100
                    ws.Range("L" & outputRow).Value = YCH
                    ws.Range("M" & outputRow).Value = PercentCH
                End If
            End If
        Next i
        For i = 2 To LastRow
            If ws.Cells(i, 12).Value >= 0 Then
                ws.Cells(i, 12).Interior.ColorIndex = 4
            Else
                ws.Cells(i, 12).Interior.ColorIndex = 3
            End If
            If GRTI < ws.Cells(i, 12).Value Then
                GRTI = ws.Cells(i, 12).Value
                GRTIN = ws.Cells(i, 11).Value
            End If
            If GRTD >= ws.Cells(i, 12).Value Then
                GRTD = ws.Cells(i, 12).Value
                GRTDN = ws.Cells(i, 11).Value
            End If
            If GTV < ws.Cells(i, 14).Value Then
                GTV = ws.Cells(i, 14).Value
                GTVN = ws.Cells(i, 11).Value
            End If
            ws.Cells(2, 18).Value = GRTI
            ws.Cells(3, 18).Value = GRTD
            ws.Cells(4, 18).Value = GTV
            ws.Cells(2, 19).Value = GRTIN
            ws.Cells(3, 19).Value = GRTDN
            ws.Cells(4, 19).Value = GTVN
            
        Next i
        ws.Cells(1, 11).Value = "Company"
        ws.Cells(1, 12).Value = "Yearly Change"
        ws.Cells(1, 13).Value = "Percent Change"
        ws.Cells(1, 14).Value = "Total Volume"
        ws.Cells(2, 17).Value = "Great % Increase"
        ws.Cells(3, 17).Value = "Great % Decrease"
        ws.Cells(4, 17).Value = "Great total Vol."
        ws.Cells(1, 18).Value = "Ticker"
        ws.Cells(1, 19).Value = "Value"
    Next name
    
End Sub


