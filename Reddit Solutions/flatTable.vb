Sub flatTable()
    Dim ws As Worksheet, ws2 As Worksheet
    Dim lRow As Integer, r As Integer, c As Integer, pRow As Integer, pCol As Integer
    Dim prodNum As Variant, prodName As Variant, size As Variant, quant As Variant, titles() As Variant, t As Variant
    
    Application.ScreenUpdating = False
    
    Set ws = ActiveSheet
    Worksheets.Add(After:=ws).Name = "Flat_Table"
    Set ws2 = Worksheets("Flat_Table")
    
    ws.Activate
    lRow = ws.Range("A" & Rows.Count).End(xlUp).Row
    pRow = 2
    
    titles = Array("Product Number", "Product Name", "Size", "Quantity")
    
    ws2.Activate
    pCol = 1
        For Each t In titles
            With ws2.Cells(1, pCol)
                .Value = t
            End With
            pCol = pCol + 1
        Next t
            
    pRow = 2
    
    ws.Activate
    For r = 2 To lRow
        prodNum = ws.Range("A" & r).Value
        prodName = ws.Range("B" & r).Value
        If Not prodNum = "size" And Not prodNum = "quantity" Then
            c = 2
            Do Until ws.Cells(r + 1, c).Value = ""
                size = ws.Cells(r + 1, c).Value
                quant = ws.Cells(r + 2, c).Value
                pCol = 1
                titles = Array(prodNum, prodName, size, quant)
                For Each t In titles
                    With ws2.Cells(pRow, pCol)
                        .Value = t
                    End With
                    pCol = pCol + 1
                Next t
                pRow = pRow + 1
                Erase titles
                c = c + 1
            Loop
        End If
    Next r
End Sub