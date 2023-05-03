Sub changeTabNames()

    Dim wsD As Worksheet
    Dim txtDate As String
    Dim i As Integer
           
    txtDate = InputBox("Starting date: ")
    
    Set wsD = ThisWorkbook.Worksheets("Admin")
    
    wsD.Range("D1").Value = CDate(txtDate)
    For i = 4 To 7
        Sheets(i).name = Left(Sheets(i).name, 21) & wsD.Range("E1").Text & wsD.Range("F1").Text & Right(Sheets(i).name, 2)
    Next i
    For i = 8 To 11
        Sheets(i).name = Left(Sheets(i).name, 21) & wsD.Range("E2").Text & wsD.Range("F2").Text & Right(Sheets(i).name, 2)
    Next i
    For i = 12 To 15
        Sheets(i).name = Left(Sheets(i).name, 21) & wsD.Range("E3").Text & wsD.Range("F3").Text & Right(Sheets(i).name, 2)
    Next i
    For i = 16 To 19
        Sheets(i).name = Left(Sheets(i).name, 21) & wsD.Range("E4").Text & wsD.Range("F4").Text & Right(Sheets(i).name, 2)
    Next i
    For i = 20 To 23
        Sheets(i).name = Left(Sheets(i).name, 21) & wsD.Range("E5").Text & wsD.Range("F5").Text & Right(Sheets(i).name, 2)
    Next i
    
End Sub