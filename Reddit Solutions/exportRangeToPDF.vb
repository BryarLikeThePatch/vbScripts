Sub ToPDF()
    Dim rng as Range
    Dim filepath As String, filename as String
    
    Set rng = ActiveSheet.Range("A1:Z99") 'whatever your range is
    filename = Format(Now(), "mm_dd_yyyy") & "_" & ActiveWorkbook.Name & ".pdf"
    filepath = ActiveWorkbook.Path & "\" & filename
        'or something like "C:\Username\Personal\PDFs\Export" to have a fixed filepath
        'trade "\" for "/" if you're rocking a unix system like Mac or Gnu/Linux
    
    
    rng.ExportAsFixedFormat Type=xlTypePDF, Filename:=filepath, OpenAfterPublish:=True

End Sub
