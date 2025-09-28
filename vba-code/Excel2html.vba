Sub GenerateStyledHTMLFilesWithPreAndIndexLink()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim linkName As String
    Dim htmlContent As String
    Dim filePath As String
    Dim indexContent As String
    Dim basePath As String

    Set ws = ThisWorkbook.Sheets(1)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    basePath = ThisWorkbook.Path & "\"

    indexContent = "<!DOCTYPE html><html><head><meta charset='UTF-8'>" & _
                   "<title>Index</title>" & _
                   "<style>" & _
                   "body { font-family: sans-serif; margin: 40px; background-color: #fff; color: #000; }" & _
                   "h1 { font-size: 24px; border-bottom: 2px solid #ccc; padding-bottom: 5px; }" & _
                   "ul { list-style-type: none; padding-left: 0; }" & _
                   "li { margin: 8px 0; }" & _
                   "a { color: #00f; text-decoration: underline; }" & _
                   "</style></head><body>" & _
                   "<h1>Index</h1><ul>"

    For i = 1 To lastRow
        linkName = Trim(ws.Cells(i, 1).Value)
        htmlContent = ws.Cells(i, 2).Value

        If linkName <> "" Then
            filePath = basePath & linkName & ".html"
            Dim pageContent As String
            pageContent = "<!DOCTYPE html><html><head><meta charset='UTF-8'>" & _
                          "<title>" & linkName & "</title>" & _
                          "<style>" & _
                          "body { font-family: sans-serif; margin: 40px; background-color: #fff; color: #000; }" & _
                          "h1 { font-size: 24px; border-bottom: 2px solid #ccc; padding-bottom: 5px; }" & _
                          "pre { font-size: 14px; background-color: #f9f9f9; padding: 10px; border: 1px solid #ccc; white-space: pre-wrap; }" & _
                          "a { color: #00f; text-decoration: underline; }" & _
                          "</style></head><body>" & _
                          "<div><a href='index.htm'>[Index]</a></div>" & _
                          "<h1>" & linkName & "</h1>" & _
                          "<pre>" & htmlContent & "</pre>" & _
                          "</body></html>"
            WriteTextFile filePath, pageContent

            indexContent = indexContent & "<li><a href='" & linkName & ".html'>" & linkName & "</a></li>"
        End If
    Next i

    indexContent = indexContent & "</ul></body></html>"
    WriteTextFile basePath & "index.htm", indexContent

    MsgBox "HTML files generation completed", vbInformation
End Sub

Private Sub WriteTextFile(filePath As String, content As String)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    With stream
        .Charset = "UTF-8"
        .Open
        .WriteText content
        .SaveToFile filePath, 2
        .Close
    End With
End Sub
