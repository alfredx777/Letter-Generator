Option Explicit

Sub GenerateLetters()
    Dim ws As Worksheet
    Dim wdApp As Object, wdDoc As Object
    Dim basePath As String, templatePath As String, outRoot As String
    Dim folderPath As String, filePath As String, baseName As String
    Dim lastRow As Long, r As Long, made As Long

    Set ws = ThisWorkbook.Worksheets("Names")
    basePath = ThisWorkbook.Path
    templatePath = basePath & "\LetterTemplate.dotx"
    outRoot = basePath & "\Output"

    If basePath = "" Then
        MsgBox "Save the workbook first.", vbExclamation
        Exit Sub
    End If
    If Dir(templatePath) = "" Then
        MsgBox "LetterTemplate.dotx not found in:" & vbCrLf & basePath, vbExclamation
        Exit Sub
    End If
    If Dir(outRoot, vbDirectory) = "" Then MkDir outRoot

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    On Error GoTo Fail

    For r = 2 To lastRow
        If Trim(ws.Cells(r, 4).Text) <> "" Then
            baseName = CleanName(ws.Cells(r, 4).Text) & "_" & CleanName(ws.Cells(r, 3).Text)
            folderPath = outRoot & "\" & baseName
            If Dir(folderPath, vbDirectory) = "" Then MkDir folderPath

            Set wdDoc = wdApp.Documents.Add(Template:=templatePath)
            ReplaceTag wdDoc, "<<RefNo>>", ws.Cells(r, 1).Text
            ReplaceTag wdDoc, "<<Title>>", ws.Cells(r, 2).Text
            ReplaceTag wdDoc, "<<FirstName>>", ws.Cells(r, 3).Text
            ReplaceTag wdDoc, "<<Surname>>", ws.Cells(r, 4).Text
            ReplaceTag wdDoc, "<<Date>>", ws.Cells(r, 5).Text

            filePath = folderPath & "\" & baseName & "_Letter.docx"
            wdDoc.SaveAs2 filePath, 16   ' 16 = .docx
            wdDoc.Close False
            Set wdDoc = Nothing
            made = made + 1
        End If
    Next r

    wdApp.Quit
    MsgBox made & " folders and letters created in:" & vbCrLf & outRoot, vbInformation
    Exit Sub

Fail:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    On Error Resume Next
    If Not wdDoc Is Nothing Then wdDoc.Close False
    wdApp.Quit
End Sub

Private Sub ReplaceTag(doc As Object, tag As String, newText As String)
    With doc.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = tag
        .Replacement.Text = newText
        .Forward = True
        .Wrap = 1            ' wdFindContinue
        .MatchWildcards = False
        .Execute Replace:=2  ' wdReplaceAll
    End With
End Sub

Private Function CleanName(s As String) As String
    Dim bad As Variant, ch As Variant
    CleanName = Trim(s)
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each ch In bad
        CleanName = Replace(CleanName, ch, "")
    Next ch
    CleanName = Replace(CleanName, " ", "_")
End Function
