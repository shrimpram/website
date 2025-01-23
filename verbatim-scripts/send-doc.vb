Sub SendDoc()
    Dim originalDoc as Document
    Set originalDoc = ActiveDocument

    ' Disable screen updating for faster execution
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Extract the folder path from the original document's file path
    Dim originalFolderPath As String
    Dim originalFilePath As String
    originalFolderPath = Left(originalDoc.FullName, InStrRev(originalDoc.FullName, Application.PathSeparator))
    originalFilePath = originalDoc.FullName

    ' Set the save path for the modified document in the same folder as the original document
    savePath = originalFolderPath & "[S] " & originalDoc.Name

    ' Check if doc has previously been saved
    If ActiveDocument.Path = "" Then
        ' If not previously saved
        MsgBox "The current document must be saved at least once."
        Exit Sub
    End If

    ' If previously saved, create a copy
    Dim sendDoc As Document
    Set sendDoc = Documents.Add(ActiveDocument.FullName)
    Selection.WholeStory
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Analytics")
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll


    Dim savePath As String
    savePath = originalFolderPath & "[S] " & originalDoc.Name
    ActiveDocument.SaveAs2 FileName:=savePath, FileFormat:=wdFormatDocumentDefault
End Sub
