Option Explicit
Sub ReplaceText(textToFind As String, replacementText As Collection)
Dim rng As Word.Range
Dim textLine As Variant

Set rng = ActiveDocument.Range
  With rng.Find
    .ClearFormatting
    .Text = textToFind
    If .Execute Then
        rng.Select
    End If
  End With
  'Now move 1 character to the right to unset the selection...
  Selection.MoveRight 1
 
  For Each textLine In replacementText
    Selection.TypeText Chr(13) + textLine
  Next textLine
End Sub
Sub InsertContent()
    Dim textToFind As String
    Dim replacementArray As Collection
    Dim filePath As String
    
    textToFind = "[content-start]"
    filePath = "c:/temp/content.txt"
    Set replacementArray = ReadFileIntoCollection(filePath)
    ReplaceText textToFind, replacementArray
End Sub
Function ReadFileIntoCollection(filePath As String) As Collection
    Open filePath For Input As #1
    Dim recordSet As Collection
    Set recordSet = New Collection
    Dim temp As String
    
    Do While Not EOF(1)
        Line Input #1, temp
        recordSet.Add temp
    Loop
    Close #1
    Set ReadFileIntoCollection = recordSet
End Function
Sub TestFileRead()
    Dim filePath As String
    Dim x As Collection

    filePath = "c:/temp/content.txt"
    Set x = ReadFileIntoCollection(filePath)
End Sub
