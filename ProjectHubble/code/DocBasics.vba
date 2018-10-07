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
Function TruncateContent(startTag As String, endTag As String)
    'without deleting the tags, delete everything between the start and end tag
    Dim rng1 As Range
    Dim rng2 As Range
    Dim rng3 As Range
    
    Dim strTheText As String

    Set rng1 = ActiveDocument.Range
    If rng1.Find.Execute(FindText:=startTag) Then
        Set rng2 = ActiveDocument.Range(rng1.End, ActiveDocument.Range.End)
        If rng2.Find.Execute(FindText:=endTag) Then
            Set rng3 = ActiveDocument.Range(rng1.End, rng2.Start)
            rng3.Select
            'hack to delete selected text and add a general purpose space
            Selection.TypeText " "
        End If
    End If
End Function

Sub TestFileRead()
    Dim filePath As String
    Dim x As Collection

    filePath = "c:/temp/content.txt"
    Set x = ReadFileIntoCollection(filePath)
End Sub
Function GetWordDelimitedRange(startTag As String, endTag As String)
    'The found string/range excludes the values in the Tags
    Dim rng1 As Range
    Dim rng2 As Range
    Dim strTheText As String

    Set rng1 = ActiveDocument.Range
    If rng1.Find.Execute(FindText:=startTag) Then
        Set rng2 = ActiveDocument.Range(rng1.End, ActiveDocument.Range.End)
        If rng2.Find.Execute(FindText:=endTag) Then
            GetWordDelimitedRange = CStr(ActiveDocument.Range(rng1.End, rng2.Start).Text)
        End If
    End If
End Function
Sub TestGetWordDelimitedRange()
    Dim a As String
    Dim b As String
    Dim retVal As String
    
    a = "metre"
    b = "Corrective"
    retVal = GetWordDelimitedRange(a, b)
    MsgBox retVal
End Sub
Sub TestTruncateContent()
    Dim a As String
    Dim b As String
    Dim retVal As String
    
    a = "[CONTENT-START]"
    b = "[CONTENT-END]"
    retVal = TruncateContent(a, b)
    

End Sub
Sub TestForIntegration()
TestTruncateContent
InsertContent

End Sub
