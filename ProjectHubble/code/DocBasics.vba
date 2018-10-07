Option Explicit
' Start Edit
Const WORKING_FOLDER As String = "C:\GitHub\VBA-sandbox\ProjectHubble\documentation"
Const DELIVERY_CONTENT_FILE = "content.txt"
Const PATCH_NARRATIVE_FILE = "narrative.txt"
' Stop Edit
Const CONTENT_START_TAG As String = "[CONTENT-START]"
Const CONTENT_END_TAG As String = "[CONTENT-END]"
Const NARRATIVE_START_TAG As String = "[NARRATIVE-START]"

Enum fileType
    DeliveryContent = 1
    PatchNarrative = 2
End Enum

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
    
    Set replacementArray = ReadFileIntoCollection(DeliveryContent)
    ReplaceText CONTENT_START_TAG, replacementArray
End Sub
Function ReadFileIntoCollection(file As fileType) As Collection
    Dim filePath As String
    
    Select Case file
        Case fileType.DeliveryContent
            filePath = WORKING_FOLDER & "/" & DELIVERY_CONTENT_FILE
        Case fileType.PatchNarrative
            filePath = WORKING_FOLDER & "/" & PATCH_NARRATIVE_FILE
        Case Else
            MsgBox "Invalid parameter in [ReadFileIntoCollection]"
    End Select
            
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
            'hack to delete selected text
            Selection.TypeText " "
        End If
    End If
End Function
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
Sub TestFileRead()
    Dim filePath As String
    Dim x As Collection

    filePath = "c:/temp/content.txt"
    Set x = ReadFileIntoCollection(filePath)
End Sub
Sub TestForIntegration()
    Dim retVal As String
    retVal = TruncateContent(CONTENT_START_TAG, CONTENT_END_TAG)
    InsertContent
End Sub
