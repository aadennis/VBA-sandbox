'Assumptions
'Any bookmarks, style names, referenced in here exist in the host document
Option Explicit
' Start Edit
Const WORKING_FOLDER As String = "C:\GitHub\VBA-sandbox\ProjectVEST\documentation"
Const DELIVERY_CONTENT_FILE = "content.txt"
Const PATCH_NARRATIVE_FILE = "patchNarrative.txt"
' Stop Edit
Const CONTENT_START_TAG As String = "CONTENT-START"
Const CONTENT_END_TAG As String = "CONTENT-END"
Const NARRATIVE_START_TAG As String = "NARRATIVE-START"

Enum fileType
    DeliveryContent = 1
    PatchNarrative = 2
End Enum

Sub ReplaceText(textToFind As String, replacementText As Collection)
'Having found TextToFind, position the curstor immediately to the right of that string,
'and insert the whole of replacementText, with a carriage return after each record
' in the collection
Dim textLine As Variant
Dim rng As Word.Range

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
Sub InsertText(replacementText As Collection)
' Insert the whole of replacementText (a collection containing 1 or more records) at the current cursor position.
' That position was likely navigated to using some other, independent method.
Dim textLine As Variant
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
Sub InsertNarrative()
    Dim replacementArray As Collection
    Dim tempElement As Variant
        
    Set replacementArray = ReadFileIntoCollection(PatchNarrative)
    ReplaceText NARRATIVE_START_TAG, replacementArray
End Sub
Function ReadFileIntoCollection2(filePath As String) As Collection
'Read a file from disk into a collection
'A key principle is that no validation of the file content happens during the read.
'As the data sizes will always be trivial, the file is wholly read into a Collection
'Any validation or trimming of the content happens in memory - files (until we decide later)
'are always read-only

    Dim recordSet As Collection: Set recordSet = New Collection
    Dim recordFound As Boolean: recordFound = False
    Dim temp As String
    
    If Dir(filePath) = "" Then
   
      MsgBox "file [" & filePath & "] does not exist. Exiting."
      Exit Function
    
    End If
    
    Open filePath For Input As #1
    Do While Not EOF(1)
        recordFound = True
        Line Input #1, temp
        recordSet.Add temp
    Loop
    Close #1
    
    If Not recordFound Then
        MsgBox "[" & filePath & "] contained no records"
        Close #1
        Exit Function
    End If
    Set ReadFileIntoCollection2 = recordSet
     
End Function

Function ReadFileIntoCollection(file As fileType) As Collection
'Read a file from disk into a collection
    Dim filePath As String
    Dim IsPatchNarrative As Boolean
    IsPatchNarrative = False
    
    Select Case file
        Case fileType.DeliveryContent
            filePath = WORKING_FOLDER & "/" & DELIVERY_CONTENT_FILE
        Case fileType.PatchNarrative
            filePath = WORKING_FOLDER & "/" & PATCH_NARRATIVE_FILE
            IsPatchNarrative = True
        Case Else
            MsgBox "Invalid parameter in [ReadFileIntoCollection]"
    End Select
            
    Open filePath For Input As #1
    
    Dim recordSet As Collection
    Set recordSet = New Collection
    Dim temp As String
    
    Dim found As Boolean
    found = False
    
    Do While Not EOF(1)
        found = True
        Line Input #1, temp
        'special case for line 1 of patch narrative (hack)...
        If (IsPatchNarrative) Then
          IsPatchNarrative = False
          temp = ConvertFacetedCodeToNarrative(temp)
        End If
        recordSet.Add temp
    Loop
    Close #1
    If Not found Then
        MsgBox "did not find [" & filePath & "]"
        Exit Function
    End If
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
Function CheckIntegrity()
    ' Check 1 - the 3 tags denoting the replacement positions must be found...
    Dim rng As Range
    Dim tagSet(2) As String
    tagSet(0) = CONTENT_START_TAG
    tagSet(1) = CONTENT_END_TAG
    tagSet(2) = NARRATIVE_START_TAG
    Dim tag As Variant
    
    For Each tag In tagSet
        Set rng = ActiveDocument.Range
        If Not rng.Find.Execute(FindText:=tag) Then
            Err.Raise vbObjectError + 513, "fatal flaw", "Did not Find " & tag & " in file. Exiting run..."
        End If
    Next tag
    MsgBox "Integrity check OK."
    
End Function
Function ConvertFacetedCodeToNarrative(code As String)
  ConvertFacetedCodeToNarrative = "EOS Release " & Mid(code, 1, 3) & " Patch " & Mid(code, 5)
End Function
Sub DoStyles()
Dim objDoc As Document
Dim head1 As Style, head2 As Style, heading3 As Style, head4 As Style
Dim NarrativeHeader As String
NarrativeHeader = "Patch Release 3.2 Patch 6"
Set objDoc = ActiveDocument


With objDoc.Content.Find
    .ClearFormatting
    .Text = NarrativeHeader
    With .Replacement
    .ClearFormatting
    .Style = heading3
    End With
    ' Here we do the actual replacement. Based on your requirements, this only replaces the
    ' first instance it finds. You could also change this to Replace:=wdReplaceAll to catch
    ' all of them.
    .Execute Wrap:=wdFindContinue, Format:=True, Replace:=wdReplaceOne
End With

End Sub
Sub SelectCurrentParagraph()
  Selection.Paragraphs(1).Range.Select
  
End Sub
Sub ApplyStyleToCurrentParagraph(styleName)
  SelectCurrentParagraph
  Selection.Style = ActiveDocument.Styles(styleName)
End Sub

Sub TestGetWordDelimitedRange()
    Dim a As String
    Dim b As String
    Dim retVal As String
    
    a = "metre"
    b = "Corrective"
    retVal = GetWordDelimitedRange(a, b)
    MsgBox retVal
End Sub
Sub TestCheckIntegrity()
    On Error GoTo errMyErrorHandler
    CheckIntegrity
errMyErrorHandler:
     MsgBox Err.Description
End Sub
Sub TestForIntegration()
    On Error GoTo errMyErrorHandler
    Dim retVal As String
    CheckIntegrity
    retVal = TruncateContent(CONTENT_START_TAG, CONTENT_END_TAG)
    InsertContent
    InsertNarrative
    
    Exit Sub
errMyErrorHandler:
    MsgBox Err.Description
End Sub
Sub TestConvertFacetedCodeToNarrative()
  MsgBox (ConvertFacetedCodeToNarrative("5.2.3999"))
End Sub
Sub TestInsertNarrative()
  InsertNarrative
End Sub
Sub TestDoStyles()
  DoStyles

End Sub
Sub bookmarkandDown(bookmarkName)
  Selection.GoTo What:=wdGoToBookmark, Name:=bookmarkName
  Selection.MoveDown Unit:=wdLine, Count:=1
End Sub

Sub TestGotoBookmark()
' Note that any use of bookmarks assumes they exist already in the host Word document
  bookmarkandDown ("NARRATIVE_START")
  Dim replacementArray As Collection
  
  'Selection.TypeText ConvertFacetedCodeToNarrative("3.2.3999") + Chr(13)
  Set replacementArray = ReadFileIntoCollection(PatchNarrative)
 InsertText replacementArray
End Sub
Sub TestApplyStyle()

ApplyStyleToCurrentParagraph ("EOSBullet2")
End Sub
Sub TestGetKeyValueCollectionFromRawCollection()
  Const delim As String = ":" 'in the real world this will be passed in
  
  Dim mockFile As New Collection
  Dim narrative As New Collection
  Dim record As Variant
  
  mockFile.Add ("EOS ID: This is the entry")
  mockFile.Add ("Problem: This is the problem")
  mockFile.Add ("")
  mockFile.Add ("Solution: This is the solution")
  mockFile.Add ("EOS ID: This is entry 2")
  mockFile.Add ("Problem: This is problem 2")
  mockFile.Add ("Solution: This is solution 2")
    
  For Each record In mockFile
    'Validate...
    '1. lines can be empty...
    If (Len(record) = 0) Then
      GoTo Continue
    End If
    'Validation ok...
    narrative.Add (Split(record, delim, 2, vbTextCompare))
    
Continue:
  Next
  
Debug.Assert mockFile.Count = 7
Debug.Assert narrative.Count = 6
'note that the row is 1-based... but the column is zero-based...
Debug.Assert narrative(1)(0) = "EOS ID"
Debug.Assert narrative(1)(1) = " This is the entry"

  Set record = Nothing
  Set mockFile = Nothing
  Set narrative = Nothing
End Sub
Sub TestReadFileIntoCollection2()
  'this test depends on a file at the location filePath,
  'with this content (3 records):
  'in the Town where I was born 22
  'lived a MaN who sailed !! to sea
  'And he told US OF HIS life
  
  Dim recordSet As Collection
  Dim filePath As String: filePath = "c:/temp/test.txt"
 
  Set recordSet = ReadFileIntoCollection2(filePath)
  
  Debug.Assert recordSet.Count = 3
  Debug.Assert recordSet(1) = "in the Town where I was born 22"
  
  Set recordSet = Nothing

End Sub
Sub AddRecordToNarrative()
  

End Sub
