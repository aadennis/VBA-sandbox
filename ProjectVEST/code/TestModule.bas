Attribute VB_Name = "TestModule"
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

Sub RunAllTests()
  TestReadFileIntoCollection2
  TestGetKeyValueCollectionFromRawCollection
  
End Sub
