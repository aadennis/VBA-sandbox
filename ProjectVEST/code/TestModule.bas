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

