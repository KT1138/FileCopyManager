Attribute VB_Name = "Test_FileCopyManager"
Option Explicit

Public Sub Test_All_Test_FileCopyManager()
  Test_GetAllBookNames
  Test_GetAllFolderNames
  Test_GetAllFilePathsContainThisWord
  Test_GetAllFolderPaths
  Test_GetAllFilePathsContainThisWord2
  Test_GetAllFolderPaths2
  Test_FindFiles
  Test_Main
  Test_GetParentFolderName
  Test_GetFileNameFromFilePath
  Test_GetConvertedFilePath
End Sub

Private Sub Test_GetAllBookNames()
  Dim lFileCopyManager As FileCopyManager: Set lFileCopyManager = New FileCopyManager
  Dim lTestTools As TestTools: Set lTestTools = New TestTools
  Dim lBookNames As Collection: Set lBookNames = New Collection
  Call lBookNames.Add("Hoge.xlsx")
  Call lBookNames.Add("Huga.xlsx")
  With lFileCopyManager
  Debug.Assert lTestTools.AreEqual(.Test_GetAllBookNames(.Test_GetFolderPathOfThisWorkbook & "Test\"), lBookNames)
  End With
End Sub

Private Sub Test_GetAllFolderNames()
  Dim lFileCopyManager As FileCopyManager: Set lFileCopyManager = New FileCopyManager
  Dim lTestTools As TestTools: Set lTestTools = New TestTools
  Dim lFolderNames As Collection: Set lFolderNames = New Collection
  Call lFolderNames.Add("Hoge")
  Call lFolderNames.Add("Huga")
  With lFileCopyManager
  Debug.Assert lTestTools.AreEqual(.Test_GetAllFolderNames(.Test_GetFolderPathOfThisWorkbook & "Test\"), lFolderNames)
  End With
End Sub

Private Sub Test_GetAllFileNamesContainThisWord()
  Dim lFileCopyManager As FileCopyManager: Set lFileCopyManager = New FileCopyManager
  Dim lTestTools As TestTools: Set lTestTools = New TestTools
  Dim lFileNames As Collection: Set lFileNames = New Collection
  Call lFileNames.Add("Hoge.xlsx")
  Call lFileNames.Add("HogeHuga.xlsx")
  With lFileCopyManager
  Debug.Assert lTestTools.AreEqual(.Test_GetAllFileNamesContainThisWord(.Test_GetFolderPathOfThisWorkbook & "Test\Hoge\", "Hoge"), lFileNames)
  End With
End Sub

Private Sub Test_GetAllFilePathsContainThisWord()
  Dim lFileCopyManager As FileCopyManager: Set lFileCopyManager = New FileCopyManager
  Dim lTestTools As TestTools: Set lTestTools = New TestTools
  Dim lFilePaths As Collection: Set lFilePaths = New Collection
  Call lFilePaths.Add(lFileCopyManager.Test_GetFolderPathOfThisWorkbook & "Test\Hoge.xlsx")
  With lFileCopyManager
  Debug.Assert lTestTools.AreEqual(.Test_GetAllFilePathsContainThisWord(lFileCopyManager.Test_GetFolderPathOfThisWorkbook & "Test\", "Hoge"), lFilePaths)
  End With
  
  Set lFilePaths = New Collection
  Call lFilePaths.Add(lFileCopyManager.Test_GetFolderPathOfThisWorkbook & "Test\Hoge\Hoge.xlsx")
  Call lFilePaths.Add(lFileCopyManager.Test_GetFolderPathOfThisWorkbook & "Test\Hoge\HogeHuga.xlsx")
  With lFileCopyManager
    Debug.Assert lTestTools.AreEqual(.Test_GetAllFilePathsContainThisWord(lFileCopyManager.Test_GetFolderPathOfThisWorkbook & "Test\Hoge\", "Hoge"), lFilePaths)
  End With
End Sub

Private Sub Test_GetAllFolderPaths()
  Dim lFileCopyManager As FileCopyManager: Set lFileCopyManager = New FileCopyManager
  Dim lTestTools As TestTools: Set lTestTools = New TestTools
  Dim lFolderPaths As Collection: Set lFolderPaths = New Collection
  Call lFolderPaths.Add(lFileCopyManager.Test_GetFolderPathOfThisWorkbook & "Test\Hoge\")
  Call lFolderPaths.Add(lFileCopyManager.Test_GetFolderPathOfThisWorkbook & "Test\Huga\")
  With lFileCopyManager
    Debug.Assert lTestTools.AreEqual(.Test_GetAllFolderPaths(lFileCopyManager.Test_GetFolderPathOfThisWorkbook & "Test\"), lFolderPaths)
  End With
End Sub

Private Sub Test_GetAllFilePathsContainThisWord2()
  Dim lFileCopyManager As FileCopyManager: Set lFileCopyManager = New FileCopyManager
  Dim lTestTools As TestTools: Set lTestTools = New TestTools
  Dim lFilePaths As Collection: Set lFilePaths = New Collection
  Call lFilePaths.Add(lFileCopyManager.Test_GetFolderPathOfThisWorkbook & "Test2\data0.txt")
  Call lFilePaths.Add(lFileCopyManager.Test_GetFolderPathOfThisWorkbook & "Test2\1\data1-1.txt")
  Dim lFolderPaths As Collection: Set lFolderPaths = New Collection
  Call lFolderPaths.Add(lFileCopyManager.Test_GetFolderPathOfThisWorkbook & "Test2\")
  Call lFolderPaths.Add(lFileCopyManager.Test_GetFolderPathOfThisWorkbook & "Test2\1\")
  With lFileCopyManager
    Debug.Assert lTestTools.AreEqual(.Test_GetAllFilePathsContainThisWord2(lFolderPaths, "data"), lFilePaths)
  End With
End Sub

Private Sub Test_GetAllFolderPaths2()
  Dim lFileCopyManager As FileCopyManager: Set lFileCopyManager = New FileCopyManager
  Dim lTestTools As TestTools: Set lTestTools = New TestTools
  Dim lFolderPaths As Collection: Set lFolderPaths = New Collection
  Call lFolderPaths.Add(lFileCopyManager.Test_GetFolderPathOfThisWorkbook & "Test2\1\")
  Call lFolderPaths.Add(lFileCopyManager.Test_GetFolderPathOfThisWorkbook & "Test2\2\")
  Call lFolderPaths.Add(lFileCopyManager.Test_GetFolderPathOfThisWorkbook & "Test2\3\")
  Call lFolderPaths.Add(lFileCopyManager.Test_GetFolderPathOfThisWorkbook & "Test2\4\")
  Call lFolderPaths.Add(lFileCopyManager.Test_GetFolderPathOfThisWorkbook & "Test2\2\2-1\")
  Call lFolderPaths.Add(lFileCopyManager.Test_GetFolderPathOfThisWorkbook & "Test2\2\dummy\")
  Dim aFolderPaths As Collection: Set aFolderPaths = New Collection
  Call aFolderPaths.Add(lFileCopyManager.Test_GetFolderPathOfThisWorkbook & "Test2\")
  Call aFolderPaths.Add(lFileCopyManager.Test_GetFolderPathOfThisWorkbook & "Test2\1\")
  Call aFolderPaths.Add(lFileCopyManager.Test_GetFolderPathOfThisWorkbook & "Test2\2\")
  With lFileCopyManager
    Debug.Assert lTestTools.AreEqual(.Test_GetAllFolderPaths2(aFolderPaths), lFolderPaths)
  End With
End Sub

Private Sub Test_FindFiles()
  Dim lFileCopyManager As FileCopyManager: Set lFileCopyManager = New FileCopyManager
  Dim lTestTools As TestTools: Set lTestTools = New TestTools
  Dim lFilePaths As Collection: Set lFilePaths = New Collection

  ' ê[ìx0
  Call lFilePaths.Add(lFileCopyManager.Test_GetFolderPathOfThisWorkbook & "Test2\data0.txt")
  With lFileCopyManager
    Debug.Assert lTestTools.AreEqual(.Test_FindFiles(lFileCopyManager.Test_GetFolderPathOfThisWorkbook & "Test2\", "data", 0), lFilePaths)
  End With
  ' ê[ìx1
  Call lFilePaths.Add(lFileCopyManager.Test_GetFolderPathOfThisWorkbook & "Test2\1\data1-1.txt")
  Call lFilePaths.Add(lFileCopyManager.Test_GetFolderPathOfThisWorkbook & "Test2\2\data2-1.txt")
  Call lFilePaths.Add(lFileCopyManager.Test_GetFolderPathOfThisWorkbook & "Test2\3\data3-1.txt")
  Call lFilePaths.Add(lFileCopyManager.Test_GetFolderPathOfThisWorkbook & "Test2\4\data4-1.txt")
  With lFileCopyManager
    Debug.Assert lTestTools.AreEqual(.Test_FindFiles(lFileCopyManager.Test_GetFolderPathOfThisWorkbook & "Test2\", "data", 1), lFilePaths)
  End With
  ' ê[ìx2
  Call lFilePaths.Add(lFileCopyManager.Test_GetFolderPathOfThisWorkbook & "Test2\2\2-1\data2-1.txt")
  Call lFilePaths.Add(lFileCopyManager.Test_GetFolderPathOfThisWorkbook & "Test2\2\2-1\data2-2.txt")
  Call lFilePaths.Add(lFileCopyManager.Test_GetFolderPathOfThisWorkbook & "Test2\3\3-1\data3-1.txt")
  Call lFilePaths.Add(lFileCopyManager.Test_GetFolderPathOfThisWorkbook & "Test2\3\3-1\data3-2.txt")
  Call lFilePaths.Add(lFileCopyManager.Test_GetFolderPathOfThisWorkbook & "Test2\4\4-1\data4-1.txt")
  Call lFilePaths.Add(lFileCopyManager.Test_GetFolderPathOfThisWorkbook & "Test2\4\4-1\data4-2.txt")
  With lFileCopyManager
    Debug.Assert lTestTools.AreEqual(.Test_FindFiles(lFileCopyManager.Test_GetFolderPathOfThisWorkbook & "Test2\", "data", 2), lFilePaths)
  End With
  ' ê[ìx3
    Call lFilePaths.Add(lFileCopyManager.Test_GetFolderPathOfThisWorkbook & "Test2\3\3-1\3-2\data3-1.txt")
    Call lFilePaths.Add(lFileCopyManager.Test_GetFolderPathOfThisWorkbook & "Test2\3\3-1\3-2\data3-2.txt")
    Call lFilePaths.Add(lFileCopyManager.Test_GetFolderPathOfThisWorkbook & "Test2\3\3-1\3-2\data3-3.txt")
    Call lFilePaths.Add(lFileCopyManager.Test_GetFolderPathOfThisWorkbook & "Test2\4\4-1\4-2\data4-1.txt")
    Call lFilePaths.Add(lFileCopyManager.Test_GetFolderPathOfThisWorkbook & "Test2\4\4-1\4-2\data4-2.txt")
    Call lFilePaths.Add(lFileCopyManager.Test_GetFolderPathOfThisWorkbook & "Test2\4\4-1\4-2\data4-3.txt")
    With lFileCopyManager
      Debug.Assert lTestTools.AreEqual(.Test_FindFiles(lFileCopyManager.Test_GetFolderPathOfThisWorkbook & "Test2\", "data", 3), lFilePaths)
    End With
  ' ê[ìx4
    Call lFilePaths.Add(lFileCopyManager.Test_GetFolderPathOfThisWorkbook & "Test2\4\4-1\4-2\4-3\data4-1.txt")
    Call lFilePaths.Add(lFileCopyManager.Test_GetFolderPathOfThisWorkbook & "Test2\4\4-1\4-2\4-3\data4-2.txt")
    Call lFilePaths.Add(lFileCopyManager.Test_GetFolderPathOfThisWorkbook & "Test2\4\4-1\4-2\4-3\data4-3.txt")
    Call lFilePaths.Add(lFileCopyManager.Test_GetFolderPathOfThisWorkbook & "Test2\4\4-1\4-2\4-3\data4-4.txt")
    With lFileCopyManager
      Debug.Assert lTestTools.AreEqual(.Test_FindFiles(lFileCopyManager.Test_GetFolderPathOfThisWorkbook & "Test2\", "data", 4), lFilePaths)
    End With
End Sub

Private Sub Test_Main()
  Dim lFileCopyManager As FileCopyManager: Set lFileCopyManager = New FileCopyManager
  Dim lTestTools As TestTools: Set lTestTools = New TestTools
  Dim lFSO As Object: Set lFSO = CreateObject("Scripting.FileSystemObject")
  Call lFileCopyManager.Test_RemoveAllFiles(lFileCopyManager.Test_GetFolderPathOfThisWorkbook & "Test3\")
  Call lFileCopyManager.Main(lFileCopyManager.Test_GetFolderPathOfThisWorkbook & "Test2\", "data", 4, lFileCopyManager.Test_GetFolderPathOfThisWorkbook & "Test3\")
  
  Dim lFileNames As Collection: Set lFileNames = New Collection
  Call lFileNames.Add("data0.txt")
  Call lFileNames.Add("data1-1.txt")
  Call lFileNames.Add("data2-1.txt")
  Call lFileNames.Add("data2-2.txt")
  Call lFileNames.Add("data3-1.txt")
  Call lFileNames.Add("data3-2.txt")
  Call lFileNames.Add("data3-3.txt")
  Call lFileNames.Add("data4-1.txt")
  Call lFileNames.Add("data4-2.txt")
  Call lFileNames.Add("data4-3.txt")
  Call lFileNames.Add("data4-4.txt")
    
  With lFileCopyManager
  Debug.Assert lTestTools.AreEqual(.Test_GetAllBookNames(.Test_GetFolderPathOfThisWorkbook & "Test3\"), lFileNames)
  End With
End Sub

Private Sub Test_GetParentFolderName()
  Dim lFileCopyManager As FileCopyManager: Set lFileCopyManager = New FileCopyManager
  Debug.Assert lFileCopyManager.Test_GetParentFolderName("C:\Users\Hoge\Desktop\FileCopyManager\bin", 1) = "FileCopyManager"
  Debug.Assert lFileCopyManager.Test_GetParentFolderName("C:\Users\Hoge\Desktop\FileCopyManager\bin", 2) = "Desktop"
  Debug.Assert lFileCopyManager.Test_GetParentFolderName("C:\Users\Hoge\Desktop\FileCopyManager\bin", 3) = "Hoge"
  Debug.Assert lFileCopyManager.Test_GetParentFolderName("C:\Users\Hoge\Desktop\FileCopyManager\bin", 4) = "Users"
  Debug.Assert lFileCopyManager.Test_GetParentFolderName("C:\Users\Hoge\Desktop\FileCopyManager\bin", 5) = "C:"
  Debug.Assert lFileCopyManager.Test_GetParentFolderName("C:\Users\Hoge\Desktop\FileCopyManager\bin", 6) = ""
End Sub

Private Sub Test_GetFileNameFromFilePath()
  Dim lFileCopyManager As FileCopyManager: Set lFileCopyManager = New FileCopyManager
  Debug.Assert lFileCopyManager.Test_GetFileNameFromFilePath("C:\Users\Hoge\Desktop\FileCopyManager\bin\Hoge.xlsx") = "Hoge.xlsx"
  Debug.Assert lFileCopyManager.Test_GetFileNameFromFilePath("C:\Users\Hoge\Desktop\FileCopyManager\Hoge.xlsx") = "Hoge.xlsx"
  Debug.Assert lFileCopyManager.Test_GetFileNameFromFilePath("C:\Users\Hoge\Desktop\Hoge.xlsx") = "Hoge.xlsx"
  Debug.Assert lFileCopyManager.Test_GetFileNameFromFilePath("C:\Users\Hoge\Desktop\") = ""
End Sub

Private Sub Test_GetConvertedFilePath()
  Dim lFileCopyManager As FileCopyManager: Set lFileCopyManager = New FileCopyManager
  Debug.Assert lFileCopyManager.Test_GetConvertedFilePath("C:\Users\Hoge\Desktop\FileCopyManager\bin\Hoge.xlsx", 2, 1, "C:\Users\Hoge\Desktop\Destination\") = _
    "C:\Users\Hoge\Desktop\Destination\FileCopyManager_bin_Hoge.xlsx"
  Debug.Assert lFileCopyManager.Test_GetConvertedFilePath("C:\Users\Hoge\Desktop\FileCopyManager\bin\Hoge.xlsx", 3, 1, "C:\Users\Hoge\Desktop\Destination\") = _
    "C:\Users\Hoge\Desktop\Destination\Desktop_bin_Hoge.xlsx"
  Debug.Assert lFileCopyManager.Test_GetConvertedFilePath("C:\Users\Hoge\Desktop\FileCopyManager\bin\Hoge.xlsx", 7, 1, "C:\Users\Hoge\Desktop\Destination\") = _
    "C:\Users\Hoge\Desktop\Destination\_bin_Hoge.xlsx"
  Debug.Assert lFileCopyManager.Test_GetConvertedFilePath("C:\Users\Hoge\Desktop\FileCopyManager\bin\Hoge.xlsx", 8, 9, "C:\Users\Hoge\Desktop\Destination\") = _
    "C:\Users\Hoge\Desktop\Destination\__Hoge.xlsx"
End Sub
