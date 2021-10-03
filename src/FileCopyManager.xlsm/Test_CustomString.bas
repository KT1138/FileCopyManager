Attribute VB_Name = "Test_CustomString"
Option Explicit

Public Sub Test_All_Test_CustomString()
  Test_GetElementCountOfThisString
  Test_GetSplitedStrings
  Test_GetSplitedString
End Sub

Private Sub Test_GetElementCountOfThisString()
  Dim CustomString As CustomString: Set CustomString = New CustomString
  Dim lTestTools As TestTools: Set lTestTools = New TestTools
  With CustomString
    Debug.Assert CustomString.GetElementCountOfThisString("C:\Users\Hoge\Desktop\", "\") = 4
  End With
End Sub

Private Sub Test_GetSplitedStrings()
  Dim CustomString As CustomString: Set CustomString = New CustomString
  Dim lTestTools As TestTools: Set lTestTools = New TestTools
  With CustomString
    Debug.Assert CustomString.GetSplitedStrings("C:\Users\Hoge\Desktop\", "\", 5) = ""
    Debug.Assert CustomString.GetSplitedStrings("C:\Users\Hoge\Desktop\", "\", 4) = "C:\Users\Hoge\Desktop\"
    Debug.Assert CustomString.GetSplitedStrings("C:\Users\Hoge\Desktop\", "\", 3) = "C:\Users\Hoge\"
    Debug.Assert CustomString.GetSplitedStrings("C:\Users\Hoge\Desktop\", "\", 2) = "C:\Users\"
    Debug.Assert CustomString.GetSplitedStrings("C:\Users\Hoge\Desktop\", "\", 1) = "C:\"
    Debug.Assert CustomString.GetSplitedStrings("C:\Users\Hoge\Desktop\", "\", 0) = ""
  End With
End Sub

Private Sub Test_GetSplitedString()
  Dim CustomString As CustomString: Set CustomString = New CustomString
  Dim lTestTools As TestTools: Set lTestTools = New TestTools
  With CustomString
    Debug.Assert CustomString.GetSplitedString("C:\Users\Hoge\Desktop\", "\", 5) = ""
    Debug.Assert CustomString.GetSplitedString("C:\Users\Hoge\Desktop\", "\", 4) = "Desktop"
    Debug.Assert CustomString.GetSplitedString("C:\Users\Hoge\Desktop\", "\", 3) = "Hoge"
    Debug.Assert CustomString.GetSplitedString("C:\Users\Hoge\Desktop\", "\", 2) = "Users"
    Debug.Assert CustomString.GetSplitedString("C:\Users\Hoge\Desktop\", "\", 1) = "C:"
    Debug.Assert CustomString.GetSplitedString("C:\Users\Hoge\Desktop\", "\", 0) = ""
    Debug.Assert CustomString.GetSplitedString("C:\Users\Hoge\Desktop\", "\", 5) = ""
  End With
End Sub

