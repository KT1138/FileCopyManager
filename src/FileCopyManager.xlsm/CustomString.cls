VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CustomString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' 指定の文字で区切られている文字列に対し、
' 区切り文字を除いた文字列を配列に格納し、
' その要素数を返す
Public Function GetElementCountOfThisString(ByVal aString As String, ByVal aDelimiter As String) As Long
  GetElementCountOfThisString = UBound(Split(aString, aDelimiter))
End Function

' 指定の文字で区切られている文字列に対し、
' 区切り文字を除いた文字列を配列に格納。
' 指定の要素数までの文字列を区切り文字で区切って返す
' 指定の要素数が多すぎる、または少なすぎる場合、空文字を返す
Public Function GetSplitedStrings(ByVal aString As String, ByVal aDelimiter As String, ByVal aCount As Long) As String
  Dim lCollection As Collection: Set lCollection = New Collection
  Dim v As Variant
  For Each v In Split(aString, aDelimiter)
    Call lCollection.Add(v)
  Next v
  If aCount <= 0 Or aCount > lCollection.Count - 1 Then
    GetSplitedStrings = ""
    Exit Function
  End If
  Dim i As Long
  For i = 1 To aCount
    GetSplitedStrings = GetSplitedStrings & lCollection(i) & "\"
  Next i
End Function

' 指定の文字で区切られている文字列に対し、
' 区切り文字を除いた文字列を配列に格納。
' 指定の要素数に対応する文字列を返す
' 指定の要素数が多すぎる、または少なすぎる場合、空文字を返す
Public Function GetSplitedString(ByVal aString As String, ByVal aDelimiter As String, ByVal aIndex As Long) As String

End Function
