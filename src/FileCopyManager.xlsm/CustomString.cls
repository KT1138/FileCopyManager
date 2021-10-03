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

' �w��̕����ŋ�؂��Ă��镶����ɑ΂��A
' ��؂蕶�����������������z��Ɋi�[���A
' ���̗v�f����Ԃ�
Public Function GetElementCountOfThisString(ByVal aString As String, ByVal aDelimiter As String) As Long
  GetElementCountOfThisString = UBound(Split(aString, aDelimiter))
End Function

' �w��̕����ŋ�؂��Ă��镶����ɑ΂��A
' ��؂蕶�����������������z��Ɋi�[�B
' �w��̗v�f���܂ł̕��������؂蕶���ŋ�؂��ĕԂ�
' �w��̗v�f������������A�܂��͏��Ȃ�����ꍇ�A�󕶎���Ԃ�
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

' �w��̕����ŋ�؂��Ă��镶����ɑ΂��A
' ��؂蕶�����������������z��Ɋi�[�B
' �w��̗v�f���ɑΉ����镶�����Ԃ�
' �w��̗v�f������������A�܂��͏��Ȃ�����ꍇ�A�󕶎���Ԃ�
Public Function GetSplitedString(ByVal aString As String, ByVal aDelimiter As String, ByVal aIndex As Long) As String

End Function