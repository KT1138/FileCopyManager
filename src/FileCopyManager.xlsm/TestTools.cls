VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' 2�̃R���N�V�������̗v�f����v���Ă��邱�Ƃ��m�F����
Public Function AreEqual(ByVal aCollection1 As Collection, ByVal aCollection2 As Collection) As Boolean
  ' �T�C�Y�̈�v���m�F
  If aCollection1.Count <> aCollection2.Count Then
    AreEqual = False
    Exit Function
  End If
  ' �v�f�̈�v���m�F
  Dim i As Long
  For i = 1 To aCollection1.Count
    If aCollection1(i) <> aCollection2(i) Then
      AreEqual = False
      Exit Function
    End If
  Next i
  
  AreEqual = True
End Function