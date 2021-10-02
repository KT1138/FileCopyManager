Attribute VB_Name = "Test_TestTools"
Option Explicit

Private Sub Test_AreEqual()
  Dim lCollection1 As Collection: Set lCollection1 = New Collection
  Dim lCollection2 As Collection: Set lCollection2 = New Collection
  Dim lTestTools As TestTools: Set lTestTools = New TestTools
  
  ' ��̃R���N�V�������r
  Debug.Assert lTestTools.AreEqual(lCollection1, lCollection2)
  
  ' �v�f���̈قȂ�R���N�V�������r
  Call lCollection1.Add("Hoge")
  Debug.Assert Not lTestTools.AreEqual(lCollection1, lCollection2)
  
  ' �v�f�̈قȂ�R���N�V�������r
  Call lCollection2.Add("Huga")
  Debug.Assert Not lTestTools.AreEqual(lCollection1, lCollection2)
  
  ' ��v����R���N�V�������r
  Set lCollection1 = New Collection
  Set lCollection2 = New Collection
  Call lCollection1.Add("Hoge")
  Call lCollection2.Add("Hoge")
  Debug.Assert lTestTools.AreEqual(lCollection1, lCollection2)
  Call lCollection1.Add("Huga")
  Call lCollection2.Add("Huga")
  Debug.Assert lTestTools.AreEqual(lCollection1, lCollection2)
End Sub
