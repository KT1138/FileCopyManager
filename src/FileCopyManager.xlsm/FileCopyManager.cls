VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileCopyManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Private mFSO As Object
'
'Public Sub Class_Initalize()
'  Set mFSO = CreateObject("Scripting.FileSystemObject")
'End Sub

' �w��̃t�H���_�̃u�b�N�����A�R���N�V�����Ɋi�[���ĕԂ�
' �����ɂ́A�t�H���_���΃p�X�ŁA������ "\" �����ĕԂ�
Private Function GetAllBookNames(ByVal lFolderPath As String) As Collection
  Set GetAllBookNames = New Collection
  Dim lBookName As String: lBookName = Dir(lFolderPath)
  Do While lBookName <> ""
    Call GetAllBookNames.Add(lBookName)
    lBookName = Dir()
  Loop
End Function
Public Function Test_GetAllBookNames(ByVal lFolderPath As String) As Collection
  Set Test_GetAllBookNames = GetAllBookNames(lFolderPath)
End Function

' �w��̃t�H���_���̃t�H���_�����A�R���N�V�����Ɋi�[���ĕԂ�
' �����ɂ́A�t�H���_���΃p�X�ŁA������"\"�����ēn��
Private Function GetAllFolderNames(ByVal lFolderPath As String) As Collection
  Set GetAllFolderNames = New Collection
  Dim lFolderName As String: lFolderName = Dir(lFolderPath, vbDirectory)
  Do While lFolderName <> ""
    If GetAttr(lFolderPath & lFolderName) <> vbDirectory Then GoTo NextLoop
    If lFolderName = "." Then GoTo NextLoop
    If lFolderName = ".." Then GoTo NextLoop
    Call GetAllFolderNames.Add(lFolderName)
NextLoop:
    lFolderName = Dir()
  Loop
End Function
Public Function Test_GetAllFolderNames(ByVal lFolderPath As String) As Collection
  Set Test_GetAllFolderNames = GetAllFolderNames(lFolderPath)
End Function

' ���̃u�b�N�̃t�H���_�����΃p�X�œ���
Private Function GetFolderPathOfThisWorkbook() As String
  GetFolderPathOfThisWorkbook = Left(ActiveWorkbook.FullName, InStrRev(ActiveWorkbook.FullName, "\"))
End Function
Public Function Test_GetFolderPathOfThisWorkbook() As String
  Test_GetFolderPathOfThisWorkbook = GetFolderPathOfThisWorkbook
End Function

' �w��̃t�H���_���̃u�b�N�̂����A�w��̕������܂ރu�b�N���݂̂� Collection �Ɋi�[���ĕԂ�
' Excel �u�b�N�ȊO�̃t�@�C�����͕Ԃ��Ȃ�
Public Function GetAllFileNamesContainThisWord(ByVal lFolderPath As String, ByVal aWord As String) As Collection
  Set GetAllFileNamesContainThisWord = New Collection
  Dim v As Variant
  For Each v In GetAllBookNames(lFolderPath)
    If InStr(v, aWord) = 0 Then GoTo NextLoop
    'If InStr(v, "xls") > 0 Or InStr(v, "xlsx") > 0 Or InStr(v, ".xlsm") > 0 Then
    Call GetAllFileNamesContainThisWord.Add(v)
    'End If
NextLoop:
  Next v
End Function
Public Function Test_GetAllFileNamesContainThisWord(ByVal lFolderPath As String, ByVal aWord As String) As Collection
  Set Test_GetAllFileNamesContainThisWord = GetAllFileNamesContainThisWord(lFolderPath, aWord)
End Function

' �����Ώۂ̃t�H���_�Ɏw��̕������܂ރt�@�C�������������ꍇ
' �������΃p�X�ŃR���N�V�����Ɋi�[���ĕԂ�
' aFolderPath �c �����ΏۂƂ���t�H���_
' aKeyWord �c �����ΏۂƂ��镶��
Private Function GetAllFilePathsContainThisWord(ByVal aFolderPath As String, ByVal aKeyWord As String) As Collection
  Set GetAllFilePathsContainThisWord = New Collection
  Dim v As Variant
  For Each v In GetAllFileNamesContainThisWord(aFolderPath, aKeyWord)
    Call GetAllFilePathsContainThisWord.Add(aFolderPath & v)
  Next v
End Function
Public Function Test_GetAllFilePathsContainThisWord(ByVal aFolderPath As String, ByVal aKeyWord As String)
  Set Test_GetAllFilePathsContainThisWord = GetAllFilePathsContainThisWord(aFolderPath, aKeyWord)
End Function
' ��L�v���V�[�W���̃R���N�V�����Ή���
Private Function GetAllFilePathsContainThisWord2(ByVal aFolderPathCollection As Collection, ByVal aKeyWord As String) As Collection
  Set GetAllFilePathsContainThisWord2 = New Collection
  Dim v As Variant
  For Each v In aFolderPathCollection
    Dim w As Variant
    For Each w In GetAllFilePathsContainThisWord(v, aKeyWord)
      Call GetAllFilePathsContainThisWord2.Add(w)
    Next w
  Next v
End Function
Public Function Test_GetAllFilePathsContainThisWord2(ByVal aFolderPaths As Collection, ByVal aKeyWord As String)
  Set Test_GetAllFilePathsContainThisWord2 = GetAllFilePathsContainThisWord2(aFolderPaths, aKeyWord)
End Function

' �����Ώۂ̃t�H���_�ɃT�u�t�H���_���������ꍇ
' �������΃p�X�ŃR���N�V�����Ɋi�[���ĕԂ�
' aFolderPath �c �����ΏۂƂ���t�H���_
Private Function GetAllFolderPaths(ByVal aFolderPath As String) As Collection
  Set GetAllFolderPaths = New Collection
  Dim v As Variant
  For Each v In GetAllFolderNames(aFolderPath)
      Call GetAllFolderPaths.Add(aFolderPath & v & "\")
  Next v
End Function
Public Function Test_GetAllFolderPaths(ByVal aFolderPath As String)
  Set Test_GetAllFolderPaths = GetAllFolderPaths(aFolderPath)
End Function
' ��L�v���V�[�W���̃R���N�V�����Ή���
Private Function GetAllFolderPaths2(ByVal aFolderPaths As Collection) As Collection
  Set GetAllFolderPaths2 = New Collection
  Dim v As Variant
  For Each v In aFolderPaths
    Dim w As Variant
    For Each w In GetAllFolderPaths(v)
      Call GetAllFolderPaths2.Add(w)
    Next w
  Next v
End Function
Public Function Test_GetAllFolderPaths2(ByVal aFolderPaths As Collection)
  Set Test_GetAllFolderPaths2 = GetAllFolderPaths2(aFolderPaths)
End Function

' �N�_�t�H���_�̃T�u�t�H���_�ɑ΂��A�w��̕������܂ރt�@�C��������������
' �t�@�C��������΁A���̐�΃p�X���R���N�V�����Ɋi�[���ĕԂ�
' aRootFolder �c �N�_�t�H���_�B��΃p�X�œ���(������"\"��t���邱��)
' aKeyWord �c ���̕������܂ރt�@�C������������
' aDepth �c ���K�w���̃T�u�t�H���_�܂Ō������s�����w��B�N�_�t�H���_���̂݌����������ꍇ�A0 ���w�肷��
Private Function FindFiles(ByVal aRootFolder As String, ByVal aKeyWord As String, ByVal aDepth As Long) As Collection
  Set FindFiles = New Collection
  Dim lFolderPaths As Collection: Set lFolderPaths = New Collection: Call lFolderPaths.Add(aRootFolder)
  Dim i As Long
  For i = 0 To aDepth
    ' �t�H���_���̃t�@�C���ɑ΂��Č�������
    Dim lFilePath As Variant
    For Each lFilePath In GetAllFilePathsContainThisWord2(lFolderPaths, aKeyWord)
      Call FindFiles.Add(lFilePath)
    Next lFilePath
    
    ' �T�u�t�H���_�ɑ΂��Ă̏���
    Dim lTempFolderPaths As Collection: Set lTempFolderPaths = New Collection
    Dim lFolderPath As Variant
    For Each lFolderPath In GetAllFolderPaths2(lFolderPaths)
        Call lTempFolderPaths.Add(lFolderPath)
    Next lFolderPath
    
    ' ���̃��[�v�p�ɕϐ��̒l���X�V
    Set lFolderPaths = lTempFolderPaths
  Next i
End Function
Public Function Test_FindFiles(ByVal aRootFolder As String, ByVal aKeyWord As String, ByVal aDepth As Long) As Collection
  Set Test_FindFiles = FindFiles(aRootFolder, aKeyWord, aDepth)
End Function

' �t�H���_���̃t�@�C�������ׂč폜����
' aTargetFolder �c �Ώۃt�H���_�B��΃p�X�œ���(������"\"��t���邱��)
Private Sub RemoveAllFiles(ByVal aTargetFolder As String)
  Dim lFSO As Object: Set lFSO = CreateObject("Scripting.FileSystemObject")
  Dim lFileName As Variant
  For Each lFileName In GetAllBookNames(aTargetFolder)
    Call lFSO.deletefile(aTargetFolder & "\" & lFileName, True)
  Next lFileName
End Sub
Public Sub Test_RemoveAllFiles(ByVal aTargetFolder As String)
  Call RemoveAllFiles(aTargetFolder)
End Sub

' �N�_�t�H���_�̃T�u�t�H���_�ɑ΂��A�w��̕������܂ރt�@�C��������������
' �t�@�C��������΁A�R�s�[��t�H���_�փR�s�[����
' �R�s�[��ɓ����t�@�C��������ꍇ�͏㏑������
' aRootFolder �c �N�_�t�H���_�B��΃p�X�œ���(������"\"��t���邱��)
' aKeyWord �c ���̕������܂ރt�@�C������������
' aDepth �c ���K�w���̃T�u�t�H���_�܂Ō������s�����w��B�N�_�t�H���_���̂݌����������ꍇ�A0 ���w�肷��
' aDestinationFolder �c �R�s�[��t�H���_�B��΃p�X�œ���(������"\"��t���邱��)
Public Sub Main(ByVal aRootFolder As String, ByVal aKeyWord As String, ByVal aDepth As Long, ByVal aDestinationFolder As String)
  Dim lFSO As Object: Set lFSO = CreateObject("Scripting.FileSystemObject")
  Dim lTargetFilePath As Variant
  For Each lTargetFilePath In FindFiles(aRootFolder, aKeyWord, aDepth)
    Dim lFile As Object: Set lFile = lFSO.GetFile(lTargetFilePath)
    lFile.Copy aDestinationFolder, True
  Next lTargetFilePath
End Sub