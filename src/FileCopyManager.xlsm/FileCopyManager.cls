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
Private mCustomString As CustomString


Public Sub Class_Initialize()
  Set mCustomString = New CustomString
End Sub

' 指定のフォルダのブック名を、コレクションに格納して返す
' 引数には、フォルダを絶対パスで、末尾に "\" をつけて返す
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

' 指定のフォルダ内のフォルダ名を、コレクションに格納して返す
' 引数には、フォルダを絶対パスで、末尾に"\"をつけて渡す
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

' このブックのフォルダ名を絶対パスで得る
Private Function GetFolderPathOfThisWorkbook() As String
  GetFolderPathOfThisWorkbook = Left(ActiveWorkbook.FullName, InStrRev(ActiveWorkbook.FullName, "\"))
End Function
Public Function Test_GetFolderPathOfThisWorkbook() As String
  Test_GetFolderPathOfThisWorkbook = GetFolderPathOfThisWorkbook
End Function

' 指定のフォルダ内のブックのうち、指定の文言を含むブック名のみを Collection に格納して返す
' Excel ブック以外のファイル名は返さない
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

' 検索対象のフォルダに指定の文言を含むファイル名があった場合
' それらを絶対パスでコレクションに格納して返す
' aFolderPath … 検索対象とするフォルダ
' aKeyWord … 検索対象とする文言
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
' 上記プロシージャのコレクション対応版
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

' 検索対象のフォルダにサブフォルダがあった場合
' それらを絶対パスでコレクションに格納して返す
' aFolderPath … 検索対象とするフォルダ
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
' 上記プロシージャのコレクション対応版
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

' 起点フォルダのサブフォルダに対し、指定の文言を含むファイルが無いか検索
' ファイルがあれば、その絶対パスをコレクションに格納して返す
' aRootFolder … 起点フォルダ。絶対パスで入力(末尾に"\"を付けること)
' aKeyWord … この文言を含むファイルを検索する
' aDepth … 何階層下のサブフォルダまで検索を行うか指定。起点フォルダ内のみ検索したい場合、0 を指定する
Private Function FindFiles(ByVal aRootFolder As String, ByVal aKeyWord As String, ByVal aDepth As Long) As Collection
  Set FindFiles = New Collection
  Dim lFolderPaths As Collection: Set lFolderPaths = New Collection: Call lFolderPaths.Add(aRootFolder)
  Dim i As Long
  For i = 0 To aDepth
    ' フォルダ内のファイルに対して検索処理
    Dim lFilePath As Variant
    For Each lFilePath In GetAllFilePathsContainThisWord2(lFolderPaths, aKeyWord)
      Call FindFiles.Add(lFilePath)
    Next lFilePath
    
    ' サブフォルダに対しての処理
    Dim lTempFolderPaths As Collection: Set lTempFolderPaths = New Collection
    Dim lFolderPath As Variant
    For Each lFolderPath In GetAllFolderPaths2(lFolderPaths)
        Call lTempFolderPaths.Add(lFolderPath)
    Next lFolderPath
    
    ' 次のループ用に変数の値を更新
    Set lFolderPaths = lTempFolderPaths
  Next i
End Function
Public Function Test_FindFiles(ByVal aRootFolder As String, ByVal aKeyWord As String, ByVal aDepth As Long) As Collection
  Set Test_FindFiles = FindFiles(aRootFolder, aKeyWord, aDepth)
End Function

' フォルダ内のファイルをすべて削除する
' aTargetFolder … 対象フォルダ。絶対パスで入力(末尾に"\"を付けること)
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

' 起点フォルダのサブフォルダに対し、指定の文言を含むファイルが無いか検索
' ファイルがあれば、コピー先フォルダへコピーする
' コピー先に同名ファイルがある場合は上書きする
' aRootFolder … 起点フォルダ。絶対パスで入力(末尾に"\"を付けること)
' aKeyWord … この文言を含むファイルを検索する
' aDepth … 何階層下のサブフォルダまで検索を行うか指定。起点フォルダ内のみ検索したい場合、0 を指定する
' aDestinationFolder … コピー先フォルダ。絶対パスで入力(末尾に"\"を付けること)
Public Sub Main(ByVal aRootFolder As String, ByVal aKeyWord As String, ByVal aDepth As Long, ByVal aDestinationFolder As String)
  Dim lFSO As Object: Set lFSO = CreateObject("Scripting.FileSystemObject")
  Dim lTargetFilePath As Variant
  For Each lTargetFilePath In FindFiles(aRootFolder, aKeyWord, aDepth)
    Dim lFile As Object: Set lFile = lFSO.GetFile(lTargetFilePath)
    lFile.Copy aDestinationFolder, True
  Next lTargetFilePath
End Sub

' 引数に取ったファイルパスに対し、変更を加えて返す
' コピー先フォルダのパス & "親フォルダ1の名前" & "_" & "親フォルダ2の名前" & "_" & ファイル名
' (コピー元のファイル名には編集を加えない)
' aFilePath … コピー元ファイル。絶対パスで入力
' aParent1Depth … 親フォルダ1の深度。一つ上のフォルダを指定したいときは1、二つ上は2のように指定する
' aParent2Depth … 親フォルダ2の深度。設定方法は親フォルダ1の場合と同じ
' aTargetFolderPath … コピー先フォルダのパス。絶対パスで入力(末尾に"\"を付けること)
Private Function GetConvertedFilePath(ByVal aFilePath As String, ByVal aParent1Depth As Long, ByVal aParent2Depth As Long, ByVal aTargetFolderPath As String) As String
  ' under construction
End Function

' 引数に取ったファイルパスに対し、指定の数だけ上位のフォルダ名を返す
' aFilePath … ファイルパス。絶対パスで入力
' aParentDepth … 親フォルダ1の深度。一つ上のフォルダを指定したいときは1、二つ上は2のように指定する
Private Function GetParentFolderName(ByVal aFilePath As String, ByVal aParentDepth As Long) As String
  Const DELIMITER As String = "\"
  GetParentFolderName = mCustomString.GetSplitedString(aFilePath, DELIMITER, mCustomString.GetCollectionFromDelimitedStrings(aFilePath, DELIMITER).Count - aParentDepth - 1)
End Function
Public Function Test_GetParentFolderName(ByVal aFilePath As String, ByVal aParentDepth As Long) As String
  Test_GetParentFolderName = GetParentFolderName(aFilePath, aParentDepth)
End Function

' 引数にとったファイルパスより、ファイル名を抜き出して返す
' フォルダパスを受け取った場合、空文字を返す
Private Function GetFileNameFromFilePath(ByVal aFilePath As String) As String
  Const DELIMITER As String = "\"
  GetFileNameFromFilePath = IIf(Right(aFilePath, 1) = DELIMITER, "", mCustomString.GetSplitedString(aFilePath, DELIMITER, mCustomString.GetCollectionFromDelimitedStrings(aFilePath, DELIMITER).Count - 1))
End Function
Public Function Test_GetFileNameFromFilePath(ByVal aFilePath As String) As String
  Test_GetFileNameFromFilePath = GetFileNameFromFilePath(aFilePath)
End Function
