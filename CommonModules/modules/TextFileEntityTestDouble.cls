VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextFileEntityTestDouble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' #############################################################################
'!
'! @brief
'! テキスト ファイルのオープン、クローズ、書き込みなどを行うための実装です。
'!
' #############################################################################

Implements ITextFileEntity

Private TUtl As UnitTestUtils

Private pFilePath As String
Private pIsOpen As Boolean
Private pAsRead As Boolean
Private pAsWrite As Boolean
Private pAsAppend As Boolean
Private pGetReadLock As Boolean
Private pGetWriteLock As Boolean

'* 読み取り回数。ReadLine によりインクリメントされ、CloseFile によりリセットされます。
Public ReadCount As Long

'* 書き込み回数。WriteLine によりインクリメントされます。リセットはされません。
Public WriteCount As Long

'* Initialize の引数→受け取り値辞書。引数をキーとして、True を値とします。
Public Initialize_Results As Dictionary

'* OpenFile の引数→受け取り値辞書。引数をキーとして、True を値とします。
Public OpenFile_Results As Dictionary

'* WriteLine の書き込み数→受け取り値辞書。WriteCount をキーとして、引数を値とします。※ GetValue ではなく、直接辞書を操作して値を確認してください。
Public WriteLine_Results As Dictionary

'* ReadLine の読み取り数→戻り値辞書。ReadCount をキーとして、戻り値を値とします。※ SetValue ではなく、直接辞書を操作して値を格納してください。
Public ReadLine_Values As Dictionary

'* CloseFile の引数→受け取り値辞書。Nothing をキーとして、True を値とします。
Public CloseFile_Results As Dictionary

Private Sub Class_Initialize()
    Set TUtl = New UnitTestUtils

    Set Initialize_Results = New Dictionary
    Set OpenFile_Results = New Dictionary
    Set WriteLine_Results = New Dictionary
    Set ReadLine_Values = New Dictionary
    Set CloseFile_Results = New Dictionary
    
End Sub

'* ファイルのパス。
Public Property Get FilePath() As String
    FilePath = pFilePath
End Property
Private Property Get ITextFileEntity_FilePath() As String
    ITextFileEntity_FilePath = FilePath
End Property

'* 読み取り用か。(未オープンの場合は False)
Public Property Get AsRead() As Boolean
    If pIsOpen Then
        AsRead = pAsRead
    Else
        AsRead = False
    End If
End Property
Private Property Get ITextFileEntity_AsRead() As Boolean
    ITextFileEntity_AsRead = AsRead
End Property

'* 書き込み用か。(未オープンの場合は False)
Public Property Get AsWrite() As Boolean
    If pIsOpen Then
        AsWrite = pAsWrite
    Else
        AsWrite = False
    End If
End Property
Private Property Get ITextFileEntity_AsWrite() As Boolean
    ITextFileEntity_AsWrite = AsWrite
End Property

'* 追記か。(未オープンの場合は False)
Public Property Get AsAppend() As Boolean
    If pIsOpen Then
        AsAppend = pAsAppend
    Else
        AsAppend = False
    End If
End Property
Private Property Get ITextFileEntity_AsAppend() As Boolean
    ITextFileEntity_AsAppend = AsAppend
End Property

'* 読み取りロックを取得しているか。(未オープンの場合は False)
Public Property Get GetReadLock() As Boolean
    If pIsOpen Then
        GetReadLock = pGetReadLock
    Else
        GetReadLock = False
    End If
End Property
Private Property Get ITextFileEntity_GetReadLock() As Boolean
    ITextFileEntity_GetReadLock = GetReadLock
End Property

'* 書き込みロックを取得しているか。(未オープンの場合は False)
Public Property Get GetWriteLock() As Boolean
    If pIsOpen Then
        GetWriteLock = pGetWriteLock
    Else
        GetWriteLock = False
    End If
End Property
Private Property Get ITextFileEntity_GetWriteLock() As Boolean
    ITextFileEntity_GetWriteLock = GetWriteLock
End Property

'* 現在ファイルがオープンされているか。
Public Property Get IsOpen() As Boolean
    IsOpen = pIsOpen
End Property
Private Property Get ITextFileEntity_IsOpen() As Boolean
    ITextFileEntity_IsOpen = IsOpen
End Property

' ファイルの最後か。(未オープンの場合は True)
Public Property Get IsEndOfFile() As Boolean
    If Not pIsOpen Or pAsWrite Then
        IsEndOfFile = True
    Else
        If ReadLine_Values.Exists(ReadCount) Then
            IsEndOfFile = False
        Else
            IsEndOfFile = True
        End If
    End If
End Property
Private Property Get ITextFileEntity_IsEndOfFile() As Boolean
    ITextFileEntity_IsEndOfFile = IsEndOfFile
End Property

'* Initialize のスパイです。
'*
'* @details
'* Call TUtl.SetValue(Initialize_Results, True, FilePath)
Public Sub Initialize(ByVal FilePath As String)
    Call TUtl.SetValue(Initialize_Results, True, FilePath)
    pFilePath = FilePath
End Sub

'* OpenFile のスパイです。引数に基づいて、プロパティも設定します。
'*
'* @details
'* Call TUtl.SetValue(OpenFile_Results, True, AsWrite, AsAppend, GetReadLock, GetWriteLock, Force)
Public Sub OpenFile( _
        Optional ByVal AsWrite As Boolean = False, Optional ByVal AsAppend As Boolean = False, _
        Optional ByVal GetReadLock As Boolean = False, Optional ByVal GetWriteLock As Boolean = False, _
        Optional ByVal Force As Boolean = False)
    
    Call TUtl.SetValue(OpenFile_Results, True, AsWrite, AsAppend, GetReadLock, GetWriteLock, Force)
    
    pAsRead = Not AsWrite
    pAsWrite = AsWrite
    pAsAppend = AsAppend
    pGetReadLock = GetReadLock
    pGetWriteLock = GetWriteLock
    
    Call CloseFile_Results.RemoveAll
    
    pIsOpen = True
End Sub
Private Sub ITextFileEntity_OpenFile( _
        Optional ByVal AsWrite As Boolean = False, Optional ByVal AsAppend As Boolean = False, _
        Optional ByVal GetReadLock As Boolean = False, Optional ByVal GetWriteLock As Boolean = False, _
        Optional ByVal Force As Boolean = False)
    
    Call OpenFile(AsWrite:=AsWrite, AsAppend:=AsAppend, GetReadLock:=GetReadLock, GetWriteLock:=GetWriteLock, Force:=Force)
End Sub

'* ReadLine のスタブです。
'*
'* @details
'* ReadLine = ReadLine_Values(ReadCount)
Public Function ReadLine() As String
    ReadLine = ReadLine_Values(ReadCount)
    ReadCount = ReadCount + 1
End Function
Private Function ITextFileEntity_ReadLine() As String
    ITextFileEntity_ReadLine = ReadLine()
End Function

'* WriteLine のスパイです。
'*
'* @details
'* Call WriteLine_Results.Add(WriteCount, Expression)
Public Sub WriteLine(ByVal Expression As String)
    Call WriteLine_Results.Add(WriteCount, Expression)
    WriteCount = WriteCount + 1
End Sub
Private Sub ITextFileEntity_WriteLine(ByVal Expression As String)
    Call WriteLine(Expression)
End Sub

'* CloseFile のスパイです。
'*
'* @details
'* Call TUtl.SetValue(CloseFile_Results, True, Nothing)
Public Sub CloseFile()
    Call TUtl.SetValue(CloseFile_Results, True, Nothing)
    ReadCount = 0
    pIsOpen = False
End Sub
Private Sub ITextFileEntity_CloseFile()
    Call CloseFile
End Sub


