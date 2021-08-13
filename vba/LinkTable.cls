VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LinkTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ERR_CODE
' 3024 : �t�@�C�������݂܂���B
' 3011 : �I�u�W�F�N�g��������܂���B
' 3078 : �N�G���e�[�u�������݂��܂���B
' 3051 : �t�@�C�����J���܂���ł����B���̃t�@�C���́A�ǂݎ���p���A�܂��͓ǂݎ���p�̋��L�f�B���N�g���ɕۑ�����Ă��܂��B
' 3027 : �e�[�u���ɍă����N�ł��܂���ł����B�ǂݎ���p���A�܂��͓ǂݎ���p�̋��L�f�B���N�g���ɕۑ�����Ă��܂��B
Public ERR_CODE As Integer
Public ERR_MSG As String
Public ERR_TABLE As String

'//////////////////////////////////////

'�e�[�u�������p�\���`�F�b�N
Public Function CheckAll(Optional ByVal testTableName As String = "") As Boolean
On Error GoTo ex
    CheckAll = False
    ERR_CODE = 0
    ERR_MSG = ""
    ERR_TABLE = ""

    Dim rs As Recordset
    Dim td As TableDef
    Dim sql0, sql1 As String
    sql0 = "SELECT TOP 1 0 AS [test] FROM [": sql1 = "]"
    
    If testTableName <> "" Then
        ERR_TABLE = testTableName
        Set rs = CurrentDb.OpenRecordset(sql0 & testTableName & sql1)
        ERR_TABLE = ""
    Else
        For Each td In CurrentDb.TableDefs
            If Len(td.Connect) > 0 Then
                ERR_TABLE = td.name
                Set rs = CurrentDb.OpenRecordset(sql0 & td.name & sql1)
                Call rs.Close
                Set rs = Nothing
                ERR_TABLE = ""
            End If
        Next
    End If
    CheckAll = True
    
    If Not rs Is Nothing Then: rs.Close
    Set rs = Nothing
    Set td = Nothing
    Exit Function
ex:
    ERR_CODE = Err
    ERR_MSG = Err.Description
    If Not rs Is Nothing Then: rs.Close
    Set rs = Nothing
    Set td = Nothing
End Function

'�����N�e�[�u���̍X�V
Public Function UpdateAll(Optional ByVal linkFile As String = "") As Boolean
    UpdateAll = True
    ERR_CODE = 0
    ERR_MSG = ""
    ERR_TABLE = ""
    
    If linkFile = "" Then
        linkFile = LinkOpenDialog
    End If
    If linkFile = "" Then
        UpdateAll = False
        Exit Function
    End If
    
    UpdateAll = RefreshLinks(linkFile)
    If Err <> 0 Then
        UpdateAll = False
        Me.ERR_CODE = Err
        Me.ERR_MSG = Err.Description
    End If
End Function


'//////////////////////////////////////
Private Function RefreshLinks(linkFile As String) As Boolean
    RefreshLinks = False
    Err = 0
    
    Dim td As TableDef
    For Each td In CurrentDb.TableDefs
        If Len(td.Connect) > 0 Then
            If Not UpdateConnect(td, linkFile) Then
                '�G���[�̔��������e�[�u����
                ERR_TABLE = td.name
                Set td = Nothing
                Exit Function
            End If
        End If
    Next
    RefreshLinks = True
    Set td = Nothing
End Function
Private Function UpdateConnect(ByRef td As TableDef, ByVal linkFile As String) As Boolean
On Error Resume Next
    Err = 0
    
    td.Connect = ";DATABASE=" & linkFile
    td.RefreshLink
    
    If Err <> 0 Then
        UpdateConnect = False
    Else
        UpdateConnect = True
    End If
End Function

'--- �t�@�C���I���_�C�A���O
Private Function LinkOpenDialog() As String
On Error GoTo ex:
    LinkOpenDialog = ""
    
    Dim ret As String
    Dim d As Object
    Set d = Application.FileDialog(3)
    d.Filters.Clear
    Call d.Filters.Add("Microsoft Access", "*.accdb; *.mdb; *.adp; *.mda; *.accda; *.mde; *.accde; *.adc")
    d.AllowMultiSelect = False
    d.InitialFileName = Application.CurrentProject.FullName
    
    d.Title = "LinkTable - MicrosoftAccess"
    If d.Show Then
        If d.SelectedItems.Count > 0 Then
            ret = d.SelectedItems(1)
            If ret = Application.CurrentProject.FullName Then
                ret = ""
            End If
        End If
    End If
    
    LinkOpenDialog = ret
ex:
End Function