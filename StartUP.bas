Attribute VB_Name = "StartUP"
Option Explicit

Public Const X As Integer = 0
Public Const Y As Integer = 1
Public Const R As Integer = 2
Public Const TH As Integer = 0
Public Const NT As Integer = 1

Public Type NCInfo
    strNCName As String
    varDrl_Inf() As Variant
    strWB_Inf() As String
    strSosu As String
End Type

Public gudtNCInfo(1) As NCInfo

Sub Main()

    If Command = "" Then
        MsgBox "�t�@�C�������w�肵�Ă�������"
        End
    End If

    Load Form1
    Form1.Show

End Sub

Public Function fMyPath() As String
    '�v���O�����I���܂Ł@MyPath�@�̓��e��ێ�
    Static MyPath As String
    '�r���Ńf�B���N�g��-���ύX����Ă��N���f�B���N�g��-���m��
    If Len(MyPath) = 0& Then
        MyPath = App.Path         '�f�B���N�g��-���擾
        '���[�g�f�B���N�g���[���̔��f
        If Right$(MyPath, 1&) <> "\" Then
            MyPath = MyPath & "\"
        End If
    End If
    fMyPath = MyPath
End Function

Public Function fTempPath() As String
    '�v���O�����I���܂Ł@TempPath�@�̓��e��ێ�
    Static TempPath As String
    '�r���Ńf�B���N�g��-���ύX����Ă�Temp�f�B���N�g��-���m��
    If Len(TempPath) = 0& Then
        TempPath = Environ("TEMP")         '�f�B���N�g��-���擾
        '���[�g�f�B���N�g���[���̔��f
        If Right$(TempPath, 1&) <> "\" Then
            TempPath = TempPath & "\"
        End If
    End If
    fTempPath = TempPath
End Function

