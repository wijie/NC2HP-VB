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
        MsgBox "ファイル名を指定してください"
        End
    End If

    Load Form1
    Form1.Show

End Sub

Public Function fMyPath() As String
    'プログラム終了まで　MyPath　の内容を保持
    Static MyPath As String
    '途中でディレクトリ-が変更されても起動ディレクトリ-を確保
    If Len(MyPath) = 0& Then
        MyPath = App.Path         'ディレクトリ-を取得
        'ルートディレクトリーかの判断
        If Right$(MyPath, 1&) <> "\" Then
            MyPath = MyPath & "\"
        End If
    End If
    fMyPath = MyPath
End Function

Public Function fTempPath() As String
    'プログラム終了まで　TempPath　の内容を保持
    Static TempPath As String
    '途中でディレクトリ-が変更されてもTempディレクトリ-を確保
    If Len(TempPath) = 0& Then
        TempPath = Environ("TEMP")         'ディレクトリ-を取得
        'ルートディレクトリーかの判断
        If Right$(TempPath, 1&) <> "\" Then
            TempPath = TempPath & "\"
        End If
    End If
    fTempPath = TempPath
End Function

