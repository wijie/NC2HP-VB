Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Module StartUP
	
	Public Const X As Short = 0
	Public Const Y As Short = 1
	Public Const R As Short = 2
	Public Const TH As Short = 0
	Public Const NT As Short = 1
	
	Public Structure NCInfo
		Dim strNCName As String
		Dim varDrl_Inf() As Object
		Dim strWB_Inf() As String
		Dim strSosu As String
	End Structure
	
	Public gudtNCInfo(1) As NCInfo
	
	'UPGRADE_WARNING: Sub Main() が完了したときにアプリケーションは終了します。 詳細については次のリンクをクリックしてください : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1047"'
	Public Sub Main()
		
		If VB.Command() = "" Then
			MsgBox("ファイル名を指定してください")
			End
		End If
		
		'UPGRADE_ISSUE: Load ステートメント はサポートされていません。 詳細については次のリンクをクリックしてください : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1039"'
		Load(Form1)
		Form1.DefInstance.Show()
		
	End Sub
	
	Public Function fMyPath() As String
		'プログラム終了まで　MyPath　の内容を保持
		Static MyPath As String
		'途中でディレクトリ-が変更されても起動ディレクトリ-を確保
		If Len(MyPath) = 0 Then
			MyPath = VB6.GetPath 'ディレクトリ-を取得
			'ルートディレクトリーかの判断
			If Right(MyPath, 1) <> "\" Then
				MyPath = MyPath & "\"
			End If
		End If
		fMyPath = MyPath
	End Function
	
	Public Function fTempPath() As String
		'プログラム終了まで　TempPath　の内容を保持
		Static TempPath As String
		'途中でディレクトリ-が変更されてもTempディレクトリ-を確保
		If Len(TempPath) = 0 Then
			TempPath = Environ("TEMP") 'ディレクトリ-を取得
			'ルートディレクトリーかの判断
			If Right(TempPath, 1) <> "\" Then
				TempPath = TempPath & "\"
			End If
		End If
		fTempPath = TempPath
	End Function
End Module