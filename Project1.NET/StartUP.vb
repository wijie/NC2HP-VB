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
	
	'UPGRADE_WARNING: Sub Main() �����������Ƃ��ɃA�v���P�[�V�����͏I�����܂��B �ڍׂɂ��Ă͎��̃����N���N���b�N���Ă������� : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1047"'
	Public Sub Main()
		
		If VB.Command() = "" Then
			MsgBox("�t�@�C�������w�肵�Ă�������")
			End
		End If
		
		'UPGRADE_ISSUE: Load �X�e�[�g�����g �̓T�|�[�g����Ă��܂���B �ڍׂɂ��Ă͎��̃����N���N���b�N���Ă������� : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1039"'
		Load(Form1)
		Form1.DefInstance.Show()
		
	End Sub
	
	Public Function fMyPath() As String
		'�v���O�����I���܂Ł@MyPath�@�̓��e��ێ�
		Static MyPath As String
		'�r���Ńf�B���N�g��-���ύX����Ă��N���f�B���N�g��-���m��
		If Len(MyPath) = 0 Then
			MyPath = VB6.GetPath '�f�B���N�g��-���擾
			'���[�g�f�B���N�g���[���̔��f
			If Right(MyPath, 1) <> "\" Then
				MyPath = MyPath & "\"
			End If
		End If
		fMyPath = MyPath
	End Function
	
	Public Function fTempPath() As String
		'�v���O�����I���܂Ł@TempPath�@�̓��e��ێ�
		Static TempPath As String
		'�r���Ńf�B���N�g��-���ύX����Ă�Temp�f�B���N�g��-���m��
		If Len(TempPath) = 0 Then
			TempPath = Environ("TEMP") '�f�B���N�g��-���擾
			'���[�g�f�B���N�g���[���̔��f
			If Right(TempPath, 1) <> "\" Then
				TempPath = TempPath & "\"
			End If
		End If
		fTempPath = TempPath
	End Function
End Module