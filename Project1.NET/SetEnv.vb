Option Strict Off
Option Explicit On
Module SetEnv
	
	Sub sReadConfig()
		
		Dim intFNo0 As Short
		Dim strValue() As String
		Dim strTmpArray1() As String
		Dim strTmpArray2() As String
		Dim bytBuf() As Byte
		Dim strConfig As String
		Dim i As Short
		Dim j As Short
		
		'DEF�t�@�C����ǂݍ���
		intFNo0 = FreeFile
		FileOpen(intFNo0, fTempPath & "NC2HPGL.TBL", OpenMode.Binary)
		ReDim bytBuf(LOF(intFNo0))
		'UPGRADE_WARNING: Get �́AFileGet �ɃA�b�v�O���[�h����A�V�������삪�w�肳��܂����B �ڍׂɂ��Ă͎��̃����N���N���b�N���Ă������� : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1041"'
		FileGet(intFNo0, bytBuf)
		FileClose(intFNo0)
		'UPGRADE_ISSUE: �萔 vbUnicode �̓A�b�v�O���[�h����܂���ł����B �ڍׂɂ��Ă͎��̃����N���N���b�N���Ă������� : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2070"'
		strConfig = StrConv(System.Text.UnicodeEncoding.Unicode.GetString(bytBuf), vbUnicode)
		
		strValue = Split(strConfig, vbCrLf, -1, CompareMethod.Text)
		With gudtNCInfo(TH)
			.strNCName = strValue(1) 'TH�̃t�@�C����
			strTmpArray1 = Split(strValue(2), " ", -1, CompareMethod.Text)
			ReDim .varDrl_Inf(UBound(strTmpArray1))
			For j = 0 To UBound(strTmpArray1)
				strTmpArray2 = Split(strTmpArray1(j), ":", -1, CompareMethod.Text)
				'UPGRADE_WARNING: Array �ɐV�������삪�w�肳��Ă��܂��B �ڍׂɂ��Ă͎��̃����N���N���b�N���Ă������� : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1041"'
				'UPGRADE_WARNING: �I�u�W�F�N�g gudtNCInfo().varDrl_Inf(j) �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��Ă͎��̃����N���N���b�N���Ă������� : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				.varDrl_Inf(j) = New Object(){CShort(Mid(strTmpArray2(0), 3)), strTmpArray2(1), strTmpArray2(2), 0} '�Ō�̂͌���
			Next 
			.strWB_Inf = Split(strValue(3), ":", -1, CompareMethod.Text)
			.strSosu = strValue(4) '"Dual" or "Multi"
		End With
		With gudtNCInfo(NT)
			.strNCName = strValue(5) 'NT�̃t�@�C����
			If .strNCName <> "null" Then
				strTmpArray1 = Split(strValue(6), " ", -1, CompareMethod.Text)
				For j = 0 To UBound(strTmpArray1)
					strTmpArray2 = Split(strTmpArray1(j), ":", -1, CompareMethod.Text)
					'UPGRADE_WARNING: Array �ɐV�������삪�w�肳��Ă��܂��B �ڍׂɂ��Ă͎��̃����N���N���b�N���Ă������� : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1041"'
					'UPGRADE_WARNING: �I�u�W�F�N�g gudtNCInfo().varDrl_Inf(j) �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��Ă͎��̃����N���N���b�N���Ă������� : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
					.varDrl_Inf(j) = New Object(){CShort(Mid(strTmpArray2(0), 3)), strTmpArray2(1), strTmpArray2(2), 0} '�Ō�̂͌���
				Next 
				.strWB_Inf = VB6.CopyArray(gudtNCInfo(TH).strWB_Inf)
				.strSosu = gudtNCInfo(TH).strSosu '"Dual" or "Multi"
			End If
		End With
		
	End Sub
End Module