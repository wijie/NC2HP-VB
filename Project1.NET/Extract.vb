Option Strict Off
Option Explicit On
Module Extract
	
	Private strEnter As String
	
	Public Sub sNCExtract(ByRef udtNCInfo As NCInfo, ByRef objBar As Object)
		
		Dim intF0 As Short
		Dim intF1 As Short
		Dim strNC As String
		Dim bytBuf() As Byte
		Dim strMainSub() As String
		'UPGRADE_WARNING: �z�� varSub �̉����� 44 ���� 0 �ɕύX����܂����B �ڍׂɂ��Ă͎��̃����N���N���b�N���Ă������� : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1033"'
		Dim varSub(97) As Object
		Dim strMain() As String
		Dim strSubTmp() As String
		Dim strEnter As String
		Dim intN As Short
		Dim blnDrillHit As Boolean
		Dim i As Integer
		Dim j As Integer
		Dim intIndex As Short
		Dim sngDrl As Single
		Dim strXY() As String
		Dim intDigit As Short
		Dim strOutFile As String
		Dim lngColor As Integer
		Dim intSubNo As Short
		Dim intTool As Short
		Dim lngCount As Integer
		
		blnDrillHit = False
		'UPGRADE_WARNING: �I�u�W�F�N�g objBar.Visible �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��Ă͎��̃����N���N���b�N���Ă������� : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		objBar.Visible = True
		'UPGRADE_WARNING: �I�u�W�F�N�g objBar.Max �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��Ă͎��̃����N���N���b�N���Ă������� : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		objBar.Max = 100
		'UPGRADE_WARNING: �I�u�W�F�N�g objBar.Min �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��Ă͎��̃����N���N���b�N���Ă������� : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		objBar.Min = 0
		
		'NC��ǂݍ���
		intF0 = FreeFile
		FileOpen(intF0, udtNCInfo.strNCName, OpenMode.Binary)
		ReDim bytBuf(LOF(intF0))
		'UPGRADE_WARNING: Get �́AFileGet �ɃA�b�v�O���[�h����A�V�������삪�w�肳��܂����B �ڍׂɂ��Ă͎��̃����N���N���b�N���Ă������� : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1041"'
		FileGet(intF0, bytBuf)
		FileClose(intF0)
		'UPGRADE_ISSUE: �萔 vbUnicode �̓A�b�v�O���[�h����܂���ł����B �ڍׂɂ��Ă͎��̃����N���N���b�N���Ă������� : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2070"'
		strNC = StrConv(System.Text.UnicodeEncoding.Unicode.GetString(bytBuf), vbUnicode)
		Erase bytBuf '�z��̃��������J������
		
		'���s�R�[�h�𒲂ׂ�
		If InStr(strNC, vbCrLf) > 0 Then
			strEnter = vbCrLf
		ElseIf InStr(strNC, vbLf) > 0 Then 
			strEnter = vbLf
		ElseIf InStr(strNC, vbCr) > 0 Then 
			strEnter = vbCr
		End If
		
		'�폜���镶�������������
		strNC = Replace(strNC, " ", "")
		'���C��,�T�u�ɕ�������
		strMainSub = Split(strNC, "G25", -1, CompareMethod.Text)
		strNC = "" '�ϐ��̃��������J������
		If UBound(strMainSub) = 1 Then
			strSubTmp = Split(strMainSub(0), "N", -1, CompareMethod.Text)
			For i = 1 To UBound(strSubTmp)
				intN = CShort(Left(strSubTmp(i), 2)) '�T�u�������̔ԍ����擾
				'UPGRADE_WARNING: �I�u�W�F�N�g varSub(intN) �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��Ă͎��̃����N���N���b�N���Ă������� : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				varSub(intN) = Split(strSubTmp(i), strEnter, -1, CompareMethod.Binary)
			Next 
			strMain = Split(strMainSub(1), strEnter, -1, CompareMethod.Binary)
		Else
			strMain = Split(strMainSub(0), strEnter, -1, CompareMethod.Binary)
		End If
		'�z��̃��������J������
		Erase strMainSub
		Erase strSubTmp
		
		strOutFile = fTempPath & "NC2HP._$$"
		'�o�͂���
		'UPGRADE_WARNING: �I�u�W�F�N�g objBar.Visible �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��Ă͎��̃����N���N���b�N���Ă������� : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		objBar.Visible = True
		'UPGRADE_WARNING: �I�u�W�F�N�g objBar.Max �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��Ă͎��̃����N���N���b�N���Ă������� : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		objBar.Max = 100
		lngCount = UBound(strMain)
		intF1 = FreeFile
		FileOpen(intF1, strOutFile, OpenMode.Output)
		For i = 0 To UBound(strMain)
			If strMain(i) Like "X*Y*" = True Then
				strXY = Split(Mid(strMain(i), 2), "Y", -1, CompareMethod.Text)
				If blnDrillHit = True Then
					With udtNCInfo
						'UPGRADE_WARNING: �I�u�W�F�N�g udtNCInfo.varDrl_Inf()() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��Ă͎��̃����N���N���b�N���Ă������� : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
						.varDrl_Inf(intIndex)(3) = CInt(.varDrl_Inf(intIndex)(3)) + 1
					End With
				End If
				WriteLine(intF1, strXY(X), strXY(Y))
			ElseIf strMain(i) Like "G81" = True Then 
				blnDrillHit = True
				WriteLine(intF1, "G81", "")
			ElseIf strMain(i) Like "G80" = True Then 
				blnDrillHit = False
				WriteLine(intF1, "G80", "")
			ElseIf strMain(i) Like "M##" = True Then 
				intSubNo = CShort(Mid(strMain(i), 2))
				If intSubNo >= 44 And intSubNo <= 97 And intSubNo <> 89 Then
					For j = 0 To UBound(varSub(intSubNo))
						'UPGRADE_WARNING: �I�u�W�F�N�g varSub(intSubNo)(j) �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��Ă͎��̃����N���N���b�N���Ă������� : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
						If varSub(intSubNo)(j) Like "X*Y*" = True Then
							'UPGRADE_WARNING: �I�u�W�F�N�g varSub()() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��Ă͎��̃����N���N���b�N���Ă������� : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
							strXY = Split(Mid(varSub(intSubNo)(j), 2), "Y", -1, CompareMethod.Text)
							If blnDrillHit = True Then
								With udtNCInfo
									'UPGRADE_WARNING: �I�u�W�F�N�g udtNCInfo.varDrl_Inf()() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��Ă͎��̃����N���N���b�N���Ă������� : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
									.varDrl_Inf(intIndex)(3) = CInt(.varDrl_Inf(intIndex)(3)) + 1
								End With
							End If
							WriteLine(intF1, strXY(X), strXY(Y))
							'UPGRADE_WARNING: �I�u�W�F�N�g varSub(intSubNo)(j) �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��Ă͎��̃����N���N���b�N���Ă������� : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
						ElseIf varSub(intSubNo)(j) Like "G81" = True Then 
							blnDrillHit = True
							WriteLine(intF1, "G81", "")
							'UPGRADE_WARNING: �I�u�W�F�N�g varSub(intSubNo)(j) �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��Ă͎��̃����N���N���b�N���Ă������� : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
						ElseIf varSub(intSubNo)(j) Like "G80" = True Then 
							blnDrillHit = False
							WriteLine(intF1, "G80", "")
						End If
					Next 
				End If
			ElseIf strMain(i) Like "T*" = True Then 
				intTool = CShort(Mid(strMain(i), 2))
				With udtNCInfo
					For intIndex = 0 To UBound(.varDrl_Inf)
						'UPGRADE_WARNING: �I�u�W�F�N�g udtNCInfo.varDrl_Inf(intIndex)(0) �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��Ă͎��̃����N���N���b�N���Ă������� : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
						If intTool = .varDrl_Inf(intIndex)(0) Then Exit For
					Next 
				End With
				WriteLine(intF1, "T" & intTool, "")
			End If
			'        objBar.Value = Int(i / lngCount * 100)
		Next 
		FileClose(intF1)
		Erase strMain '�z��̃��������J������
		'UPGRADE_WARNING: �I�u�W�F�N�g objBar.Visible �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��Ă͎��̃����N���N���b�N���Ă������� : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		objBar.Visible = False
		
	End Sub
End Module