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
		
		'DEFファイルを読み込む
		intFNo0 = FreeFile
		FileOpen(intFNo0, fTempPath & "NC2HPGL.TBL", OpenMode.Binary)
		ReDim bytBuf(LOF(intFNo0))
		'UPGRADE_WARNING: Get は、FileGet にアップグレードされ、新しい動作が指定されました。 詳細については次のリンクをクリックしてください : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1041"'
		FileGet(intFNo0, bytBuf)
		FileClose(intFNo0)
		'UPGRADE_ISSUE: 定数 vbUnicode はアップグレードされませんでした。 詳細については次のリンクをクリックしてください : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2070"'
		strConfig = StrConv(System.Text.UnicodeEncoding.Unicode.GetString(bytBuf), vbUnicode)
		
		strValue = Split(strConfig, vbCrLf, -1, CompareMethod.Text)
		With gudtNCInfo(TH)
			.strNCName = strValue(1) 'THのファイル名
			strTmpArray1 = Split(strValue(2), " ", -1, CompareMethod.Text)
			ReDim .varDrl_Inf(UBound(strTmpArray1))
			For j = 0 To UBound(strTmpArray1)
				strTmpArray2 = Split(strTmpArray1(j), ":", -1, CompareMethod.Text)
				'UPGRADE_WARNING: Array に新しい動作が指定されています。 詳細については次のリンクをクリックしてください : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1041"'
				'UPGRADE_WARNING: オブジェクト gudtNCInfo().varDrl_Inf(j) の既定プロパティを解決できませんでした。 詳細については次のリンクをクリックしてください : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				.varDrl_Inf(j) = New Object(){CShort(Mid(strTmpArray2(0), 3)), strTmpArray2(1), strTmpArray2(2), 0} '最後のは穴数
			Next 
			.strWB_Inf = Split(strValue(3), ":", -1, CompareMethod.Text)
			.strSosu = strValue(4) '"Dual" or "Multi"
		End With
		With gudtNCInfo(NT)
			.strNCName = strValue(5) 'NTのファイル名
			If .strNCName <> "null" Then
				strTmpArray1 = Split(strValue(6), " ", -1, CompareMethod.Text)
				For j = 0 To UBound(strTmpArray1)
					strTmpArray2 = Split(strTmpArray1(j), ":", -1, CompareMethod.Text)
					'UPGRADE_WARNING: Array に新しい動作が指定されています。 詳細については次のリンクをクリックしてください : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1041"'
					'UPGRADE_WARNING: オブジェクト gudtNCInfo().varDrl_Inf(j) の既定プロパティを解決できませんでした。 詳細については次のリンクをクリックしてください : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
					.varDrl_Inf(j) = New Object(){CShort(Mid(strTmpArray2(0), 3)), strTmpArray2(1), strTmpArray2(2), 0} '最後のは穴数
				Next 
				.strWB_Inf = VB6.CopyArray(gudtNCInfo(TH).strWB_Inf)
				.strSosu = gudtNCInfo(TH).strSosu '"Dual" or "Multi"
			End If
		End With
		
	End Sub
End Module