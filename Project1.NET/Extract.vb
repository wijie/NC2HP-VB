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
		'UPGRADE_WARNING: 配列 varSub の下限が 44 から 0 に変更されました。 詳細については次のリンクをクリックしてください : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1033"'
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
		'UPGRADE_WARNING: オブジェクト objBar.Visible の既定プロパティを解決できませんでした。 詳細については次のリンクをクリックしてください : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		objBar.Visible = True
		'UPGRADE_WARNING: オブジェクト objBar.Max の既定プロパティを解決できませんでした。 詳細については次のリンクをクリックしてください : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		objBar.Max = 100
		'UPGRADE_WARNING: オブジェクト objBar.Min の既定プロパティを解決できませんでした。 詳細については次のリンクをクリックしてください : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		objBar.Min = 0
		
		'NCを読み込む
		intF0 = FreeFile
		FileOpen(intF0, udtNCInfo.strNCName, OpenMode.Binary)
		ReDim bytBuf(LOF(intF0))
		'UPGRADE_WARNING: Get は、FileGet にアップグレードされ、新しい動作が指定されました。 詳細については次のリンクをクリックしてください : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1041"'
		FileGet(intF0, bytBuf)
		FileClose(intF0)
		'UPGRADE_ISSUE: 定数 vbUnicode はアップグレードされませんでした。 詳細については次のリンクをクリックしてください : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2070"'
		strNC = StrConv(System.Text.UnicodeEncoding.Unicode.GetString(bytBuf), vbUnicode)
		Erase bytBuf '配列のメモリを開放する
		
		'改行コードを調べる
		If InStr(strNC, vbCrLf) > 0 Then
			strEnter = vbCrLf
		ElseIf InStr(strNC, vbLf) > 0 Then 
			strEnter = vbLf
		ElseIf InStr(strNC, vbCr) > 0 Then 
			strEnter = vbCr
		End If
		
		'削除する文字列を処理する
		strNC = Replace(strNC, " ", "")
		'メイン,サブに分割する
		strMainSub = Split(strNC, "G25", -1, CompareMethod.Text)
		strNC = "" '変数のメモリを開放する
		If UBound(strMainSub) = 1 Then
			strSubTmp = Split(strMainSub(0), "N", -1, CompareMethod.Text)
			For i = 1 To UBound(strSubTmp)
				intN = CShort(Left(strSubTmp(i), 2)) 'サブメモリの番号を取得
				'UPGRADE_WARNING: オブジェクト varSub(intN) の既定プロパティを解決できませんでした。 詳細については次のリンクをクリックしてください : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				varSub(intN) = Split(strSubTmp(i), strEnter, -1, CompareMethod.Binary)
			Next 
			strMain = Split(strMainSub(1), strEnter, -1, CompareMethod.Binary)
		Else
			strMain = Split(strMainSub(0), strEnter, -1, CompareMethod.Binary)
		End If
		'配列のメモリを開放する
		Erase strMainSub
		Erase strSubTmp
		
		strOutFile = fTempPath & "NC2HP._$$"
		'出力する
		'UPGRADE_WARNING: オブジェクト objBar.Visible の既定プロパティを解決できませんでした。 詳細については次のリンクをクリックしてください : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		objBar.Visible = True
		'UPGRADE_WARNING: オブジェクト objBar.Max の既定プロパティを解決できませんでした。 詳細については次のリンクをクリックしてください : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		objBar.Max = 100
		lngCount = UBound(strMain)
		intF1 = FreeFile
		FileOpen(intF1, strOutFile, OpenMode.Output)
		For i = 0 To UBound(strMain)
			If strMain(i) Like "X*Y*" = True Then
				strXY = Split(Mid(strMain(i), 2), "Y", -1, CompareMethod.Text)
				If blnDrillHit = True Then
					With udtNCInfo
						'UPGRADE_WARNING: オブジェクト udtNCInfo.varDrl_Inf()() の既定プロパティを解決できませんでした。 詳細については次のリンクをクリックしてください : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
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
						'UPGRADE_WARNING: オブジェクト varSub(intSubNo)(j) の既定プロパティを解決できませんでした。 詳細については次のリンクをクリックしてください : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
						If varSub(intSubNo)(j) Like "X*Y*" = True Then
							'UPGRADE_WARNING: オブジェクト varSub()() の既定プロパティを解決できませんでした。 詳細については次のリンクをクリックしてください : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
							strXY = Split(Mid(varSub(intSubNo)(j), 2), "Y", -1, CompareMethod.Text)
							If blnDrillHit = True Then
								With udtNCInfo
									'UPGRADE_WARNING: オブジェクト udtNCInfo.varDrl_Inf()() の既定プロパティを解決できませんでした。 詳細については次のリンクをクリックしてください : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
									.varDrl_Inf(intIndex)(3) = CInt(.varDrl_Inf(intIndex)(3)) + 1
								End With
							End If
							WriteLine(intF1, strXY(X), strXY(Y))
							'UPGRADE_WARNING: オブジェクト varSub(intSubNo)(j) の既定プロパティを解決できませんでした。 詳細については次のリンクをクリックしてください : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
						ElseIf varSub(intSubNo)(j) Like "G81" = True Then 
							blnDrillHit = True
							WriteLine(intF1, "G81", "")
							'UPGRADE_WARNING: オブジェクト varSub(intSubNo)(j) の既定プロパティを解決できませんでした。 詳細については次のリンクをクリックしてください : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
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
						'UPGRADE_WARNING: オブジェクト udtNCInfo.varDrl_Inf(intIndex)(0) の既定プロパティを解決できませんでした。 詳細については次のリンクをクリックしてください : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
						If intTool = .varDrl_Inf(intIndex)(0) Then Exit For
					Next 
				End With
				WriteLine(intF1, "T" & intTool, "")
			End If
			'        objBar.Value = Int(i / lngCount * 100)
		Next 
		FileClose(intF1)
		Erase strMain '配列のメモリを開放する
		'UPGRADE_WARNING: オブジェクト objBar.Visible の既定プロパティを解決できませんでした。 詳細については次のリンクをクリックしてください : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		objBar.Visible = False
		
	End Sub
End Module