Option Strict Off
Option Explicit On
Module Covert
	
	Private intFNo0 As Short
	
	Sub sHPConv(ByRef typNCInfo As NCInfo)
		
		Dim i As Object
		Dim j As Short
		Dim intDrl As Short
		Dim sngSize As Single
		Dim intColor As Short
		Dim strXY(1) As String
		Dim blnDrillHit As Boolean
		Dim strA As String
		Dim strB As String
		Dim intF1 As Short
		
		With typNCInfo
			'出力する
			intFNo0 = FreeFile
			FileOpen(intFNo0, "NC2HP.HP", OpenMode.Output)
			PrintLine(intFNo0, "")
			PrintLine(intFNo0, "DF;")
			PrintLine(intFNo0, "")
			Call sPrtWB(typNCInfo) 'ワークボードのプロット
			Call sPrtToolList(typNCInfo) 'ツールリストのプロット
			intF1 = FreeFile
			FileOpen(intF1, fTempPath & "NC2HP._$$", OpenMode.Input)
			Do While Not EOF(intF1)
				Input(intF1, strA)
				Input(intF1, strB)
				If strB <> "" Then
					strXY(X) = CStr(CInt(strA))
					strXY(Y) = CStr(CInt(strB))
					PrintLine(intFNo0, "PU " & CDbl(strXY(X)) / 2.5 & "," & CDbl(strXY(Y)) / 2.5 & ";")
					If blnDrillHit = True Then
						PrintLine(intFNo0, "CI " & sngSize & ";")
					End If
				ElseIf strA = "G81" Then 
					blnDrillHit = True
				ElseIf strA = "G80" Then 
					blnDrillHit = False
				ElseIf strA Like "T*" = True Then 
					For j = 0 To UBound(.varDrl_Inf)
						'UPGRADE_WARNING: オブジェクト typNCInfo.varDrl_Inf()() の既定プロパティを解決できませんでした。 詳細については次のリンクをクリックしてください : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
						If CShort(Mid(strA, 2)) = CShort(.varDrl_Inf(j)(0)) Then
							'UPGRADE_WARNING: オブジェクト typNCInfo.varDrl_Inf(j)(1) の既定プロパティを解決できませんでした。 詳細については次のリンクをクリックしてください : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
							PrintLine(intFNo0, "SP " & .varDrl_Inf(j)(1) & ";")
							'UPGRADE_WARNING: オブジェクト typNCInfo.varDrl_Inf()() の既定プロパティを解決できませんでした。 詳細については次のリンクをクリックしてください : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
							sngSize = (CSng(.varDrl_Inf(j)(2)) / 2 - 0.25) / 0.025
							If sngSize < 0 Then sngSize = 0
						End If
					Next 
				End If
			Loop 
			PrintLine(intFNo0, "SP 1;")
			PrintLine(intFNo0, "PU " & -1 * (2.5 / 0.025) & "," & 2.5 / 0.025 & ";")
			PrintLine(intFNo0, "PD " & 5 / 0.025 & "," & -1 * (5 / 0.025) & ";")
			PrintLine(intFNo0, "PU 0," & 5 / 0.025 & ";")
			PrintLine(intFNo0, "PD " & -1 * (5 / 0.025) & "," & -1 * (5 / 0.025) & ";")
			PrintLine(intFNo0, "")
			PrintLine(intFNo0, "PU;SP 0;")
			PrintLine(intFNo0, "")
			FileClose(intFNo0)
			FileClose(intF1)
		End With
		
	End Sub
	
	Sub sPrtWB(ByRef typNCInfo As NCInfo)
		
		With typNCInfo
			Print(intFNo0, "PA;PU ")
			PrintLine(intFNo0, -1 * CDbl(.strWB_Inf(0)) / 2 / 0.025 & "," & -1 * CDbl(.strWB_Inf(1)) / 2 / 0.025 & ";")
			PrintLine(intFNo0, "PR;")
			PrintLine(intFNo0, "SP 1;") 'ペン番号1を選択
			PrintLine(intFNo0, "PD " & CDbl(.strWB_Inf(0)) / 0.025 & ",0;")
			PrintLine(intFNo0, "PD 0," & CDbl(.strWB_Inf(1)) / 0.025 & ";")
			PrintLine(intFNo0, "PD " & -1 * CDbl(.strWB_Inf(0)) / 0.025 & ",0;")
			PrintLine(intFNo0, "PD 0," & -1 * CDbl(.strWB_Inf(1)) / 0.025 & ";")
			PrintLine(intFNo0, "")
		End With
	End Sub
	
	Sub sPrtToolList(ByRef typNCInfo As NCInfo)
		
		Dim i As Short
		Dim lngTotal As Integer
		
		With typNCInfo
			Print(intFNo0, "PA;PU ")
			PrintLine(intFNo0, -1 * CDbl(.strWB_Inf(0)) / 2 / 0.025 & "," & -1 * (CDbl(.strWB_Inf(1)) / 2 + 6.35) / 0.025 & ";")
			PrintLine(intFNo0, "SI.30,.40;LB" & .strNCName & Chr(3))
			PrintLine(intFNo0, "")
			For i = 0 To UBound(.varDrl_Inf)
				Print(intFNo0, "PA;PU ")
				PrintLine(intFNo0, -1 * CDbl(.strWB_Inf(0)) / 2 / 0.025 & "," & -1 * (CDbl(.strWB_Inf(1)) / 2 + 6.35 + (i + 1) * 5.08) / 0.025 & ";")
				'UPGRADE_WARNING: オブジェクト typNCInfo.varDrl_Inf(i)(1) の既定プロパティを解決できませんでした。 詳細については次のリンクをクリックしてください : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				PrintLine(intFNo0, "SP " & .varDrl_Inf(i)(1) & ";")
				Print(intFNo0, "SI.15,.20;LB")
				'UPGRADE_WARNING: オブジェクト typNCInfo.varDrl_Inf()() の既定プロパティを解決できませんでした。 詳細については次のリンクをクリックしてください : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				Print(intFNo0, "T" & VB6.Format(.varDrl_Inf(i)(0), "0#") & "/")
				'UPGRADE_WARNING: オブジェクト typNCInfo.varDrl_Inf()() の既定プロパティを解決できませんでした。 詳細については次のリンクをクリックしてください : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				Print(intFNo0, VB6.Format(.varDrl_Inf(i)(2), "!@@@@@") & "mm/")
				'UPGRADE_WARNING: オブジェクト typNCInfo.varDrl_Inf()() の既定プロパティを解決できませんでした。 詳細については次のリンクをクリックしてください : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				PrintLine(intFNo0, VB6.Format(.varDrl_Inf(i)(3), "@@@@@@") & Chr(3))
				'UPGRADE_WARNING: オブジェクト typNCInfo.varDrl_Inf()() の既定プロパティを解決できませんでした。 詳細については次のリンクをクリックしてください : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				lngTotal = lngTotal + CInt(.varDrl_Inf(i)(3))
			Next 
			Print(intFNo0, "PA;PU ")
			PrintLine(intFNo0, -1 * CDbl(.strWB_Inf(0)) / 2 / 0.025 & "," & -1 * (CDbl(.strWB_Inf(1)) / 2 + 6.35 + (i + 1) * 5.08) / 0.025 & ";")
			PrintLine(intFNo0, "SP 1;")
			Print(intFNo0, "SI.15,.20;LB    Total  /")
			PrintLine(intFNo0, VB6.Format(lngTotal, "@@@@@@") & Chr(3))
			PrintLine(intFNo0, "")
			Print(intFNo0, "PA;PU " & -1 * (CDbl(.strWB_Inf(0)) / 2) / 0.025 & ",")
			PrintLine(intFNo0, -1 * (CDbl(.strWB_Inf(1)) / 2) / 0.025 & ";")
			Print(intFNo0, "PR;PU " & CDbl(.strWB_Inf(2)) / 0.025 & ",")
			PrintLine(intFNo0, CDbl(.strWB_Inf(3)) / 0.025 & ";")
			PrintLine(intFNo0, "")
		End With
		
	End Sub
End Module