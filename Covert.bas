Attribute VB_Name = "Covert"
Option Explicit

Private intFNo0 As Integer

Sub sHPConv(typNCInfo As NCInfo)

    Dim i, j As Integer
    Dim intDrl As Integer
    Dim sngSize As Single
    Dim intColor As Integer
    Dim strXY(1) As String
    Dim blnDrillHit As Boolean
    Dim strA As String
    Dim strB As String
    Dim intF1 As Integer

    With typNCInfo
        '出力する
        intFNo0 = FreeFile
        Open "NC2HP.HP" For Output As #intFNo0
        Print #intFNo0, ""
        Print #intFNo0, "DF;"
        Print #intFNo0, ""
        Call sPrtWB(typNCInfo) 'ワークボードのプロット
        Call sPrtToolList(typNCInfo) 'ツールリストのプロット
        intF1 = FreeFile
        Open fTempPath & "NC2HP._$$" For Input As #intF1
        Do While Not EOF(intF1)
            Input #intF1, strA, strB
            If strB <> "" Then
                strXY(X) = CLng(strA)
                strXY(Y) = CLng(strB)
                Print #intFNo0, "PU " & _
                                strXY(X) / 2.5 & "," & _
                                strXY(Y) / 2.5 & ";"
                If blnDrillHit = True Then
                    Print #intFNo0, "CI " & sngSize & ";"
                End If
            ElseIf strA = "G81" Then
                blnDrillHit = True
            ElseIf strA = "G80" Then
                blnDrillHit = False
            ElseIf strA Like "T*" = True Then
                For j = 0 To UBound(.varDrl_Inf)
                    If CInt(Mid(strA, 2)) = CInt(.varDrl_Inf(j)(0)) Then
                        Print #intFNo0, "SP " & .varDrl_Inf(j)(1) & ";"
                        sngSize = (CSng(.varDrl_Inf(j)(2)) / 2 - 0.25) / 0.025
                        If sngSize < 0 Then sngSize = 0
                    End If
                Next
            End If
        Loop
        Print #intFNo0, "SP 1;"
        Print #intFNo0, "PU " & -1 * (2.5 / 0.025) & "," & 2.5 / 0.025 & ";"
        Print #intFNo0, "PD " & 5 / 0.025 & "," & -1 * (5 / 0.025) & ";"
        Print #intFNo0, "PU 0," & 5 / 0.025 & ";"
        Print #intFNo0, "PD " & -1 * (5 / 0.025) & "," & -1 * (5 / 0.025) & ";"
        Print #intFNo0, ""
        Print #intFNo0, "PU;SP 0;"
        Print #intFNo0, ""
        Close #intFNo0
        Close #intF1
    End With

End Sub

Sub sPrtWB(typNCInfo As NCInfo)

    With typNCInfo
        Print #intFNo0, "PA;PU ";
        Print #intFNo0, -1 * .strWB_Inf(0) / 2 / 0.025 & "," & _
                        -1 * .strWB_Inf(1) / 2 / 0.025 & ";"
        Print #intFNo0, "PR;"
        Print #intFNo0, "SP 1;" 'ペン番号1を選択
        Print #intFNo0, "PD " & .strWB_Inf(0) / 0.025 & ",0;"
        Print #intFNo0, "PD 0," & .strWB_Inf(1) / 0.025 & ";"
        Print #intFNo0, "PD " & -1 * .strWB_Inf(0) / 0.025 & ",0;"
        Print #intFNo0, "PD 0," & -1 * .strWB_Inf(1) / 0.025 & ";"
        Print #intFNo0, ""
    End With
End Sub

Sub sPrtToolList(typNCInfo As NCInfo)

    Dim i As Integer
    Dim lngTotal As Long

    With typNCInfo
        Print #intFNo0, "PA;PU ";
        Print #intFNo0, -1 * .strWB_Inf(0) / 2 / 0.025 & "," & _
                        -1 * (.strWB_Inf(1) / 2 + 6.35) / 0.025 & ";"
        Print #intFNo0, "SI.30,.40;LB" & .strNCName & Chr(3)
        Print #intFNo0, ""
        For i = 0 To UBound(.varDrl_Inf)
            Print #intFNo0, "PA;PU ";
            Print #intFNo0, -1 * .strWB_Inf(0) / 2 / 0.025 & "," & _
                            -1 * (.strWB_Inf(1) / 2 + 6.35 + (i + 1) * 5.08) / 0.025 & ";"
            Print #intFNo0, "SP " & .varDrl_Inf(i)(1) & ";"
            Print #intFNo0, "SI.15,.20;LB";
            Print #intFNo0, "T" & Format(.varDrl_Inf(i)(0), "0#") & "/";
            Print #intFNo0, Format(.varDrl_Inf(i)(2), "!@@@@@") & "mm/";
            Print #intFNo0, Format(.varDrl_Inf(i)(3), "@@@@@@") & Chr(3)
            lngTotal = lngTotal + CLng(.varDrl_Inf(i)(3))
        Next
        Print #intFNo0, "PA;PU ";
        Print #intFNo0, -1 * .strWB_Inf(0) / 2 / 0.025 & "," & _
                        -1 * (.strWB_Inf(1) / 2 + 6.35 + (i + 1) * 5.08) / 0.025 & ";"
        Print #intFNo0, "SP 1;"
        Print #intFNo0, "SI.15,.20;LB    Total  /";
        Print #intFNo0, Format(lngTotal, "@@@@@@") & Chr(3)
        Print #intFNo0, ""
        Print #intFNo0, "PA;PU " & -1 * (.strWB_Inf(0) / 2) / 0.025 & ",";
        Print #intFNo0, -1 * (.strWB_Inf(1) / 2) / 0.025 & ";"
        Print #intFNo0, "PR;PU " & .strWB_Inf(2) / 0.025 & ",";
        Print #intFNo0, .strWB_Inf(3) / 0.025 & ";"
        Print #intFNo0, ""
    End With

End Sub
