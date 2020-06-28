Attribute VB_Name = "Extract"
Option Explicit

Private strEnter As String

Public Sub sNCExtract(ByRef udtNCInfo As NCInfo, _
                      ByRef objBar As Object)

    Dim intF0 As Integer
    Dim intF1 As Integer
    Dim strNC As String
    Dim bytBuf() As Byte
    Dim strMainSub() As String
    Dim varSub(44 To 97) As Variant
    Dim strMain() As String
    Dim strSubTmp() As String
    Dim strEnter As String
    Dim intN As Integer
    Dim blnDrillHit As Boolean
    Dim i As Long
    Dim j As Long
    Dim intIndex As Integer
    Dim sngDrl As Single
    Dim strXY() As String
    Dim intDigit As Integer
    Dim strOutFile As String
    Dim lngColor As Long
    Dim intSubNo As Integer
    Dim intTool As Integer
    Dim lngCount As Long

    blnDrillHit = False
    objBar.Visible = True
    objBar.Max = 100
    objBar.Min = 0

    'NCを読み込む
    intF0 = FreeFile
    Open udtNCInfo.strNCName For Binary As #intF0
    ReDim bytBuf(LOF(intF0))
    Get #intF0, , bytBuf
    Close #intF0
    strNC = StrConv(bytBuf, vbUnicode)
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
    strMainSub = Split(strNC, "G25", -1, vbTextCompare)
    strNC = "" '変数のメモリを開放する
    If UBound(strMainSub) = 1 Then
        strSubTmp = Split(strMainSub(0), "N", -1, vbTextCompare)
        For i = 1 To UBound(strSubTmp)
            intN = Left(strSubTmp(i), 2) 'サブメモリの番号を取得
            varSub(intN) = Split(strSubTmp(i), strEnter, -1, vbBinaryCompare)
        Next
        strMain = Split(strMainSub(1), strEnter, -1, vbBinaryCompare)
    Else
        strMain = Split(strMainSub(0), strEnter, -1, vbBinaryCompare)
    End If
    '配列のメモリを開放する
    Erase strMainSub
    Erase strSubTmp

    strOutFile = fTempPath & "NC2HP._$$"
    '出力する
    objBar.Visible = True
    objBar.Max = 100
    lngCount = UBound(strMain)
    intF1 = FreeFile
    Open strOutFile For Output As #intF1
    For i = 0 To UBound(strMain)
        If strMain(i) Like "X*Y*" = True Then
            strXY = Split(Mid(strMain(i), 2), "Y", -1, vbTextCompare)
            If blnDrillHit = True Then
                With udtNCInfo
                    .varDrl_Inf(intIndex)(3) = CLng(.varDrl_Inf(intIndex)(3)) + 1
                End With
            End If
            Write #intF1, strXY(X), strXY(Y)
        ElseIf strMain(i) Like "G81" = True Then
            blnDrillHit = True
            Write #intF1, "G81", ""
        ElseIf strMain(i) Like "G80" = True Then
            blnDrillHit = False
            Write #intF1, "G80", ""
        ElseIf strMain(i) Like "M##" = True Then
            intSubNo = CInt(Mid(strMain(i), 2))
            If intSubNo >= 44 And intSubNo <= 97 And intSubNo <> 89 Then
                For j = 0 To UBound(varSub(intSubNo))
                    If varSub(intSubNo)(j) Like "X*Y*" = True Then
                        strXY = Split(Mid(varSub(intSubNo)(j), 2), "Y", -1, vbTextCompare)
                        If blnDrillHit = True Then
                            With udtNCInfo
                                .varDrl_Inf(intIndex)(3) = CLng(.varDrl_Inf(intIndex)(3)) + 1
                            End With
                        End If
                        Write #intF1, strXY(X), strXY(Y)
                    ElseIf varSub(intSubNo)(j) Like "G81" = True Then
                        blnDrillHit = True
                        Write #intF1, "G81", ""
                    ElseIf varSub(intSubNo)(j) Like "G80" = True Then
                        blnDrillHit = False
                        Write #intF1, "G80", ""
                    End If
                Next
            End If
        ElseIf strMain(i) Like "T*" = True Then
            intTool = CInt(Mid(strMain(i), 2))
            With udtNCInfo
                For intIndex = 0 To UBound(.varDrl_Inf)
                    If intTool = .varDrl_Inf(intIndex)(0) Then Exit For
                Next
            End With
            Write #intF1, "T" & intTool, ""
        End If
'        objBar.Value = Int(i / lngCount * 100)
    Next
    Close #intF1
    Erase strMain '配列のメモリを開放する
    objBar.Visible = False

End Sub
