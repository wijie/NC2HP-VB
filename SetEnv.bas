Attribute VB_Name = "SetEnv"
Option Explicit

Sub sReadConfig()

    Dim intFNo0 As Integer
    Dim strValue() As String
    Dim strTmpArray1() As String
    Dim strTmpArray2() As String
    Dim bytBuf() As Byte
    Dim strConfig As String
    Dim i As Integer
    Dim j As Integer

    'DEFファイルを読み込む
    intFNo0 = FreeFile
    Open fTempPath & "NC2HPGL.TBL" For Binary As #intFNo0
    ReDim bytBuf(LOF(intFNo0))
    Get #intFNo0, , bytBuf
    Close #intFNo0
    strConfig = StrConv(bytBuf, vbUnicode)

    strValue = Split(strConfig, vbCrLf, -1, vbTextCompare)
    With gudtNCInfo(TH)
        .strNCName = strValue(1) 'THのファイル名
        strTmpArray1 = Split(strValue(2), " ", -1, vbTextCompare)
        ReDim .varDrl_Inf(UBound(strTmpArray1))
        For j = 0 To UBound(strTmpArray1)
            strTmpArray2 = Split(strTmpArray1(j), ":", -1, vbTextCompare)
            .varDrl_Inf(j) = Array(CInt(Mid(strTmpArray2(0), 3)), _
                                   strTmpArray2(1), _
                                   strTmpArray2(2), _
                                   0) '最後のは穴数
        Next
        .strWB_Inf = Split(strValue(3), ":", -1, vbTextCompare)
        .strSosu = strValue(4) '"Dual" or "Multi"
    End With
    With gudtNCInfo(NT)
        .strNCName = strValue(5) 'NTのファイル名
        If .strNCName <> "null" Then
            strTmpArray1 = Split(strValue(6), " ", -1, vbTextCompare)
            For j = 0 To UBound(strTmpArray1)
                strTmpArray2 = Split(strTmpArray1(j), ":", -1, vbTextCompare)
                .varDrl_Inf(j) = Array(CInt(Mid(strTmpArray2(0), 3)), _
                                      strTmpArray2(1), _
                                      strTmpArray2(2), _
                                      0) '最後のは穴数
            Next
            .strWB_Inf = gudtNCInfo(TH).strWB_Inf
            .strSosu = gudtNCInfo(TH).strSosu '"Dual" or "Multi"
        End If
    End With

End Sub

