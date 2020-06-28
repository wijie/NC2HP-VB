Attribute VB_Name = "Module1"
Option Explicit

Public Const X As Integer = 0
Public Const Y As Integer = 1
Public Const R As Integer = 2
Public gstrNC() As String
Private strEnter As String

Sub Main()

    If Command = "" Then
        MsgBox "�t�@�C�������w�肵�Ă�������"
        End
    End If

    Load Form1
    Form1.Show

End Sub

Public Sub sGetNC()

    Dim intF0 As Integer
    Dim strNC As String
    Dim bytBuf() As Byte

    'NC��ǂݍ���
    intF0 = FreeFile
    Open Command For Binary As #intF0
    ReDim bytBuf(LOF(intF0))
    Get #intF0, , bytBuf
    Close #intF0
    strNC = StrConv(bytBuf, vbUnicode)

    '���s�R�[�h�𒲂ׂ�
    If InStr(strNC, vbCrLf) > 0 Then
        strEnter = vbCrLf
    ElseIf InStr(strNC, vbLf) > 0 Then
        strEnter = vbLf
    ElseIf InStr(strNC, vbCr) > 0 Then
        strEnter = vbCr
    End If

    '�T�u��������W�J����
    Call NCExtract(strNC)

    '���������J������
    Erase bytBuf
    strNC = ""
End Sub

Sub NCExtract(ByVal strNC As String)

    Dim strMainSub() As String
    Dim strSub(44 To 97) As String
    Dim strSubTmp() As String
    Dim intN As Integer
    Dim intSubList() As Integer
    Dim varSubNo As Variant
    Dim strTmp As String
    Dim intFNo0 As Integer
    Dim varDelStr() As Variant
    Dim varStr As Variant
    Dim i As Integer

    strMainSub = Split(strNC, "G25" & strEnter, -1, vbTextCompare)
    If UBound(strMainSub) = 1 Then
        strSubTmp = Split(strMainSub(0), "N", -1, vbTextCompare)
        ReDim intSubList(UBound(strSubTmp) - 1)
        For i = 1 To UBound(strSubTmp)
            intN = Left(strSubTmp(i), 2) '�T�u�������̔ԍ����擾
            strSub(intN) = Mid(strSubTmp(i), 3) '���W�f�[�^���擾
            intSubList(i - 1) = intN
        Next
        '�T�u��������W�J����
        For Each varSubNo In intSubList
            strTmp = Replace(strMainSub(1), "M" & varSubNo, strSub(varSubNo), 1, -1)
            strMainSub(1) = strTmp
        Next
    End If

    '�폜���镶�������������
    varDelStr = Array("G26", "M00", "M02", "M99", "%", " ") '�폜���镶����
    For Each varStr In varDelStr
        strMainSub(1) = Replace(strMainSub(1), varStr, "")
    Next
    While InStr(strMainSub(1), strEnter & strEnter) <> 0
        strMainSub(1) = Replace(strMainSub(1), strEnter & strEnter, strEnter)
    Wend

    gstrNC = Split(strMainSub(1), strEnter, -1, vbTextCompare)
    Call sHPConv(gstrNC)

    '�o�͂���
    intFNo0 = FreeFile
    Open "NC2HP-VB.DAT" For Output As #intFNo0
    Print #intFNo0, strMainSub(1)
    Close #intFNo0

    '���������J������
    Erase strSubTmp
    Erase intSubList
    Erase strMainSub
    strTmp = ""
    End
End Sub
