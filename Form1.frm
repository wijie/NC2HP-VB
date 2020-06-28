VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000000&
   Caption         =   "Form1"
   ClientHeight    =   2484
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   3540
   ForeColor       =   &H80000007&
   LinkTopic       =   "Form1"
   ScaleHeight     =   2484
   ScaleWidth      =   3540
   StartUpPosition =   3  'Windows ‚ÌŠù’è’l
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   252
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3252
      _ExtentX        =   5736
      _ExtentY        =   445
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   732
      Left            =   1080
      TabIndex        =   0
      Top             =   1440
      Width           =   1332
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    Call sReadConfig
    Call sNCExtract(gudtNCInfo(TH), ProgressBar1)
    If gudtNCInfo(NT).strNCName <> "null" Then
        Call sNCExtract(gudtNCInfo(NT), ProgressBar1)
    End If
    Call sHPConv(gudtNCInfo(TH))
End Sub

