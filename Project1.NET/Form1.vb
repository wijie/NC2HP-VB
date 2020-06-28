Option Strict Off
Option Explicit On
Friend Class Form1
	Inherits System.Windows.Forms.Form
#Region "Windows フォーム デザイナによって生成されたコード"
	Public Sub New()
		MyBase.New()
		If m_vb6FormDefInstance Is Nothing Then
			If m_InitializingDefInstance Then
				m_vb6FormDefInstance = Me
			Else
				Try 
					'スタートアップ フォームについては、最初に作成されたインスタンスが既定インスタンスになります。
					If System.Reflection.Assembly.GetExecutingAssembly.EntryPoint.DeclaringType Is Me.GetType Then
						m_vb6FormDefInstance = Me
					End If
				Catch
				End Try
			End If
		End If
		'この呼び出しは、Windows フォーム デザイナで必要です。
		InitializeComponent()
	End Sub
	'Form は、コンポーネント一覧に後処理を実行するために dispose をオーバーライドします。
	Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Windows フォーム デザイナで必要です。
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents ProgressBar1 As AxMSComctlLib.AxProgressBar
	Public WithEvents Command1 As System.Windows.Forms.Button
	'メモ : 以下のプロシージャは Windows フォーム デザイナで必要です。
	'Windows フォーム デザイナを使って変更できます。
	'コード エディタを使って修正しないでください。
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Form1))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ToolTip1.Active = True
		Me.ProgressBar1 = New AxMSComctlLib.AxProgressBar
		Me.Command1 = New System.Windows.Forms.Button
		CType(Me.ProgressBar1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.BackColor = System.Drawing.SystemColors.ScrollBar
		Me.Text = "Form1"
		Me.ClientSize = New System.Drawing.Size(236, 166)
		Me.Location = New System.Drawing.Point(4, 19)
		Me.ForeColor = System.Drawing.SystemColors.MenuText
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 12)
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.MaximizeBox = True
		Me.MinimizeBox = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "Form1"
		ProgressBar1.OcxState = CType(resources.GetObject("ProgressBar1.OcxState"), System.Windows.Forms.AxHost.State)
		Me.ProgressBar1.Size = New System.Drawing.Size(217, 17)
		Me.ProgressBar1.Location = New System.Drawing.Point(8, 24)
		Me.ProgressBar1.TabIndex = 1
		Me.ProgressBar1.Name = "ProgressBar1"
		Me.Command1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Command1.Text = "Command1"
		Me.Command1.Size = New System.Drawing.Size(89, 49)
		Me.Command1.Location = New System.Drawing.Point(72, 96)
		Me.Command1.TabIndex = 0
		Me.Command1.BackColor = System.Drawing.SystemColors.Control
		Me.Command1.CausesValidation = True
		Me.Command1.Enabled = True
		Me.Command1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Command1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Command1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Command1.TabStop = True
		Me.Command1.Name = "Command1"
		Me.Controls.Add(ProgressBar1)
		Me.Controls.Add(Command1)
		CType(Me.ProgressBar1, System.ComponentModel.ISupportInitialize).EndInit()
	End Sub
#End Region 
#Region "アップグレード ウィザードのサポート コード"
	Private Shared m_vb6FormDefInstance As Form1
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As Form1
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New Form1()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region 
	
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		
		Call sReadConfig()
		Call sNCExtract(gudtNCInfo(TH), ProgressBar1)
		If gudtNCInfo(NT).strNCName <> "null" Then
			Call sNCExtract(gudtNCInfo(NT), ProgressBar1)
		End If
		Call sHPConv(gudtNCInfo(TH))
	End Sub
End Class