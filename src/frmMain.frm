VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{888CB5C1-D2D1-44D7-A65A-A025AAC95417}#1.0#0"; "wodHttp.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "subdomain bruteforce"
   ClientHeight    =   4140
   ClientLeft      =   5910
   ClientTop       =   4305
   ClientWidth     =   6870
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   6870
   StartUpPosition =   2  'CenterScreen
   Begin wodHttpDLXLibCtl.wodHttpDLX wScan 
      Index           =   0
      Left            =   7320
      OleObjectBlob   =   "frmMain.frx":048A
      Top             =   1920
   End
   Begin VB.Timer tmrTimeout 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   5000
      Left            =   7320
      Top             =   2400
   End
   Begin VB.Timer tmrSocket 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   1
      Left            =   7320
      Top             =   2880
   End
   Begin VB.ListBox lstWordlist 
      Height          =   2790
      Left            =   120
      TabIndex        =   11
      Top             =   960
      Width           =   2895
   End
   Begin VB.TextBox txtResults 
      Height          =   2775
      Left            =   3120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   960
      Width           =   3615
   End
   Begin MSComctlLib.StatusBar statusBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   8
      Top             =   3825
      Width           =   6870
      _ExtentX        =   12118
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5292
            Text            =   "status: idle"
            TextSave        =   "status: idle"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3351
            MinWidth        =   3351
            Text            =   "attempts: 0"
            TextSave        =   "attempts: 0"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3351
            MinWidth        =   3351
            Text            =   "discovered: 0"
            TextSave        =   "discovered: 0"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtTimeout 
      Height          =   315
      Left            =   5160
      TabIndex        =   6
      Text            =   "500"
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox txtPort 
      Height          =   315
      Left            =   4080
      TabIndex        =   4
      Text            =   "80"
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox txtTarget 
      Height          =   315
      Left            =   1080
      TabIndex        =   2
      Text            =   "hackerone.com"
      Top             =   360
      Width           =   2895
   End
   Begin VB.TextBox txtSockets 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Text            =   "25"
      Top             =   360
      Width           =   855
   End
   Begin VB.Label lblDiscovered 
      Caption         =   "0"
      Height          =   255
      Left            =   7440
      TabIndex        =   14
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "sockets"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblAttempts 
      Caption         =   "0"
      Height          =   255
      Left            =   7440
      TabIndex        =   12
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "results.txt"
      Height          =   255
      Left            =   3120
      TabIndex        =   10
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Label6 
      Caption         =   "wordlist.txt"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   3015
   End
   Begin VB.Label Label5 
      Caption         =   "socket timeout (ms)"
      Height          =   255
      Left            =   5160
      TabIndex        =   5
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "port"
      Height          =   255
      Left            =   4080
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "host"
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.Menu start 
      Caption         =   "start"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'welcome 2 subdomain bruteforce for windows 98 because idea/execution > language

Dim sTarget(1 To 500) As String
Dim sTimeout(1 To 500) As Integer

Private Sub Form_Load()

    FlatBorder txtSockets.hwnd
    FlatBorder txtTarget.hwnd
    FlatBorder txtPort.hwnd
    FlatBorder txtTimeout.hwnd
    FlatBorder lstWordlist.hwnd
    FlatBorder txtResults.hwnd

End Sub

Private Sub start_Click()

    start.Enabled = False
    lstWordlist.Enabled = False
    txtTarget.Enabled = False
    txtPort.Enabled = False
    txtTimeout.Enabled = False

    statusBar.Panels(1).Text = "status: loading wordlist"
    Call LoadListFromFile(App.Path & "\wordlist.txt", lstWordlist)
    statusBar.Panels(1).Text = "status: loaded"
    lstWordlist.Visible = True

    For I = 1 To txtSockets.Text
        Load tmrTimeout(I)
        Load tmrSocket(I)
        Load wScan(I)
        tmrTimeout(I).Interval = txtTimeout.Text \ 2
        tmrSocket(I).Enabled = True
    Next

    statusBar.Panels(1).Text = "status: sockets loaded"
    txtSockets.Enabled = False

End Sub

Private Sub tmrSocket_Timer(Index As Integer)

    If Not lstWordlist.ListIndex = lstWordlist.ListCount - 1 Then
        lstWordlist.ListIndex = lstWordlist.ListIndex + 1
        sTarget(Index) = lstWordlist.Text & "." & txtTarget.Text
        lblAttempts.Caption = lblAttempts.Caption + 1
        statusBar.Panels(2).Text = "attempts: " & lblAttempts.Caption
        statusBar.Panels(1).Text = sTarget(Index)
        wScan(Index).Disconnect
        If txtPort.Text = 80 Then
            wScan(Index).URL = "http://" & sTarget(Index)
        Else
            wScan(Index).URL = "https://" & sTarget(Index)
        End If
        wScan(Index).Get
        tmrTimeout(Index).Enabled = True
        tmrSocket(Index).Enabled = False
        Exit Sub
    Else
        tmrSocket(Index).Enabled = False
        Exit Sub
    End If

End Sub

Private Sub tmrTimeout_Timer(Index As Integer)

    If sTimeout(Index) = 2 Then
        sTimeout(Index) = 0
        wScan(Index).Disconnect
        tmrTimeout(Index).Enabled = False
        tmrSocket(Index).Enabled = True
    Else
        sTimeout(Index) = sTimeout(Index) + 1
    End If

End Sub

Private Sub txtResults_Change()

    txtResults.SelStart = Len(txtResults.Text)

End Sub

Private Sub wScan_Done(Index As Integer, ByVal ErrorCode As Long, ByVal ErrorText As String)

    If Len(wScan(Index).Response.Headers.ToString) > 10 Then
        txtResults.Text = txtResults.Text & sTarget(Index) & vbNewLine
        lblDiscovered.Caption = lblDiscovered.Caption + 1
        Call SaveText(txtResults.Text, App.Path & "\results.txt")
        statusBar.Panels(3).Text = "discovered: " & lblDiscovered.Caption
    End If

    sTimeout(Index) = 0
    tmrTimeout(Index).Enabled = False
    tmrSocket(Index).Enabled = True

End Sub

