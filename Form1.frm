VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   Caption         =   "Unknown Browser - Opening home page.."
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10650
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   7980
   ScaleWidth      =   10650
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   300
      Left            =   6840
      TabIndex        =   11
      Top             =   7680
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   529
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   9360
      Top             =   1080
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   4920
      Picture         =   "Form1.frx":0874
      ScaleHeight     =   855
      ScaleWidth      =   5970
      TabIndex        =   10
      Top             =   120
      Width           =   5970
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1635
      Left            =   0
      Picture         =   "Form1.frx":3530
      ScaleHeight     =   1635
      ScaleWidth      =   30000
      TabIndex        =   1
      Top             =   0
      Width           =   30000
      Begin VB.CommandButton Command1 
         Caption         =   "GO"
         Default         =   -1  'True
         Height          =   280
         Left            =   6720
         TabIndex        =   9
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   0
         TabIndex        =   8
         Top             =   1320
         Width           =   6495
      End
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   5
         Left            =   0
         Picture         =   "Form1.frx":7D6D
         ScaleHeight     =   210
         ScaleWidth      =   1020
         TabIndex        =   7
         Top             =   1080
         Width           =   1020
      End
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   1
         Left            =   960
         Picture         =   "Form1.frx":828C
         ScaleHeight     =   975
         ScaleWidth      =   945
         TabIndex        =   6
         Top             =   0
         Width           =   945
      End
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   945
         Index           =   2
         Left            =   2040
         Picture         =   "Form1.frx":8B9A
         ScaleHeight     =   945
         ScaleWidth      =   915
         TabIndex        =   5
         Top             =   0
         Width           =   915
      End
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   945
         Index           =   3
         Left            =   3000
         Picture         =   "Form1.frx":940E
         ScaleHeight     =   945
         ScaleWidth      =   930
         TabIndex        =   4
         Top             =   0
         Width           =   930
      End
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   945
         Index           =   4
         Left            =   3960
         Picture         =   "Form1.frx":C294
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   3
         Top             =   0
         Width           =   945
      End
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   0
         Left            =   0
         Picture         =   "Form1.frx":CBEA
         ScaleHeight     =   975
         ScaleWidth      =   960
         TabIndex        =   2
         Top             =   0
         Width           =   960
      End
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   1635
      Width           =   10455
      ExtentX         =   18441
      ExtentY         =   10186
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   12
      Top             =   7650
      Width           =   10650
      _ExtentX        =   18785
      _ExtentY        =   582
      SimpleText      =   "test"
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   18256
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command1_Click()
On Error Resume Next
WebBrowser1.Navigate Text1.Text
End Sub



Private Sub Form_Load()
On Error Resume Next
WebBrowser1.GoHome
End Sub

Private Sub Form_Resize()
On Error Resume Next
'StatusBar1.Panels(1).Width = Form1.Width
WebBrowser1.Width = Form1.Width - 100
WebBrowser1.Height = Form1.Height - 2450
Text1.Width = Form1.Width - 1000
Picture1.Width = Form1.Width
Picture3.Left = Form1.Width - Picture3.Width - 200
ProgressBar1.Left = Form1.Width - ProgressBar1.Width - 200 - 200
ProgressBar1.Top = Form1.Height - ProgressBar1.Height - 200 - 310
Dim INTwh As Integer
INTwh = Form1.Width - 120
INTwh = INTwh - Command1.Width
Command1.Left = INTwh
End Sub



Private Sub Picture2_Click(Index As Integer)
On Error Resume Next
If Index = 0 Then
WebBrowser1.GoBack
End If
If Index = 1 Then
WebBrowser1.GoForward
End If
If Index = 2 Then
WebBrowser1.GoHome
End If
If Index = 3 Then
WebBrowser1.Stop
End If
If Index = 4 Then
WebBrowser1.Refresh
End If


End Sub

Private Sub WebBrowser1_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
Text1.Text = URL
End Sub

Private Sub WebBrowser1_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
On Error Resume Next
ProgressBar1.Max = ProgressMax
ProgressBar1.Value = Progress
End Sub

Private Sub WebBrowser1_StatusTextChange(ByVal Text As String)
On Error Resume Next
StatusBar1.Panels(1).Text = Text
End Sub

Private Sub WebBrowser1_TitleChange(ByVal Text As String)
On Error Resume Next
Form1.Caption = "Unknown Browser - " & Text
End Sub
