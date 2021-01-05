VERSION 5.00
Begin VB.Form Progress 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4860
   FillStyle       =   5  'Downward Diagonal
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   99
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   324
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   3278
      TabIndex        =   3
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   128
      TabIndex        =   2
      Top             =   720
      Width           =   1455
   End
   Begin VB.PictureBox picProgress 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   128
      ScaleHeight     =   390
      ScaleWidth      =   4545
      TabIndex        =   0
      Top             =   120
      Width           =   4605
      Begin VB.PictureBox picProgressSlide 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   0
         ScaleHeight     =   345
         ScaleWidth      =   2040
         TabIndex        =   1
         Top             =   15
         Width           =   2040
      End
   End
   Begin VB.Timer tmrUpdate 
      Interval        =   100
      Left            =   3840
      Top             =   -1680
   End
End
Attribute VB_Name = "Progress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim Counter As Integer
Public Sub StartAndStopProgress()
    With picProgress
        .Cls
        .BackColor = vbHighlightText
        .ForeColor = vbHighlight
        .Visible = False
    End With
    With picProgressSlide
        .Cls
        .BackColor = vbHighlight
        .ForeColor = vbHighlightText
        .Move 0, 0, 1, picProgress.ScaleHeight
        .Visible = False
    End With
End Sub

Public Sub UpdateProgress(ByVal Percentage As Single)

Dim lTextTop As Long, lTextLeft As Long
    
    picProgress.Visible = True
    picProgressSlide.Visible = True
    
    picProgress.Cls
    picProgressSlide.Cls
    
    picProgressSlide.Width = picProgress.ScaleWidth * Percentage
    
    lTextTop = (picProgress.ScaleHeight - picProgress.TextHeight(Percentage * 100 & " %")) / 2
    lTextLeft = (picProgress.ScaleWidth - picProgress.TextWidth(Percentage * 100 & " %")) / 2
    picProgress.CurrentX = lTextLeft
    picProgress.CurrentY = lTextTop
    picProgressSlide.CurrentX = lTextLeft
    picProgressSlide.CurrentY = lTextTop
    picProgress.Print Percentage * 100 & " %"
    picProgressSlide.Print Percentage * 100 & " %"
    picProgress.Refresh
    picProgressSlide.Refresh
End Sub


Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            tmrUpdate.Enabled = True
        Case 1
            End
    End Select
End Sub

Private Sub Form_Load()
    StartAndStopProgress
    tmrUpdate.Enabled = False
    Me.Show
End Sub


Private Sub tmrUpdate_Timer()
    Randomize
    tmrUpdate.Interval = Rnd * 300 + 1
    Counter = Counter + 1
    If Counter = 101 Then
        StartAndStopProgress
        tmrUpdate.Enabled = False
        Counter = 0
        Exit Sub
    End If
    UpdateProgress (Counter / 100)
End Sub
