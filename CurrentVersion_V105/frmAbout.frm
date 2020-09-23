VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About "
   ClientHeight    =   3600
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5700
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2484.784
   ScaleMode       =   0  'User
   ScaleWidth      =   5352.595
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5160
      Top             =   2880
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5160
      Top             =   2880
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   5160
      Top             =   2880
   End
   Begin VB.CommandButton Command1 
      DownPicture     =   "frmAbout.frx":000C
      Height          =   390
      Left            =   4260
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmAbout.frx":1366
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2895
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   2100
      Picture         =   "frmAbout.frx":26C0
      ScaleHeight     =   255
      ScaleWidth      =   1335
      TabIndex        =   1
      Top             =   1065
      Width           =   1395
      Begin VB.Label Label2 
         Caption         =   "huleek Co."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   255
         TabIndex        =   2
         Top             =   0
         Width           =   1290
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000B&
      Caption         =   $"frmAbout.frx":3502
      Height          =   810
      Left            =   90
      TabIndex        =   3
      Top             =   2715
      Width           =   3615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000000FF&
      FillStyle       =   7  'Diagonal Cross
      Height          =   1665
      Left            =   375
      Top             =   375
      Width           =   4905
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00FFFFFF&
      X1              =   225.372
      X2              =   5070.88
      Y1              =   1822.175
      Y2              =   1822.175
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   225.372
      X2              =   5070.88
      Y1              =   1822.175
      Y2              =   1822.175
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000005&
      X1              =   4958.193
      X2              =   4958.193
      Y1              =   248.478
      Y2              =   1408.044
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000005&
      X1              =   4958.193
      X2              =   338.059
      Y1              =   1408.044
      Y2              =   1408.044
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00808080&
      X1              =   338.059
      X2              =   338.059
      Y1              =   1408.044
      Y2              =   248.478
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00808080&
      X1              =   338.059
      X2              =   4958.193
      Y1              =   248.478
      Y2              =   248.478
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   5070.88
      X2              =   5070.88
      Y1              =   165.652
      Y2              =   1490.87
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      X1              =   225.372
      X2              =   5070.88
      Y1              =   1490.87
      Y2              =   1490.87
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   5070.88
      X2              =   225.372
      Y1              =   165.652
      Y2              =   165.652
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   225.372
      X2              =   225.372
      Y1              =   165.652
      Y2              =   1490.87
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intDirec As Integer
Dim tempA As Integer, tempB As Integer

Private Sub Command1_Click()

    Unload Me

End Sub

Private Sub Form_Click()

    onClick
    
End Sub

Private Sub Form_Load()

    Me.Top = (Screen.Height / 2) - (Me.Height / 2)
    Me.Left = (Screen.Width / 2) - (Me.Width / 2)
    
    intDirec = 10

End Sub

Private Sub Label2_Click()

    onClick
    
End Sub

Private Sub Picture2_Click()

    onClick

End Sub

Private Sub Timer1_Timer()

    Picture2.Left = Picture2.Left + intDirec
    If Picture2.Left >= Shape1.Width - Picture2.Width + 300 Then intDirec = -10
    If Picture2.Left <= Shape1.Left Then intDirec = 10

End Sub

Private Sub onClick()

    'Timer1.Enabled = Not Timer1.Enabled
    Timer2.Enabled = Not Timer2.Enabled
    If Timer2.Enabled = False Then
        Timer3.Enabled = True
    Else
        Timer3.Enabled = False
    End If

End Sub

Private Sub Timer2_Timer()

    On Error GoTo erHand

    Picture2.Width = Picture2.Width - 10
    Picture2.Height = Picture2.Height - 1.5
    If Picture2.Width <= 0 Then Picture2.Visible = False
    If Picture2.Height <= 0 Then Picture2.Visible = False
    Exit Sub
    
erHand:
    Picture2.Visible = False
    Exit Sub
    

End Sub

Private Sub Timer3_Timer()

    Picture2.Visible = True
    Picture2.Width = Picture2.Width + 10
    Picture2.Height = Picture2.Height + 1.5
    If Picture2.Width >= 1309.977 Then Picture2.Width = 1309.977: tempA = 10
    If Picture2.Height >= 217.419 Then Picture2.Height = 217.419: tempB = 10
    If tempA = 10 And tempB = 10 Then
        Timer3.Enabled = False
        tempA = 0
        tempB = 0
    End If

End Sub
