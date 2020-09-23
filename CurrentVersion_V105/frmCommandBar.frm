VERSION 5.00
Begin VB.Form frmCommandBar 
   BackColor       =   &H004080FF&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4200
   LinkTopic       =   "Form2"
   MousePointer    =   15  'Size All
   Picture         =   "frmCommandBar.frx":0000
   ScaleHeight     =   435
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete Report"
      Height          =   300
      Left            =   2835
      MousePointer    =   1  'Arrow
      TabIndex        =   2
      Top             =   75
      Width           =   1215
   End
   Begin VB.CommandButton cmdUnload 
      Caption         =   "Close Toolbar"
      Height          =   300
      Left            =   150
      MousePointer    =   1  'Arrow
      TabIndex        =   1
      Top             =   75
      Width           =   1215
   End
   Begin VB.CommandButton cmdCloseNotepad 
      Caption         =   "Close Notepad"
      Height          =   300
      Left            =   1485
      MousePointer    =   1  'Arrow
      TabIndex        =   0
      Top             =   75
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   2760
      Top             =   1920
   End
End
Attribute VB_Name = "frmCommandBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
  (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Declare Function BringWindowToTop Lib "user32" _
  (ByVal hwnd As Long) As Long

Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Const WM_CLOSE = &H10

Const GWL_STYLE = (-16)

Const LWA_COLORKEY = &H1
Const GWL_EXSTYLE = (-20)
Const WS_EX_LAYERED = &H80000
Const BM_SETSTATE = &HF3

Const LWA_ALPHA = &H2
Const LWA_BOTH = &H3

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40

Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Dim topMost As Boolean
Dim rslt As Boolean
Dim lHandle As Long
Dim theX As Integer
Dim theY As Integer

Sub setTrans()

    Dim col As Long
    Dim Ret As Long
    Dim attrib As Long
    Dim intTrans As Integer
    Dim flag As Byte

    intTrans = 80
    col = RGB(255, 128, 64)
    flag = 0
    flag = flag Or LWA_COLORKEY
    flag = flag Or LWA_ALPHA
    attrib = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    SetWindowLong Me.hwnd, GWL_EXSTYLE, attrib Or WS_EX_LAYERED
    SetLayeredWindowAttributes Me.hwnd, col, intTrans, flag
    
End Sub

Public Sub setTopMost()
    If topMost Then
        SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    Else
        SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    End If
End Sub

Private Sub cmdCloseNotepad_Click()

    Dim theRlt As Long
    
    theRlt = PostMessage(lHandle, WM_CLOSE, &O0, &O0)

End Sub

Private Sub cmdDelete_Click()

    frmMain.mnuDeleteReportItem_Click
    Exit Sub

'    Dim fso As FileSystemObject
'    Dim fso_Fil As File
'
'    On Error GoTo erHand
'
'    Set fso = New FileSystemObject
'    Set fso_Fil = fso.GetFile(frmMain.Frame1.Tag)
'    fso_Fil.Delete True
'    isReport = False
'    strReportPath = ""
'    intTotalProgs = -1
'    fil = ""
'    frmMain.cmdReport.Enabled = False
'    frmMain.mnuDeleteReportItem.Enabled = False
'    frmMain.mnuViewReportItem.Enabled = False
'
'    Exit Sub
'
'erHand:
'
'    MsgBox "Error finding the file. Error number : " & Err.Number, vbInformation, Err.Description
    
'    Exit Sub
End Sub

Private Sub cmdUnload_Click()

    Me.Hide

End Sub

Private Sub Form_Load()
    
    topMost = True
    setTopMost
    setTrans
    
    Me.Top = (Screen.Height / 2) - (Me.Height / 2)
    Me.Left = (Screen.Width / 2) - (Me.Width / 2)
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
        Me.Tag = "move"
        theX = X
        theY = Y
    Else
        Me.Tag = ""
    End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Me.Tag = "move" Then
        Me.Top = Me.Top + Y - theY
        Me.Left = Me.Left + X - theX
    End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Me.Tag = ""

End Sub

Private Sub Form_Unload(Cancel As Integer)

    topMost = False
    setTopMost

End Sub

Private Sub Timer1_Timer()

    lHandle = FindWindow("notepad", frmMain.cmdExit.Tag & " - Notepad")
    If lHandle = 0 Then Timer1.Interval = 500: Me.Hide
    Timer1.Interval = 25

End Sub
