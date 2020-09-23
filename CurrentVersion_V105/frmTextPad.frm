VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmTextPad 
   Caption         =   "TextEditor"
   ClientHeight    =   4965
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5925
   Icon            =   "frmTextPad.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   5925
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cmnDlgTE 
      Left            =   3840
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4440
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":08A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":0E3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":0F98
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":10F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":124C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":13A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":1500
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":165A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":2FEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":3586
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":3B20
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":3C7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":4214
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":47AE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Save Text As"
            Object.Tag             =   "save"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Print Text"
            Object.Tag             =   "print"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Copy Selected Text"
            Object.Tag             =   "copy"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Find Text"
            Object.Tag             =   "find"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Select All"
            Object.Tag             =   "selectall"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Zoom In"
            Object.Tag             =   "zoomin"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Zoom Out"
            Object.Tag             =   "zoomout"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   14
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Help"
            Object.Tag             =   "help"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Close TextEditor"
            Object.Tag             =   "close"
            ImageIndex      =   13
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   1
      Top             =   4725
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   423
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1402
            MinWidth        =   1411
            TextSave        =   "5:14 AM"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1535
            MinWidth        =   1411
            TextSave        =   "8/20/2006"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "SCRL"
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox RTB1 
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   450
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   3625
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      TextRTF         =   $"frmTextPad.frx":4D48
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSaveItem 
         Caption         =   "Save As..."
      End
      Begin VB.Menu mnuPrintItem 
         Caption         =   "Print..."
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExitItem 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "&Menu"
      Begin VB.Menu mnuCopyItem 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuFindItem 
         Caption         =   "Find"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuZoom 
         Caption         =   "Zoom   +/-"
         Begin VB.Menu mnuZoomInItem 
            Caption         =   "Zoom In"
         End
         Begin VB.Menu mnuZoomOutItem 
            Caption         =   "Zoom Out"
         End
      End
      Begin VB.Menu mnuSelAllItem 
         Caption         =   "Select All"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAboutItem 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmTextPad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim theSize As Integer

Private Sub Form_Load()

    theSize = 0
    Me.Height = (Screen.Height / 6) * 5
    Me.Width = (Screen.Width / 6) * 5
    Me.Top = (Screen.Height / 2) - (Me.Height / 2)
    Me.Left = (Screen.Width / 2) - (Me.Width / 2)

End Sub

Private Sub Form_Resize()

    RTB1.Top = 50 + Toolbar1.Height
    RTB1.Left = 0
    
    RTB1.Height = Me.Height - Toolbar1.Height - StatusBar1.Height - 750
    RTB1.Width = Me.Width - 120

End Sub

Private Sub mnuAboutItem_Click()

    frmAbout.Show vbModeless, Me

End Sub

Private Sub mnuCopyItem_Click()
    
    If RTB1.SelLength > 0 Then
        Clipboard.Clear
        Clipboard.SetText RTB1.SelText
        RTB1.SetFocus
    End If

End Sub

Private Sub mnuExitItem_Click()

    Unload Me

End Sub

Private Sub mnuPrintItem_Click()

    cmnDlgTE.ShowPrinter

End Sub

Private Sub mnuSaveItem_Click()

    cmnDlgTE.Filter = "Text Document (*.txt)|*.txt|All Files|*.*"
    cmnDlgTE.FileName = "InstalledProgramLog"
    cmnDlgTE.ShowSave
    On Error GoTo erHand
    If Len(cmnDlgTE.FileName) > 0 Then
        RTB1.SaveFile cmnDlgTE.FileName, rtfText
    Else
        Exit Sub
    End If
    
    Exit Sub
    
erHand:
    Exit Sub

End Sub

Private Sub mnuSelAllItem_Click()
    
    RTB1.SelStart = 0
    RTB1.SelLength = Len(RTB1.Text)
    RTB1.SetFocus
    
End Sub

Private Sub mnuZoomInItem_Click()
    
    If theSize < 6 Then
        RTB1.SelStart = 0
        RTB1.SelLength = Len(RTB1.Text)
        RTB1.SelFontSize = RTB1.SelFontSize + 2
        RTB1.SelStart = RTB1.SelStart + Len(RTB1.SelText)
        RTB1.SelLength = 0
        RTB1.SetFocus
        theSize = theSize + 1
    End If
    
End Sub

Private Sub mnuZoomOutItem_Click()
    
    If theSize > 0 Then
        RTB1.SelStart = 0
        RTB1.SelLength = Len(RTB1.Text)
        RTB1.SelFontSize = RTB1.SelFontSize - 2
        RTB1.SelStart = RTB1.SelStart + Len(RTB1.SelText)
        RTB1.SelLength = 0
        RTB1.SetFocus
        theSize = theSize - 1
    End If
        
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Tag
    Case "save"
        mnuSaveItem_Click
    Case "print"
        mnuPrintItem_Click
    Case "copy"
        mnuCopyItem_Click
    Case "find"
        MsgBox "Find"
    Case "selectall"
        mnuSelAllItem_Click
    Case "zoomin"
        mnuZoomInItem_Click
    Case "zoomout"
        mnuZoomOutItem_Click
    Case "help"
        mnuAboutItem_Click
    Case "close"
        mnuExitItem_Click
    End Select

End Sub
