VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Installed Programs"
   ClientHeight    =   6405
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   8745
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   8745
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      ToolTipText     =   "Exit the program. The loacation of the report will be saved in the registry,"
      Top             =   5640
      Width           =   1815
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "View Report"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      ToolTipText     =   "View report in notepad."
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Caption         =   "Choose the view"
      Height          =   1215
      Left            =   1440
      TabIndex        =   11
      ToolTipText     =   "Choose the view."
      Top             =   4800
      Width           =   1575
      Begin VB.OptionButton optLV 
         Caption         =   "Graphical View"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Will display report with icons."
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton optList 
         Caption         =   "Standard View"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Will display standard-view report."
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin MSComctlLib.ImageList imgLst1 
      Left            =   6720
      Top             =   5670
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":030A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LV1 
      Height          =   4290
      Left            =   45
      TabIndex        =   8
      ToolTipText     =   "Graphic report, with icons."
      Top             =   270
      Visible         =   0   'False
      Width           =   8700
      _ExtentX        =   15346
      _ExtentY        =   7567
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDragMode     =   1
      OLEDropMode     =   1
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   0
   End
   Begin MSComDlg.CommonDialog cmnDlg1 
      Left            =   6750
      Top             =   5730
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox List3 
      Height          =   255
      Left            =   6915
      TabIndex        =   6
      Top             =   5910
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.CommandButton cmdToTxt 
      Caption         =   "Write To External File..."
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      ToolTipText     =   "Will save the list to a text file of your choice. Once such report exists, 'View Report' button will be enabled."
      Top             =   4920
      Width           =   1815
   End
   Begin VB.ListBox List2 
      Height          =   4350
      Left            =   1440
      Sorted          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "Standard view."
      Top             =   240
      Width           =   6015
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   75
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   3360
      TabIndex        =   5
      Top             =   4680
      Width           =   2055
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   6735
      Picture         =   "frmMain.frx":0BE4
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   10
      Top             =   5700
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   555
      Left            =   6705
      ScaleHeight     =   495
      ScaleWidth      =   540
      TabIndex        =   9
      Top             =   5685
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Currently installed programs on your computer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   8655
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Menu"
      Begin VB.Menu mnuRefresh 
         Caption         =   "Total Installed Programs"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWriteDataItem 
         Caption         =   "Write Data To File"
      End
      Begin VB.Menu mnuViewReportItem 
         Caption         =   "View Report"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDeleteReportItem 
         Caption         =   "Delete Report"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete All Settings"
         Enabled         =   0   'False
         Begin VB.Menu mnuConfirm 
            Caption         =   "Confirm (Path of the report will be deleted)"
            Begin VB.Menu mnuNoItem 
               Caption         =   "No"
            End
            Begin VB.Menu mnuDeleteItem 
               Caption         =   "Delete"
            End
         End
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExitItem 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAboutItem 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''Registry APIs''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''Locate System Directory API''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''Run an application API'''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''Extract Icon APIs''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Private Declare Function ExtractAssociatedIcon Lib "shell32.dll" Alias "ExtractAssociatedIconA" (ByVal hInst As Long, ByVal lpIconPath As String, lpiIcon As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''Run Application Constant''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Const SW_SHOWNORMAL = 1

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''Registry Constants''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Const ERROR_NO_MORE_ITEMS = 259&
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const UNINSTALL_KEY = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''Extract Icon Constants''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Const DI_MASK = &H1
Const DI_IMAGE = &H2
Const DI_NORMAL = DI_MASK Or DI_IMAGE

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''Form's Declarations'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Const APP_NAME = "InstalledProgramList_ByZHULEEK"
Const SECTION_NAME = "StartUp"

Dim lHandle As Long
Dim fil As String
Dim intCnt As Integer
Dim intViewStyle As Integer
Dim strReportFileName As String
Dim fso As FileSystemObject


Function openRegKey(KeyName As String) As Long

    openRegKey = RegOpenKey(HKEY_LOCAL_MACHINE, KeyName, lHandle)

End Function

Private Sub closeRegKey(keyHandle As Long)

    RegCloseKey keyHandle

End Sub



Private Sub cmdExit_Click()

    Unload Me

End Sub

Private Sub cmdReport_Click()

    On Error GoTo erHand
    
    Dim isOK As Long
    
    If Len(strReportPath) > 0 Then
        Frame1.Tag = strReportPath
        isOK = ShellExecute(Me.hwnd, vbNullString, strReportPath, vbNullString, vbNullString, SW_SHOWNORMAL)
        If isOK > 32 Then
            Unload frmCommandBar
            frmCommandBar.Show
        Else
            Load frmTextPad
            frmTextPad.RTB1.LoadFile strReportPath
            frmTextPad.Show
        End If
    Else
        Exit Sub
    End If
    
    Exit Sub
    
erHand:

    MsgBox "I was unable to locate a valid path to the report, because it was moved or deleted."
    cmdReport.Enabled = False
    fil = ""
    isReport = False
    saveSettings

End Sub

Private Sub cmdToTxt_Click()

    Dim ff As Long
    
    ff = FreeFile
    
    setNormalAttributes

    On Error GoTo erHand
    cmnDlg1.Filter = "Text File (*.txt)|*.txt"
    cmnDlg1.FileName = "ConfigReport"
    cmnDlg1.Flags = cdlOFNCreatePrompt Or cdlOFNNoReadOnlyReturn Or cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
    cmnDlg1.DefaultExt = ".txt"
    cmnDlg1.ShowOpen
    
    If Len(cmnDlg1.FileName) > 0 Then
        cmdReport.Enabled = True
        mnuViewReportItem.Enabled = True
        mnuDeleteReportItem.Enabled = True
        fil = cmnDlg1.FileName
        If LCase(Right(fil, 4)) <> ".txt" Then fil = fil & ".txt"
        
        Open fil For Output As #ff
        Print #ff, "This list was produced on : " & Date & ", @ : " & Time
        Print #ff, ""
        Print #ff, "The total number of programs listed is : " & List2.ListCount - 1
        Print #ff, ""
        Print #ff, "Installed Programs as of " & Date & " " & Time
        Dim i As Integer
        For i = 0 To List2.ListCount - 1
            Print #ff, List2.List(i)
        Next i
        Print #ff, ""
        Print #ff, "End of list."
        Close #ff
        strReportPath = fil
        intTotalProgs = List2.ListCount - 1
        isReport = True
        extractFileName fil
        saveSettings
        
    Else
        Exit Sub
    End If
    
erHand:
    Exit Sub
    
End Sub


Private Sub Form_Load()

    Dim intCurrentProgs As Integer
    
    getSettings
    
    Show
    
    intCnt = 1
    
    With LV1
        .ColumnHeaders.Add Text:="Icon"
        .ColumnHeaders.Add Text:="Uninstall Path"
        '.ColumnHeaders.Add Text:="Path"
        .LabelEdit = lvwManual
        .View = lvwReport
        
        .ColumnHeaders(1).Width = .Width / 2 '- 50
        .ColumnHeaders(2).Width = .Width  '/ 2 '- 50
        '.ColumnHeaders(3).Width = .Width / 3 - 50
        
    End With
    
    With imgLst1
        .UseMaskColor = True
        .MaskColor = &H8000000A
    End With
    
    
    getAllKeys
    closeRegKey lHandle
    
    intCurrentProgs = List2.ListCount - 1
    
    If isReport Then
        If intCurrentProgs <> intTotalProgs Then
            Dim reply As VbMsgBoxResult
            reply = MsgBox("The number of installed programs on your computer has changed." & vbCrLf & _
                "The saved report is not up to date. Would you like to see the saved report?", vbYesNo, "Attention")
                If reply = vbYes Then
                    If Len(fil) > 0 Then
                        On Error GoTo erHand
                        ShellExecute Me.hwnd, vbNullString, fil, vbNullString, vbNullString, SW_SHOWNORMAL
                    Else
                        GoTo erHand
                    End If
                End If
        End If
    End If
    
    Exit Sub

erHand:
    MsgBox "I was unable to locate a valid path to the report, because it was moved or deleted."
    cmdReport.Enabled = False
    fil = ""
    isReport = False
    saveSettings

End Sub


Sub getAllKeys()

    Dim isOpen As Long
    Dim s As Integer
    
    isOpen = openRegKey(UNINSTALL_KEY)
    
    If isOpen = 0 Then
        
        Dim Ret As Long, cnt As Long, sName As String
        Dim sData As String, retData As Long
        Const BUFFER_SIZE = 255
        
        sName = Space(BUFFER_SIZE)
        Ret = BUFFER_SIZE
        
        While RegEnumKeyEx(lHandle, cnt, sName, Ret, ByVal 0&, vbNullString, ByVal 0&, ByVal 0&) <> ERROR_NO_MORE_ITEMS
            
            Dim tempKey As String
            
            tempKey = Left$(sName, Ret)
            List1.AddItem tempKey
            
            openRegKey (UNINSTALL_KEY & "\" & tempKey)
            
            Dim lRslt As Long, lValueType As Long, strBuf As String
            
            strBuf = String(255, Chr(0))
            
            lRslt = RegQueryValueEx(lHandle, "DisplayName", 0, 0, ByVal strBuf, BUFFER_SIZE)
            
            If lRslt = 0 Then
                strBuf = Left(strBuf, 255)
                
                Dim tp As String
                
                List3.Clear
                List3.AddItem strBuf
                tp = List3.List(0)
                If List2.ListCount = 0 Then
                    List2.AddItem strBuf
                Else
                    Dim addIt As Boolean
                    addIt = True
                    For s = 0 To List2.ListCount - 1
                        If Left(List2.List(s), 200) = Left(tp, 200) Then
                            addIt = False
                        End If
                    Next s
                    If addIt = True Then
                        List2.AddItem tp 'strBuf
                        
                        
                        
                        Dim uninRslt As Long
                        Dim strUninPath As String
                        
                        strUninPath = Space(BUFFER_SIZE)
                        
                        uninRslt = RegQueryValueEx(lHandle, "UninstallString", 0, 0, ByVal strUninPath, BUFFER_SIZE)
                        
                        If uninRslt = 0 Then
                            List3.Clear
                            List3.AddItem strUninPath
                            strUninPath = ""
                            strUninPath = List3.List(0)
                            List3.Clear
                        Else
                            strUninPath = "Not Available"
                        End If
                        
                        
                                              
                        Dim icnRslt As Long
                        Dim strIconPath As String
                        
                        strIconPath = Space(BUFFER_SIZE)
                        
                        icnRslt = RegQueryValueEx(lHandle, "DisplayIcon", 0, 0, ByVal strIconPath, BUFFER_SIZE)
                        If icnRslt = 0 Then
                            intCnt = intCnt + 1
                            Picture1.Cls
                            List3.Clear
                            List3.AddItem strIconPath
                            strIconPath = ""
                            strIconPath = List3.List(0)
                            List3.Clear
                            
                            Dim mIcon As Long
                            
                            mIcon = ExtractAssociatedIcon(App.hInstance, strIconPath, 1)
                            
                            DrawIconEx Picture1.hdc, 0, 0, mIcon, 0, 0, 0, 0, DI_NORMAL
                            DestroyIcon mIcon
                            
                            With imgLst1
                                .ListImages.Add Picture:=Picture1.Image
                            End With

                            With LV1
                                .SmallIcons = imgLst1
                                With .ListItems.Add(, , tp, , intCnt) 'SmallIcon:=1)
                                    .ListSubItems.Add Text:=strUninPath
'                                    .ListSubItems.Add Text:=strIconPath
                                End With
                            End With
                            
                        Else
                            
                            With LV1
                                .SmallIcons = imgLst1
                                With .ListItems.Add(, , tp, SmallIcon:=1)
                                    .ListSubItems.Add Text:=strUninPath
'                                    .ListSubItems.Add Text:=""
                                End With
                            End With
                        End If
                    End If
                End If
            Else
                Debug.Print Left(strBuf, Ret)
            End If
            
            cnt = cnt + 1
            sName = Space(BUFFER_SIZE)
            Ret = BUFFER_SIZE
            closeRegKey (lHandle)
            openRegKey (UNINSTALL_KEY)
        Wend

    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    saveSettings
    setROAttributes
    End

End Sub

Private Sub List2_DblClick()

'    For h = 1 To Len(List2.Text)
'        If Left(List2.Text, h) = " " Then
'            MsgBox "Y"
'        Else
'            MsgBox Left(List2.Text, h)
'        End If
'    Next h

End Sub

Private Sub mnuAboutItem_Click()

    frmAbout.Show vbModeless, Me

End Sub

Private Sub mnuDeleteItem_Click()

    DeleteSetting APP_NAME

End Sub

Public Sub mnuDeleteReportItem_Click()

    If isReport Then
        
        On Error GoTo erHand
        
        Dim fso As FileSystemObject
        Dim fso_Fil As File
    
        Set fso = New FileSystemObject
        Set fso_Fil = fso.GetFile(strReportPath)
        fso_Fil.Delete True
        isReport = False
        strReportPath = ""
        intTotalProgs = -1
        fil = ""
        cmdReport.Enabled = False
        mnuDeleteReportItem.Enabled = False
        mnuViewReportItem.Enabled = False
        
    End If
    
    Exit Sub

erHand:
    
    MsgBox "Error occured. Error number : " & Err.Number, vbInformation, Err.Description
    Exit Sub
    
End Sub

Private Sub mnuExitItem_Click()

    Unload Me
    
End Sub

Private Sub mnuNoItem_Click()

    Exit Sub

End Sub

Private Sub mnuRefresh_Click()

 MsgBox "Total of programs currently installed" & vbCrLf & Space(11) & _
        "on your system is : " & List2.ListCount - 1

End Sub

Private Sub mnuViewReportItem_Click()

    cmdReport_Click

End Sub

Private Sub mnuWriteDataItem_Click()

    cmdToTxt_Click

End Sub

Private Sub optList_Click()

    LV1.Visible = Not LV1.Visible
    List2.Visible = Not List2.Visible

End Sub

Private Sub optLV_Click()

    LV1.Visible = Not LV1.Visible
    List2.Visible = Not List2.Visible

End Sub

Sub getSettings()

    isReport = CBool(GetSetting(APP_NAME, SECTION_NAME, "ReportExists", "False"))
    intViewStyle = CInt(GetSetting(APP_NAME, SECTION_NAME, "ReportView", "1"))
    Me.Top = CInt(GetSetting(APP_NAME, SECTION_NAME, "FormTop", (Screen.Height / 2) - (Me.Height / 2)))
    Me.Left = CInt(GetSetting(APP_NAME, SECTION_NAME, "FormLeft", (Screen.Width / 2) - (Me.Width / 2)))
    
    If intViewStyle = 1 Then
        optList.Value = True
        optLV.Value = False
        List2.Visible = True
        LV1.Visible = False
    Else
        optList.Value = False
        optLV.Value = True
        List2.Visible = False
        LV1.Visible = True
    End If
    
    If isReport Then
        strReportPath = GetSetting(APP_NAME, SECTION_NAME, "ReportLocation", "")
        intTotalProgs = CInt(GetSetting(APP_NAME, SECTION_NAME, "TotalProgramsInstalled", -1))
        cmdReport.Enabled = True
        fil = strReportPath
        mnuDelete.Enabled = True
        extractFileName strReportPath
        mnuDeleteReportItem.Enabled = True
        mnuViewReportItem.Enabled = True
        setNormalAttributes
    Else
        mnuDelete.Enabled = False
        cmdReport.Enabled = False
        mnuDeleteReportItem.Enabled = False
        mnuViewReportItem.Enabled = False
    End If

End Sub

Sub saveSettings()

    If optList.Value = True Then
        intViewStyle = 1
    Else
        intViewStyle = 2
    End If

    SaveSetting APP_NAME, SECTION_NAME, "ReportExists", CStr(isReport)
    SaveSetting APP_NAME, SECTION_NAME, "ReportLocation", strReportPath
    SaveSetting APP_NAME, SECTION_NAME, "TotalProgramsInstalled", intTotalProgs
    SaveSetting APP_NAME, SECTION_NAME, "ReportView", intViewStyle
    SaveSetting APP_NAME, SECTION_NAME, "FormTop", Me.Top
    SaveSetting APP_NAME, SECTION_NAME, "FormLeft", Me.Left

End Sub

Sub extractFileName(ByVal pth As String)

    Dim r As Integer

    r = InStrRev(pth, "\", Len(pth), vbTextCompare)
    pth = Right(pth, Len(pth) - r)
    cmdExit.Tag = pth

End Sub

Sub setNormalAttributes()

    Set fso = New FileSystemObject
        Dim theFil As File
        Dim isTheFil As Boolean
        isTheFil = fso.FileExists(strReportPath)
        
    If isTheFil Then
        Set theFil = fso.GetFile(strReportPath)
        theFil.Attributes = Normal
    End If
    
    Set fso = Nothing

End Sub


Sub setROAttributes()

    Set fso = New FileSystemObject
        Dim theFil As File
        Dim isTheFil As Boolean
        isTheFil = fso.FileExists(strReportPath)
        
    If isTheFil Then
        Set theFil = fso.GetFile(strReportPath)
        theFil.Attributes = ReadOnly
    End If

End Sub
