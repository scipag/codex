VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H00C0C0C0&
   Caption         =   "codEX"
   ClientHeight    =   8280
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11985
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar tbrMenu1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgListMenu1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   100
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Open project"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Save project"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Search"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Arrange windows"
            ImageIndex      =   4
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList imgListMenu1 
         Left            =   8160
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":030A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":06A4
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":0A3E
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":0DD8
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.Toolbar tbrMenu2 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgListMenu2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   100
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Open source code view"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Open MetaCode view"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Show controls"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Show variables"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Show strings"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Show comments"
            ImageIndex      =   6
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList imgListMenu2 
         Left            =   8160
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1172
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":150C
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":18A6
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1C40
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1FDA
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2374
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.Toolbar tbrPosition 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   873
      ButtonWidth     =   609
      ButtonHeight    =   714
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      _Version        =   393216
      Begin VB.ComboBox cboPosition 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   120
         Width           =   1815
      End
      Begin MSComctlLib.ProgressBar pbrPointer 
         Height          =   315
         Left            =   2040
         TabIndex        =   1
         Top             =   120
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         MousePointer    =   10
         Scrolling       =   1
      End
   End
   Begin MSComDlg.CommonDialog cdgOpen 
      Left            =   120
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "PHP files|*.php*"
      DialogTitle     =   "Open file"
      FileName        =   "index.php"
      Filter          =   $"frmMain.frx":270E
   End
   Begin MSComDlg.CommonDialog cdgSaveAs 
      Left            =   720
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "ASCII files|*.txt"
      DialogTitle     =   "Save MetaCode"
      FileName        =   "project1-metacode.txt"
      Filter          =   "ASCII files|*.txt|All Files (*.*)|*.*"
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpenItem 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExitItem 
         Caption         =   "&Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuAnalysis 
      Caption         =   "&Analysis"
      Begin VB.Menu mnuAnalysisCreateMetaCodeItem 
         Caption         =   "&Create MetaCode"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuAnalysisSaveMetaCode 
         Caption         =   "&Save MetaCode"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnuWindows 
      Caption         =   "&Windows"
      Begin VB.Menu mnuWindowsResetDesktopItem 
         Caption         =   "&Reset Desktop"
      End
      Begin VB.Menu mnuWindowsSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowsSourceCodeViewItem 
         Caption         =   "Source Code View"
         Shortcut        =   +{F1}
      End
      Begin VB.Menu mnuWindowsMetaCodeViewItem 
         Caption         =   "MetaCode View"
         Shortcut        =   +{F2}
      End
      Begin VB.Menu mnuWindowsSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowsControlsViewItem 
         Caption         =   "Controls View"
         Shortcut        =   +{F3}
      End
      Begin VB.Menu mnuWindowsVariablesViewItem 
         Caption         =   "Variables View"
         Shortcut        =   +{F4}
      End
      Begin VB.Menu mnuWindowsStringsViewItem 
         Caption         =   "Strings View"
         Shortcut        =   +{F5}
      End
      Begin VB.Menu mnuWindowsCommentsViewItem 
         Caption         =   "Comments View"
         Enabled         =   0   'False
         Shortcut        =   +{F6}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAboutItem 
         Caption         =   "&About"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpDemoInformationItem 
         Caption         =   "&Demo Information"
      End
      Begin VB.Menu mnuHelpSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpCodexHomePageItem 
         Caption         =   "codEX home page"
      End
   End
   Begin VB.Menu mnuLog 
      Caption         =   "Log"
      Visible         =   0   'False
      Begin VB.Menu mnuLogClearItem 
         Caption         =   "Clear"
      End
      Begin VB.Menu mnuLogCopyItem 
         Caption         =   "Copy"
         Shortcut        =   ^{INSERT}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboPosition_Click()
    Call SelectPointerInBar(GetAddressPosition(cboPosition.Text))
    Call SelectAddressInSourceCode(cboPosition.Text, False)
End Sub

Private Sub cboPosition_KeyUp(KeyCode As Integer, Shift As Integer)
    Call SelectPointerInBar(GetAddressPosition(cboPosition.Text))
    Call SelectAddressInSourceCode(cboPosition.Text, True)
End Sub

Private Sub MDIForm_Load()
    Me.Caption = APP_NAME
    Me.pbrPointer.Height = cboPosition.Height
    Call ResetWindowOrder
    Call DemoInformation
End Sub

Private Sub MDIForm_Resize()
    If Me.WindowState <> vbMinimized Then
        If Me.Height < 8000 Then
            Me.Height = 8000
        End If
        
        If Me.Width < 10000 Then
            Me.Width = 10000
        End If
    End If
End Sub

Private Sub mnuAnalysisCreateMetaCodeItem_Click()
    Call LockWindows
    Call CreateMetaCode(frmSourceCode.txtSourceCode.Text)
    Call UnlockWindows
End Sub

Private Sub mnuAnalysisSaveMetaCode_Click()
    Dim sFileName As String
    Dim sOutput As String
    Dim lMetaCodeCount As Long
    Dim lCount As Long
    
    cdgSaveAs.InitDir = App.Path
    
    On Error Resume Next
    cdgSaveAs.FileName = "project1-metacode.txt"
    cdgSaveAs.ShowSave
    sFileName = cdgSaveAs.FileName
    
    lMetaCodeCount = frmMetaCode.lsvMetaCode.ListItems.Count
    For lCount = 1 To lMetaCodeCount
        sOutput = sOutput & frmMetaCode.lsvMetaCode.ListItems(lCount).Text & " " & _
            frmMetaCode.lsvMetaCode.ListItems(lCount).ListSubItems(1).Text & " " & _
            frmMetaCode.lsvMetaCode.ListItems(lCount).ListSubItems(2).Text & vbCrLf
    Next lCount
    
    Open sFileName For Output As #1
        Print #1, sOutput
    Close
End Sub

Private Sub mnuFileExitItem_Click()
    Unload Me
End Sub

Private Sub mnuFileOpenItem_Click()
    frmNewAnalysis.Show vbModal, frmMain
End Sub

Public Sub ResetWindowOrder()
    frmSourceCode.WindowState = vbNormal
    frmSourceCode.Top = 0
    frmSourceCode.Left = 0
    frmSourceCode.Height = (frmMain.Height / 9) * 4 - (tbrPosition.Height + tbrMenu1.Height + tbrMenu2.Height)
    frmSourceCode.Width = (frmMain.Width / 5) * 3
    frmSourceCode.Visible = True

    frmMetaCode.WindowState = vbNormal
    frmMetaCode.Top = frmSourceCode.Top + frmSourceCode.Height
    frmMetaCode.Left = frmSourceCode.Left
    frmMetaCode.Height = frmSourceCode.Height
    frmMetaCode.Width = frmSourceCode.Width
    frmMetaCode.Visible = True

    frmLog.Top = frmMetaCode.Top + frmMetaCode.Height
    frmLog.Height = (frmMain.Height / 9) * 1
    frmLog.Left = frmSourceCode.Left
    frmLog.Width = frmMain.Width - 260
    frmLog.Visible = True

    frmControls.WindowState = vbNormal
    frmControls.Top = frmSourceCode.Top
    frmControls.Left = frmSourceCode.Left + frmSourceCode.Width
    frmControls.Height = (frmSourceCode.Height + frmMetaCode.Height) / 4
    frmControls.Width = (frmMain.Width / 5) * 2 - 260
    frmControls.Visible = True
    
    frmVariables.WindowState = vbNormal
    frmVariables.Top = frmControls.Top + frmControls.Height
    frmVariables.Left = frmSourceCode.Left + frmSourceCode.Width
    frmVariables.Height = frmControls.Height
    frmVariables.Width = frmControls.Width
    frmVariables.Visible = True

    frmStrings.WindowState = vbNormal
    frmStrings.Top = frmVariables.Top + frmVariables.Height
    frmStrings.Left = frmVariables.Left
    frmStrings.Height = frmVariables.Height
    frmStrings.Width = frmVariables.Width
    frmStrings.Visible = True

    frmComments.WindowState = vbNormal
    frmComments.Top = frmStrings.Top + frmStrings.Height
    frmComments.Left = frmVariables.Left
    frmComments.Height = frmVariables.Height
    frmComments.Width = frmVariables.Width
    frmComments.Visible = True

    If (frmSourceCode.Visible = True) Then
        frmSourceCode.SetFocus
    End If
End Sub

Private Sub mnuHelpAboutItem_Click()
    frmAbout.Show vbModeless, frmMain
End Sub

Private Sub mnuHelpCodexHomePageItem_Click()
    Call OpenProjectWebsite
End Sub

Private Sub mnuHelpDemoInformationItem_Click()
    Call DemoInformation
End Sub

Private Sub mnuLogClearItem_Click()
    frmLog.lstLog.Clear
End Sub

Private Sub mnuLogCopyItem_Click()
    Dim i As Integer
    Dim iAllItems As Integer
    Dim sItems As String
    
    iAllItems = frmLog.lstLog.ListCount - 1
    
    For i = 1 To iAllItems
        sItems = sItems & frmLog.lstLog.List(i) & vbCrLf
    Next i
    
    Clipboard.SetText sItems, vbCFText
End Sub

Private Sub mnuWindowsCommentsViewItem_Click()
    Call ViewClick(frmComments)
End Sub

Private Sub mnuWindowsControlsViewItem_Click()
    Call ViewClick(frmControls)
End Sub

Private Sub mnuWindowsMetaCodeViewItem_Click()
    Call ViewClick(frmMetaCode)
End Sub

Private Sub mnuWindowsResetDesktopItem_Click()
    Call ResetWindowOrder
End Sub

Public Sub LockWindows()
    frmMain.Enabled = False
    Screen.MousePointer = 11
End Sub

Public Sub UnlockWindows()
    frmMain.Enabled = True
    Screen.MousePointer = vbDefault
End Sub

Private Sub DemoInformation()
    Dim iUsability As String
    Dim iCompiler As Integer
    Dim iControls As Integer
    Dim iVariables As Integer
    Dim iVulnerabilities As Integer
    Dim iPercentage As Integer

    iUsability = 55
    iCompiler = 42
    iControls = 15
    iVariables = 10
    iVulnerabilities = 0

    iPercentage = (iUsability + iCompiler + iControls + iVariables + iVulnerabilities) / 5

    MsgBox "This is an early pre-release of the codEX application. You are not" & vbCrLf & _
        "allowed to copy, re-distribute or sell this edition without permission" & vbCrLf & _
        "of the author. This version contains not all of the final features" & vbCrLf & _
        "(approx. " & iPercentage & " %) and therefore is for testing purposes only:" & vbCrLf & vbCrLf & _
        vbTab & "Usability Features" & vbTab & vbTab & iUsability & " %" & vbCrLf & _
        vbTab & "Compiler for MetaCode" & vbTab & iCompiler & " %" & vbCrLf & _
        vbTab & "Control Path Analysis" & vbTab & iControls & " %" & vbCrLf & _
        vbTab & "Variables Path Analysis" & vbTab & iVariables & " %" & vbCrLf & _
        vbTab & "Vulnerability Identification" & vbTab & iVulnerabilities & " %" & vbCrLf & vbCrLf & _
        "Feel free to send feedback to marc.ruef@computec.ch. Further" & vbCrLf & _
        "details about the application and the project are available at " & vbCrLf & _
        APP_WEBSITE_URL, _
        vbOKOnly + vbInformation, "codEX Pre-Release Information"
End Sub

Private Sub mnuWindowsSourceCodeViewItem_Click()
    Call ViewClick(frmSourceCode)
End Sub

Private Sub ViewClick(ByRef fForm As Form)
    If (fForm.WindowState <> vbNormal) Then
        fForm.WindowState = vbNormal
    End If
    
    If (fForm.Visible = False) Then
        fForm.Visible = True
    End If
    
    fForm.SetFocus
End Sub

Private Sub mnuWindowsStringsViewItem_Click()
    Call ViewClick(frmStrings)
End Sub

Private Sub mnuWindowsVariablesViewItem_Click()
    Call ViewClick(frmVariables)
End Sub

Private Sub pbrPointer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim iIndex As Integer
    Dim iPointer As Integer
    
    If (frmMetaCode.lsvMetaCode.ListItems.Count) Then
        iPointer = (X / pbrPointer.Width) * 100
        iIndex = CInt((frmMetaCode.lsvMetaCode.ListItems.Count / 100) * iPointer) + 1
        
        If (frmMetaCode.lsvMetaCode.ListItems.Count >= iIndex) Then
            Me.pbrPointer.Value = iPointer
            
            If (frmMetaCode.Visible = False) Then
                frmMetaCode.Visible = True
            End If
            
            frmMain.cboPosition.Text = frmMetaCode.lsvMetaCode.ListItems(iIndex).Text
            
            frmMetaCode.lsvMetaCode.ListItems(iIndex).EnsureVisible
            frmMetaCode.lsvMetaCode.SetFocus
        End If
    End If
End Sub

Private Sub pbrPointer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim iPointer As Integer
    
    iPointer = (X / pbrPointer.Width) * 100
    
    pbrPointer.ToolTipText = "Position: " & iPointer & " % of MetaCode"
End Sub

Private Sub tbrMenu1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim iIndex As Integer
    
    iIndex = Button.Index

    If (iIndex = 2) Then
        Call mnuFileOpenItem_Click
    ElseIf (iIndex = 7) Then
        Call mnuWindowsResetDesktopItem_Click
    End If
End Sub

Private Sub tbrMenu2_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim iIndex As Integer
    
    iIndex = Button.Index

    If (iIndex = 2) Then
        Call mnuWindowsSourceCodeViewItem_Click
    ElseIf (iIndex = 3) Then
        Call mnuWindowsMetaCodeViewItem_Click
    'ElseIf (iIndex = 4) Then
    '    Call mnuAnalysisCreateMetaCodeItem_Click
    ElseIf (iIndex = 5) Then
        Call mnuWindowsControlsViewItem_Click
    ElseIf (iIndex = 6) Then
        Call mnuWindowsVariablesViewItem_Click
    ElseIf (iIndex = 7) Then
        Call mnuWindowsStringsViewItem_Click
    End If
End Sub

