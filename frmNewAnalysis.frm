VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmNewAnalysis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New source analysis"
   ClientHeight    =   4665
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   6960
   Icon            =   "frmNewAnalysis.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog cdgOpen 
      Left            =   720
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Select Source File to analyze"
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   120
      Top             =   4200
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewAnalysis.frx":038A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewAnalysis.frx":1EDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewAnalysis.frx":3A2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewAnalysis.frx":5580
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewAnalysis.frx":70D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewAnalysis.frx":8C24
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewAnalysis.frx":A776
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewAnalysis.frx":C2C8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lsvSources 
      Height          =   3612
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   6732
      _ExtentX        =   11880
      _ExtentY        =   6376
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      OLEDragMode     =   1
      _Version        =   393217
      Icons           =   "imlIcons"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      OLEDragMode     =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   372
      Left            =   3600
      TabIndex        =   1
      Top             =   4200
      Width           =   972
   End
   Begin VB.CommandButton cmdOkay 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   372
      Left            =   2280
      TabIndex        =   0
      Top             =   4200
      Width           =   1092
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "select the source code file format to analyze"
      Height          =   252
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6732
   End
End
Attribute VB_Name = "frmNewAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOkay_Click()
    Call OpenFile
End Sub

Private Sub OpenFile()
    Dim sApplicationFileName As String
    
    Me.Visible = False
    
    cdgOpen.InitDir = App.Path & "\sample_files"
    cdgOpen.Filter = PrepareFiles
    cdgOpen.ShowOpen
    If (cdgOpen.CancelError = False) Then
        sApplicationFileName = cdgOpen.FileName
            
        If LenB(sApplicationFileName) Then
            Call frmMain.LockWindows
            If (Dir$(sApplicationFileName, 16) <> "") Then
                frmSourceCode.txtSourceCode.Text = ReadFile(sApplicationFileName, vbNull)
                sourcecode_length = Len(frmSourceCode.txtSourceCode.Text)
                Call CreateMetaCode(frmSourceCode.txtSourceCode.Text)
            End If
            Call frmMain.UnlockWindows
        End If
    End If
    
    Unload Me
End Sub

Private Sub Form_Load()
    Call ListViewAddLanguages
End Sub

Private Sub ListViewAddLanguages()
    Dim List As ListItem
    
    Set List = Me.lsvSources.ListItems.Add(, , "PHP", 1)
    Set List = Me.lsvSources.ListItems.Add(, , "C/C++", 2)
    'Set List = Me.lsvSources.ListItems.Add(, , "Perl", 3)
    'Set List = Me.lsvSources.ListItems.Add(, , "Shell", 4)
    'Set List = Me.lsvSources.ListItems.Add(, , "Java", 5)
    'Set List = Me.lsvSources.ListItems.Add(, , ".NET", 6)
    'Set List = Me.lsvSources.ListItems.Add(, , "Ruby", 7)
    'Set List = Me.lsvSources.ListItems.Add(, , "BAT", 8)
End Sub

Public Function ReadFile(ByRef sFileName As String, ByRef sFilePath As String) As String
    Dim sFileContent As String
    
    If InStrB(1, sFileName, "\", vbBinaryCompare) Then
        sFileName = sFileName
    Else
        sFileName = sFilePath & "\" & sFileName
    End If

    If LenB(Dir$(sFileName)) Then
        Open sFileName For Input As #1
            sFileContent = Input(LOF(1), #1)
        Close

        frmMain.Caption = APP_NAME & " - " & Mid$(sFileName, InStrRev(sFileName, "\", , vbBinaryCompare) + 1)
        sFileContent = Replace(sFileContent, vbCr, "", 1, , vbBinaryCompare)
        ReadFile = Replace(sFileContent, vbLf, vbCrLf, 1, , vbBinaryCompare)
    End If
End Function

Private Sub lsvSources_Click()
    Me.cmdOkay.Enabled = True
End Sub

Private Sub lsvSources_DblClick()
    Call OpenFile
End Sub

Private Function PrepareFiles() As String
    Dim sSelected As String
    Dim sFilter As String
    
    sSelected = lsvSources.SelectedItem.Text
    
    If (sSelected = "PHP") Then
        sFilter = "PHP Files (*.php; *.phps)|*.php*"
    ElseIf (sSelected = "Perl") Then
        sFilter = "Perl Files (*.pl)|*.pl"
    ElseIf (sSelected = "C/C++") Then
        sFilter = "C/C++ Files (*.c; *.cpp)|*.c*"
    ElseIf (sSelected = "Java") Then
        sFilter = "Java Files (*.java)|*.java"
    ElseIf (sSelected = "Ruby") Then
        sFilter = "Ruby Files (*.rb)|*.rb"
    ElseIf (sSelected = "Shell") Then
        sFilter = "Shell Files (*.sh)|*.sh"
    ElseIf (sSelected = ".NET") Then
        sFilter = ".NET Files (*.*)|*.*"
    ElseIf (sSelected = "BAT") Then
        sFilter = "Batch Files (*.bat)|*.bat"
    End If
    
    PrepareFiles = sFilter & "|All Files (*.*)|*.*"
End Function
