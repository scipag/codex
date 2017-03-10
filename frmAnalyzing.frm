VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAnalyzing 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Analyzing ..."
   ClientHeight    =   1365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7140
   Icon            =   "frmAnalyzing.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MousePointer    =   11  'Hourglass
   Moveable        =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer timIncrease 
      Interval        =   1000
      Left            =   6240
      Top             =   120
   End
   Begin MSComctlLib.ProgressBar pbrAnalysis 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lblAnalyzingObject 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Please wait!"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   6855
   End
   Begin VB.Label lblAnalysisStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Analyzing ..."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmAnalyzing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Top = (frmMain.Height / 2) - (Me.Height)
    Me.Left = (frmMain.Width / 2) - (Me.Width / 2)
    Me.pbrAnalysis.Value = 0
End Sub

Public Sub SetAnalysisStatus(ByRef lValue As Long, ByRef sDescription As String)
    Me.lblAnalysisStatus.Caption = "Analyze ... (" & lValue & "%)"
    Me.pbrAnalysis.Value = lValue
    frmMain.pbrPointer.Value = lValue
    Me.lblAnalyzingObject.Caption = sDescription
    DoEvents
End Sub

Private Sub timIncrease_Timer()
    If (Me.pbrAnalysis.Value < 100) Then
        Me.pbrAnalysis.Value = Me.pbrAnalysis.Value + 1
    Else
        Me.timIncrease.Enabled = False
        Unload Me
    End If
End Sub
