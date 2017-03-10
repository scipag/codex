VERSION 5.00
Begin VB.Form frmLog 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Log"
   ClientHeight    =   2100
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   8436
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   8436
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstLog 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1884
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8415
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If (UnloadMode = vbFormControlMenu) Then
        Cancel = True
        Me.Hide
    End If
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        If Me.Height < 1000 Then
            Me.Height = 1000
        End If
        
        If Me.Width < 1000 Then
            Me.Width = 1000
        End If
        
        Me.lstLog.Height = Me.Height - 280
        Me.lstLog.Width = Me.Width - 120
    End If
End Sub

Public Sub AddLogEntry(ByRef sMessage As String)
    Me.lstLog.AddItem sMessage
    lstLog.ListIndex = Me.lstLog.ListCount - 1
End Sub

Private Sub lstLog_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = 2) Then
        Call OpenContextMenu(Me, frmMain.mnuLog)
    End If
End Sub
