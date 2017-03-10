VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSourceCode 
   Caption         =   "Source Code View"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5775
   Icon            =   "frmSourceCode.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4455
   ScaleWidth      =   5775
   Begin MSComctlLib.StatusBar stbStatus 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   4080
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   2487
            MinWidth        =   2469
            Text            =   "Start: 0"
            TextSave        =   "Start: 0"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2469
            MinWidth        =   2469
            Text            =   "End: 0"
            TextSave        =   "End: 0"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2469
            MinWidth        =   2469
            Text            =   "Length: 0"
            TextSave        =   "Length: 0"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Address: unknown"
            TextSave        =   "Address: unknown"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtSourceCode 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   5775
   End
End
Attribute VB_Name = "frmSourceCode"
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
        
        Me.txtSourceCode.Height = Me.Height - Me.stbStatus.Height - 400
        Me.txtSourceCode.Width = Me.Width - 120
    End If
End Sub

Private Sub txtSourceCode_KeyUp(KeyCode As Integer, Shift As Integer)
    Call UpdateCoordinates("unknown")
End Sub

Public Sub UpdateCoordinates(Optional ByRef sAddress As String)
    Dim lStart As Long
    Dim lEnd As Long
    Dim lLength As Long
    
    lStart = Me.txtSourceCode.SelStart + 1
    lLength = Me.txtSourceCode.SelLength
    lEnd = lStart + lLength
    
    Me.stbStatus.Panels(1).Text = "Start: " & lStart
    Me.stbStatus.Panels(2).Text = "End: " & lEnd
    Me.stbStatus.Panels(3).Text = "Length: " & lLength
    Me.stbStatus.Panels(4).Text = "Address: " & sAddress

    Call SelectPointerInBar(lEnd)
End Sub

Private Sub txtSourceCode_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UpdateCoordinates("unknown")
End Sub
