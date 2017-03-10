VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMetaCode 
   Caption         =   "MetaCode View"
   ClientHeight    =   4560
   ClientLeft      =   675
   ClientTop       =   345
   ClientWidth     =   5895
   Icon            =   "frmMetaCode.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4557.01
   ScaleMode       =   0  'User
   ScaleWidth      =   5895
   Begin MSComctlLib.ListView lsvMetaCode 
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   8070
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Address"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Token"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Operand"
         Object.Width           =   7066
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Type"
         Object.Width           =   1413
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "If"
         Object.Width           =   707
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Ch"
         Object.Width           =   707
      EndProperty
   End
End
Attribute VB_Name = "frmMetaCode"
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
    Call FormResizer(Me, Me.lsvMetaCode)
End Sub

Private Sub lsvMetaCode_DblClick()
    If (lsvMetaCode.ListItems.Count) Then
        Call SelectAddressInSourceCode(lsvMetaCode.SelectedItem.Text, True)
        Call frmSourceCode.UpdateCoordinates(lsvMetaCode.SelectedItem.Text)
    End If
End Sub

Private Sub lsvMetaCode_ItemClick(ByVal Item As MSComctlLib.ListItem)
    'There is an error if multiple lines are selected; need to be fixed
    On Error Resume Next
    Call SelectAddressInSourceCode(lsvMetaCode.SelectedItem.Text, True)
    Call SelectPointerInBar(GetAddressPosition(lsvMetaCode.SelectedItem.Text))
End Sub

Private Sub lsvMetaCode_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (lsvMetaCode.ListItems.Count) Then
        Call SelectPointerInBar(GetAddressPosition(lsvMetaCode.SelectedItem.Text))
    End If
End Sub
