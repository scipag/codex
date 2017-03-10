VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmStrings 
   Caption         =   "Strings window"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
   Icon            =   "frmStrings.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4245
   ScaleWidth      =   4950
   Begin MSComctlLib.ListView lsvStrings 
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   7435
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Address"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Length"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Type"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "String"
         Object.Width           =   10583
      EndProperty
   End
End
Attribute VB_Name = "frmStrings"
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
    Call FormResizer(Me, Me.lsvStrings)
End Sub

Private Sub lsvStrings_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call ListViewSort(lsvStrings, ColumnHeader)
End Sub

Private Sub lsvStrings_DblClick()
    If (lsvStrings.ListItems.Count) Then
        Call SelectAddressInMetaCode(Me.lsvStrings.SelectedItem.Text)
    End If
End Sub

Private Sub lsvStrings_KeyUp(KeyCode As Integer, Shift As Integer)
    If (lsvStrings.ListItems.Count) Then
        Call SelectPointerInBar(GetAddressPosition(lsvStrings.SelectedItem.Text))
    End If
End Sub

Private Sub lsvStrings_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (lsvStrings.ListItems.Count) Then
        Call SelectPointerInBar(GetAddressPosition(lsvStrings.SelectedItem.Text))
    End If
End Sub
