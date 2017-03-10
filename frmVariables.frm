VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmVariables 
   Caption         =   "Variables window"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   Icon            =   "frmVariables.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4920
   ScaleWidth      =   7455
   Begin MSComctlLib.ListView lsvVariables 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   8705
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
         Text            =   "Name"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Address"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Context"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Infection"
         Object.Width           =   1411
      EndProperty
   End
End
Attribute VB_Name = "frmVariables"
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
    Call FormResizer(Me, Me.lsvVariables)
End Sub

Private Sub lsvVariables_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call ListViewSort(lsvVariables, ColumnHeader)
End Sub

Private Sub lsvVariables_DblClick()
    If (lsvVariables.ListItems.Count) Then
        Call SelectAddressInMetaCode(lsvVariables.SelectedItem.SubItems(1))
    End If
End Sub

Private Sub lsvVariables_KeyUp(KeyCode As Integer, Shift As Integer)
    If (lsvVariables.ListItems.Count) Then
        Call SelectPointerInBar(GetAddressPosition(lsvVariables.SelectedItem.SubItems(1)))
    End If
End Sub

Private Sub lsvVariables_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (lsvVariables.ListItems.Count) Then
        Call SelectPointerInBar(GetAddressPosition(lsvVariables.SelectedItem.SubItems(1)))
    End If
End Sub
