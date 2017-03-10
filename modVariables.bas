Attribute VB_Name = "modVariables"
Option Explicit

Public Sub AddVariableToList(ByRef sMnemonic As String)
    Dim List As ListItem
    
    Set List = frmVariables.lsvVariables.ListItems.Add(, , metacode_operand)
        List.SubItems(1) = metacode_address
        
    If (sMnemonic = "varget") Then
        List.SubItems(2) = "get"
    ElseIf (sMnemonic = "varpost") Then
        List.SubItems(2) = "post"
    ElseIf (sMnemonic = "varcookie") Then
        List.SubItems(2) = "cookie"
    Else
        List.SubItems(2) = "var"
    End If
End Sub

