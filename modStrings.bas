Attribute VB_Name = "modStrings"
Option Explicit

Public Sub AddStringToList(ByRef sString As String)
    Dim List As ListItem
    Dim iStringLength As Integer
    
    iStringLength = Len(sString) - 2
    
    Set List = frmStrings.lsvStrings.ListItems.Add(, , metacode_address)
        List.SubItems(1) = iStringLength
        
        If (iStringLength) Then
            If (IsNumeric(Mid$(sString, 2, iStringLength))) Then
                List.SubItems(2) = "int"
            Else
                List.SubItems(2) = "char"
            End If
        Else
            List.SubItems(2) = "empty"
        End If
        
        List.SubItems(3) = sString
End Sub
