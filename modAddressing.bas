Attribute VB_Name = "modAddressing"
Option Explicit

Public sourcecode_length As Long

Public Function GenerateAddress(ByRef lSourceCodeLine As Long, ByRef lMetaCodeToken As Long, ByRef lMetacodePointer As Long) As String
    GenerateAddress = PadZeros(lSourceCodeLine, Len(CStr(sourcecode_length))) & ":" & PadZeros(lMetaCodeToken, Len(CStr(sourcecode_length))) & ":" & PadZeros(lMetacodePointer, Len(CStr(sourcecode_length)))
End Function

Public Function PadZeros(ByRef lValue As Long, ByRef iTotalLength As Integer) As String
    Dim iPadLength As Integer
    Dim sPadString As String
    Dim i As Integer
    
    iPadLength = iTotalLength - Len(CStr(lValue))
    For i = 1 To iPadLength
        sPadString = sPadString & "0"
    Next i
    
    PadZeros = sPadString & lValue
End Function

Public Sub SelectAddressInMetaCode(ByRef sAddress As String)
    Dim iListItemCount As Integer
    Dim i As Integer
    
    iListItemCount = frmMetaCode.lsvMetaCode.ListItems.Count
    
    For i = 1 To iListItemCount
        If sAddress = frmMetaCode.lsvMetaCode.ListItems(i).Text Then
            frmMetaCode.lsvMetaCode.SelectedItem = frmMetaCode.lsvMetaCode.ListItems(i)
            frmMetaCode.lsvMetaCode.SelectedItem.EnsureVisible
            Call SelectPointerInBar(GetAddressPosition(frmMetaCode.lsvMetaCode.SelectedItem.Text))
            Exit For
        End If
    Next i
    
    If (frmMetaCode.Visible = True) Then
        frmMetaCode.SetFocus
    End If
End Sub

Public Sub SelectAddressInSourceCode(ByRef sAddress As String, ByRef sPositionUpdate As Boolean)
    Dim sAddressBlocks() As String
    
    sAddressBlocks = Split(sAddress, ":", , vbBinaryCompare)
    frmSourceCode.stbStatus.Panels(4).Text = "Address: " & sAddress
    frmSourceCode.txtSourceCode.SelStart = 0
    frmSourceCode.txtSourceCode.SelLength = sAddressBlocks(2)
    If (frmSourceCode.txtSourceCode.Visible = True) Then
        frmSourceCode.txtSourceCode.SetFocus
    End If
    
    If (sPositionUpdate = True) Then
        frmMain.cboPosition.Text = sAddress
    End If
End Sub

Public Function SelectPointerInBar(ByRef sAddress As Long) As Boolean
    Dim lLength As Long

    lLength = Len(frmSourceCode.txtSourceCode.Text) + 1
    
    If (sAddress < lLength) Then
        frmMain.pbrPointer.Value = (sAddress / lLength) * 100
    ElseIf (sAddress = lLength) Then
        frmMain.pbrPointer.Value = 100
    End If
End Function

Public Function GetAddressPosition(ByRef sAddress As String) As Long
    GetAddressPosition = Val(Mid$(sAddress, InStrRev(sAddress, ":", , vbBinaryCompare) + 1))
End Function
