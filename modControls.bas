Attribute VB_Name = "modControls"
Option Explicit

Public Sub AddControlToList(ByRef sAction As String, ByRef sControl)
    Dim List As ListItem
    
    Set List = frmControls.lsvControls.ListItems.Add(, , metacode_address)
        List.SubItems(1) = sControl
        List.SubItems(2) = sAction
        List.SubItems(3) = metacode_iflevel
End Sub

Public Sub GenerateLogic()
    Dim List As ListItem
    Dim iIndex As Integer
    Dim sAddressStart As String
    Dim lCount As Long
    Dim sLogikToken As String
    Dim sLogic As String
    Dim bEntryPoint As Boolean
    Dim lControlsCount As Long
    Dim lMetaCodeCount As Long
    
    lControlsCount = frmControls.lsvControls.ListItems.Count
    Call frmLog.AddLogEntry("Generating logic for " & lControlsCount & " controls ...")
    If (lControlsCount > 0) Then
        For iIndex = 1 To lControlsCount
            sLogic = vbNullString
            sAddressStart = frmControls.lsvControls.ListItems(iIndex).Text
            
            If (iIndex Mod 50) = 0 Then
                Call frmAnalyzing.SetAnalysisStatus(iIndex / (lControlsCount / 100), "Second Pass (Logic): " & sAddressStart)
            End If
            
            lMetaCodeCount = frmMetaCode.lsvMetaCode.ListItems.Count
            For lCount = 1 To lMetaCodeCount
                If (frmMetaCode.lsvMetaCode.ListItems(lCount).ListSubItems(1).Text <> MetaDefinition("condition_endif")) Then
                    If (bEntryPoint = True) Then
                        sLogikToken = frmMetaCode.lsvMetaCode.ListItems(lCount).ListSubItems(1).Text
                        
                        If (sLogikToken <> MetaDefinition("condition_then")) Then
                            If (sLogikToken = MetaDefinition("compare_not")) Then
                                sLogic = sLogic & "!"
                            ElseIf (sLogikToken = MetaDefinition("parenthesis_open")) Then
                                sLogic = sLogic & "("
                            ElseIf (sLogikToken = MetaDefinition("parenthesis_close")) Then
                                sLogic = sLogic & ")"
                            ElseIf (sLogikToken = MetaDefinition("condition_land")) Then
                                sLogic = sLogic & "&&"
                            ElseIf (sLogikToken = MetaDefinition("condition_lor")) Then
                                sLogic = sLogic & "||"
                            ElseIf (sLogikToken = MetaDefinition("compare_equal")) Then
                                sLogic = sLogic & "="
                            ElseIf (sLogikToken = MetaDefinition("compare_lthan")) Then
                                sLogic = sLogic & "<"
                            ElseIf (sLogikToken = MetaDefinition("compare_gthan")) Then
                                sLogic = sLogic & ">"
                            ElseIf (sLogikToken = MetaDefinition("compare_lethan")) Then
                                sLogic = sLogic & "=<"
                            ElseIf (sLogikToken = MetaDefinition("compare_gethan")) Then
                                sLogic = sLogic & ">="
                            Else
                                sLogic = sLogic & sLogikToken
                            End If
                        Else
                            bEntryPoint = False
                        End If
                    End If
                    
                    If (frmMetaCode.lsvMetaCode.ListItems(lCount).Text = sAddressStart) Then
                        bEntryPoint = True
                    End If
                End If
            Next lCount
            
            Set List = frmControls.lsvControls.ListItems.Item(iIndex)
                List.SubItems(4) = sLogic
        Next iIndex
    End If
    Call frmLog.AddLogEntry("Dissection of control structures and logic conditions done.")
End Sub
