Attribute VB_Name = "modMetaCode"
Option Explicit

Public metacode_pointer As Long
Public metacode_line As Long
Public metacode_tokenid As Long
Public metacode_address As String
Public metacode_mnemonic As String
Public metacode_operand As String
Public metacode_type As String
Public metacode_childlevel As Integer
Public metacode_iflevel As Integer

Public metacode_chunk As clsConcat

Public Const metacode_definition_string As String = "string"
Public Const metacode_definition_var As String = "var"
Public Const metacode_definition_varget As String = "varget"
Public Const metacode_definition_varpost As String = "varpost"
Public Const metacode_definition_varcookie As String = "varcookie"
Public Const metacode_definition_varserver As String = "varserver"
Public Const metacode_definition_varunknown As String = "varunknown"

Public Sub CreateMetaCode(ByRef sSourceCodeFull As String)
    Dim sMetaCodeFull As String
    Dim lSourceCodeLength As Long
    Dim lOperandPosition As Long
    Dim i As Long
    Dim lInput As Long
    Dim iSourceCodeDefinitions As Integer
    
    Dim metacode_iflevelstop As Boolean
    Dim metacode_childlevelstop As Boolean
    Dim metacode_ignoreerrors As Boolean
    Dim metacode_errorcount As Integer
    
    Set metacode_chunk = New clsConcat
    
    lSourceCodeLength = Len(sSourceCodeFull)
    Call frmLog.AddLogEntry("Compiling MetaCode out of " & lSourceCodeLength & " bytes source code ...")
    
    Call InitializeCodeDefinitions(App.Path & sourcecode_lexer_directory & "php.lexer", sourcecode_definitions)
    Call InitializeCodeDefinitions(App.Path & sourcecode_lexer_directory & "metacode.lexer", metacode_definitions)
    
    frmMain.cboPosition.Clear
    frmControls.lsvControls.ListItems.Clear
    frmMetaCode.lsvMetaCode.ListItems.Clear
    frmVariables.lsvVariables.ListItems.Clear
    frmStrings.lsvStrings.ListItems.Clear
    metacode_chunk.Clear
    metacode_tokenid = 0
    metacode_line = 1
    metacode_childlevel = 0
    metacode_iflevel = 0
    
    metacode_iflevelstop = False
    metacode_childlevelstop = False
    metacode_ignoreerrors = False
    metacode_errorcount = 0
    
    For metacode_pointer = 1 To lSourceCodeLength
        metacode_chunk.Concat Mid$(sSourceCodeFull, metacode_pointer, 1)
        
        If (metacode_pointer Mod 5000) = 0 Then
            Call frmAnalyzing.SetAnalysisStatus(metacode_pointer / (lSourceCodeLength / 100), "First pass (MetaCode): " & metacode_address)
        End If
        
        If (Mid$(sSourceCodeFull, metacode_pointer, 2) = vbCrLf) Then
            metacode_line = metacode_line + 1
        End If
        
        If InStrB(1, metacode_chunk.Value, sourcecode_definition_quotedouble, vbBinaryCompare) Then
            If InStr(metacode_pointer + 1, sSourceCodeFull, sourcecode_definition_quotedouble, vbBinaryCompare) Then
                metacode_operand = Mid$(sSourceCodeFull, metacode_pointer, (InStr(metacode_pointer + 1, sSourceCodeFull, sourcecode_definition_quotedouble, vbBinaryCompare) - metacode_pointer) + 1)
                metacode_pointer = metacode_pointer + Len(metacode_operand) - 1
                Call AddMetaCode(metacode_definition_string, metacode_operand)
            End If
        ElseIf InStrB(1, metacode_chunk.Value, sourcecode_definition_quotesingle, vbBinaryCompare) Then
            If InStr(metacode_pointer + 1, sSourceCodeFull, sourcecode_definition_quotesingle, vbBinaryCompare) Then
                metacode_operand = Mid$(sSourceCodeFull, metacode_pointer, (InStr(metacode_pointer + 1, sSourceCodeFull, sourcecode_definition_quotesingle, vbBinaryCompare) - metacode_pointer) + 1)
                metacode_pointer = metacode_pointer + Len(metacode_operand) - 1
                Call AddMetaCode(metacode_definition_string, metacode_operand)
            End If
        ElseIf InStrB(1, metacode_chunk.Value, "$", vbBinaryCompare) Then
            If (Mid$(sSourceCodeFull, metacode_pointer + 1, 1) = "_") Then
                If (Mid$(sSourceCodeFull, metacode_pointer + 2, 3) = "GET") Then
                    metacode_operand = Mid$(sSourceCodeFull, metacode_pointer + 7, (InStr(metacode_pointer + 7, sSourceCodeFull, "'", vbBinaryCompare) - 7) - (metacode_pointer))
                    metacode_pointer = metacode_pointer + Len(metacode_operand) + 8
                    Call AddMetaCode(metacode_definition_varget, metacode_operand)
                ElseIf (Mid$(sSourceCodeFull, metacode_pointer + 2, 4) = "POST") Then
                    metacode_operand = Mid$(sSourceCodeFull, metacode_pointer + 8, (InStr(metacode_pointer + 8, sSourceCodeFull, "'", vbBinaryCompare) - 8) - (metacode_pointer))
                    metacode_pointer = metacode_pointer + Len(metacode_operand) + 9
                    Call AddMetaCode(metacode_definition_varpost, metacode_operand)
                ElseIf (Mid$(sSourceCodeFull, metacode_pointer + 2, 6) = "COOKIE") Then
                    metacode_operand = Mid$(sSourceCodeFull, metacode_pointer + 10, (InStr(metacode_pointer + 10, sSourceCodeFull, "'", vbBinaryCompare) - 10) - (metacode_pointer))
                    metacode_pointer = metacode_pointer + Len(metacode_operand) + 11
                    Call AddMetaCode(metacode_definition_varcookie, metacode_operand)
                ElseIf (Mid$(sSourceCodeFull, metacode_pointer + 2, 6) = "SERVER") Then
                    metacode_operand = Mid$(sSourceCodeFull, metacode_pointer + 10, (InStr(metacode_pointer + 10, sSourceCodeFull, "'", vbBinaryCompare) - 10) - (metacode_pointer))
                    metacode_pointer = metacode_pointer + Len(metacode_operand) + 11
                    Call AddMetaCode(metacode_definition_varserver, metacode_operand)
                Else
                    Call AddMetaCode(metacode_definition_varunknown, "")
                End If
            Else
                For lOperandPosition = 1 To lSourceCodeLength
                    metacode_operand = Mid$(sSourceCodeFull, metacode_pointer + lOperandPosition, 1)
                    
                    If InStrB(1, AllowedCharsInVariables, metacode_operand, vbBinaryCompare) = 0 Then
                        metacode_operand = Mid$(sSourceCodeFull, metacode_pointer + 1, lOperandPosition - 1)
                        Exit For
                    End If
                Next lOperandPosition
                
                metacode_pointer = metacode_pointer + Len(metacode_operand)
                Call AddMetaCode(metacode_definition_var, metacode_operand)
            End If
        Else
            'Real recursive compiler
            iSourceCodeDefinitions = UBound(sourcecode_definitions)
            For i = 0 To iSourceCodeDefinitions
                'Lookahead mechanism
                If (LCase(Mid$(sSourceCodeFull, metacode_pointer, Len(sourcecode_definitions(i, 1)))) = LCase(sourcecode_definitions(i, 1))) Then
                    metacode_pointer = metacode_pointer + Len(sourcecode_definitions(i, 1)) - 1
                    Call AddMetaCode(Source2Meta(sourcecode_definitions(i, 1)), "")
                    Exit For
                ElseIf InStrB(1, LCase(metacode_chunk.Value), LCase(sourcecode_definitions(i, 1)), vbBinaryCompare) Then
                    Call AddMetaCode(Source2Meta(sourcecode_definitions(i, 1)), "")
                    Exit For
                End If
            Next i
        End If
    
        'Lexical error detection
        If (metacode_childlevel < 0) Then
            metacode_errorcount = metacode_errorcount + 1
            If (metacode_ignoreerrors = False) Then
                metacode_childlevelstop = True
                Call frmLog.AddLogEntry("Error #" & metacode_errorcount & ": Syntax error (child_level) found at " & metacode_address & ".")
            End If
        End If
        
        If (metacode_iflevel < 0 And metacode_ignoreerrors = False) Then
            metacode_errorcount = metacode_errorcount + 1
            If (metacode_ignoreerrors = False) Then
                metacode_iflevelstop = True
                Call frmLog.AddLogEntry("Error #" & metacode_errorcount & ": Syntax error (condition_level) found at " & metacode_address & ".")
            End If
        End If
        
        If (metacode_childlevelstop = True And metacode_ignoreerrors = False) Then
            lInput = MsgBox("During compilation of the source code file an error within" & vbCrLf & _
                            "the child definitions has been found. Would you like to resume" & vbCrLf & _
                            "or to abort the compilation? (Note: Ignoring such errors might" & vbCrLf & _
                            "result in inadequate program modelling and analysis.)", vbAbortRetryIgnore + vbInformation, "Compilation error child definition")
            
            If (lInput = 4) Then
                metacode_childlevelstop = False
                metacode_childlevel = 0
            ElseIf (lInput = 3) Then
                Exit For
            ElseIf (lInput = 5) Then
                metacode_ignoreerrors = True
            End If
        End If
    
        If (metacode_iflevelstop = True And metacode_ignoreerrors = False) Then
            lInput = MsgBox("During compilation of the source code file an error within" & vbCrLf & _
                            "the if level definitions has been found. Would you like to resume" & vbCrLf & _
                            "or to abort the compilation? (Note: Ignoring such errors might" & vbCrLf & _
                            "result in inadequate program modelling and analysis.)", vbAbortRetryIgnore + vbInformation, "Compilation error if definition")
            
            If (lInput = 4) Then
                metacode_iflevelstop = False
                metacode_iflevel = 0
            ElseIf (lInput = 3) Then
                Exit For
            ElseIf (lInput = 5) Then
                metacode_ignoreerrors = True
            End If
        End If
    
    Next metacode_pointer
    
    Call frmLog.AddLogEntry("MetaCode compilation done. " & frmMetaCode.lsvMetaCode.ListItems.Count & " tokens determined. (" & metacode_errorcount & " errors)")
    Call GenerateLogic
    Call frmLog.AddLogEntry("The initial autoanalysis has been finished.")
    Unload frmAnalyzing
End Sub

Public Function AddMetaCode(ByRef sMnemonic As String, ByRef sOperand As String) As String
    Dim List As ListItem
    
    metacode_tokenid = metacode_tokenid + 1
    metacode_address = GenerateAddress(metacode_line, metacode_tokenid, metacode_pointer)

    If (Mid$(sMnemonic, 1, Len(metacode_definition_var)) = metacode_definition_var) Then
        Call AddVariableToList(sMnemonic)
        metacode_type = "Var"
    ElseIf (sMnemonic = metacode_definition_string) Then
        Call AddStringToList(sOperand)
        metacode_type = "Stmt"
    ElseIf (sMnemonic = MetaDefinition("condition_if")) Then
        metacode_iflevel = metacode_iflevel + 1
        Call AddControlToList("+", sMnemonic)
        metacode_type = "Cond"
    ElseIf (sMnemonic = MetaDefinition("condition_while")) Then
        metacode_iflevel = metacode_iflevel + 1
        Call AddControlToList("+", sMnemonic)
        metacode_type = "Cond"
    ElseIf (sMnemonic = MetaDefinition("condition_for")) Then
        metacode_iflevel = metacode_iflevel + 1
        Call AddControlToList("+", sMnemonic)
        metacode_type = "Cond"
    ElseIf (sMnemonic = MetaDefinition("condition_do")) Then
        metacode_iflevel = metacode_iflevel + 1
        Call AddControlToList("+", sMnemonic)
        metacode_type = "Cond"
    ElseIf (sMnemonic = MetaDefinition("condition_switch")) Then
        metacode_iflevel = metacode_iflevel + 1
        Call AddControlToList("+", sMnemonic)
        metacode_type = "Cond"
    ElseIf (sMnemonic = MetaDefinition("condition_elseif")) Then
        metacode_iflevel = metacode_iflevel + 1
        Call AddControlToList("+", sMnemonic)
        metacode_type = "Cond"
    ElseIf (sMnemonic = MetaDefinition("condition_else")) Then
        metacode_iflevel = metacode_iflevel + 1
        Call AddControlToList("+", sMnemonic)
        metacode_type = "Cond"
    ElseIf (sMnemonic = MetaDefinition("condition_endif")) Then
        metacode_iflevel = metacode_iflevel - 1
        Call AddControlToList("-", sMnemonic)
        metacode_type = "Cond"
    ElseIf (sMnemonic = MetaDefinition("parenthesis_open")) Then
        metacode_childlevel = metacode_childlevel + 1
        metacode_type = "Chld"
    ElseIf (sMnemonic = MetaDefinition("parenthesis_close")) Then
        metacode_childlevel = metacode_childlevel - 1
        metacode_type = "Chld"
    End If
    
    Set List = frmMetaCode.lsvMetaCode.ListItems.Add(, , metacode_address)
        List.SubItems(1) = sMnemonic
        List.SubItems(2) = metacode_operand
        List.SubItems(3) = metacode_type
        List.SubItems(4) = metacode_iflevel
        List.SubItems(5) = metacode_childlevel
    
    frmMain.cboPosition.AddItem metacode_address
    
    Call ResetMetaCodeToken
End Function

Public Sub ResetMetaCodeToken()
    metacode_chunk.Clear
    metacode_operand = ""
    metacode_type = ""
End Sub

Public Function MetaDefinition(ByRef sStatement As String) As String
    Dim i As Integer
    Dim iMetaCodeDefinitionArrayUpper As Integer
    
    iMetaCodeDefinitionArrayUpper = UBound(metacode_definitions)
    
    For i = 0 To iMetaCodeDefinitionArrayUpper
        If metacode_definitions(i, 0) = sStatement Then
            MetaDefinition = metacode_definitions(i, 1)
            Exit For
        End If
    Next i
End Function
