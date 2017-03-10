Attribute VB_Name = "modSourceCode"
Option Explicit

Public Const AllowedCharsInVariables As String = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890_-[]'"
Public Const AllowedCharsInNumerals As String = "1234567890,."

Public Const sourcecode_definition_quotedouble As String = """"
Public Const sourcecode_definition_quotesingle As String = "'"

Public Const sourcecode_lexer_directory As String = "\lexer\"

Public sourcecode_definitions() As String
Public metacode_definitions() As String

Public Sub InitializeCodeDefinitions(ByRef sLexerFileName As String, ByRef sPublicArray() As String)
    Dim i As Integer
    Dim sInput As String
    Dim sLines() As String
    Dim sEntry() As String
    Dim iLinesArrayUpper As Integer
    
    If LenB(Dir$(sLexerFileName)) Then
        Open sLexerFileName For Input As #1
            sInput = Input(LOF(1), #1)
        Close
        sLines = Split(sInput, vbCrLf, , vbBinaryCompare)
        iLinesArrayUpper = UBound(sLines)
        
        ReDim sPublicArray(0 To iLinesArrayUpper, 0 To 1)
        
        For i = 0 To iLinesArrayUpper
            sEntry = Split(sLines(i), ";", , vbBinaryCompare)
            
            sPublicArray(i, 0) = sEntry(0)
            sPublicArray(i, 1) = sEntry(1)
        Next i
    
        Call frmLog.AddLogEntry("Loading lexer module " & sLexerFileName & " ...")
    End If
End Sub

Public Function Source2Meta(ByRef sSource As String) As String
    Dim i As Integer
    Dim j As Integer
    Dim iSourceCodeDefinitionArrayUpper As Integer
    Dim iMetaCodeDefinitionArrayUpper As Integer
    
    iSourceCodeDefinitionArrayUpper = UBound(sourcecode_definitions)
    iMetaCodeDefinitionArrayUpper = UBound(metacode_definitions)
    
    For i = 0 To iSourceCodeDefinitionArrayUpper
        If sourcecode_definitions(i, 1) = sSource Then
            For j = 0 To iMetaCodeDefinitionArrayUpper
                If sourcecode_definitions(i, 0) = metacode_definitions(j, 0) Then
                    Source2Meta = metacode_definitions(j, 1)
                    Exit Function
                End If
            Next j
        End If
    Next i
End Function

Public Function Meta2Source(ByRef sMeta As String) As String
    Dim i As Integer
    Dim j As Integer
    Dim iSourceCodeDefinitionArrayUpper As Integer
    Dim iMetaCodeDefinitionArrayUpper As Integer
    
    iSourceCodeDefinitionArrayUpper = UBound(sourcecode_definitions)
    iMetaCodeDefinitionArrayUpper = UBound(metacode_definitions)
    
    For i = 0 To iMetaCodeDefinitionArrayUpper
        If metacode_definitions(i, 1) = sMeta Then
            For j = 0 To iSourceCodeDefinitionArrayUpper
                If sourcecode_definitions(i, 0) = metacode_definitions(j, 0) Then
                    Meta2Source = sourcecode_definitions(j, 1)
                    Exit Function
                End If
            Next j
        End If
    Next i
End Function

Public Function SourceDefinition(ByRef sStatement As String) As String
    Dim i As Integer
    Dim iSourceCodeDefinitionArrayUpper As Integer
    
    iSourceCodeDefinitionArrayUpper = UBound(sourcecode_definitions)
    
    For i = 0 To iSourceCodeDefinitionArrayUpper
        If sourcecode_definitions(i, 0) = sStatement Then
            SourceDefinition = sourcecode_definitions(i, 1)
            Exit For
        End If
    Next i
End Function
