Attribute VB_Name = "modFormResize"
Option Explicit

Public Sub FormResizer(ByRef fForm As Form, ByRef lListView As ListView)
    If fForm.WindowState <> vbMinimized Then
        If fForm.Height < 1000 Then
            fForm.Height = 1000
        End If
        
        If fForm.Width < 1000 Then
            fForm.Width = 1000
        End If
        
        lListView.Height = fForm.Height - 360
        lListView.Width = fForm.Width - 120
    End If
End Sub
