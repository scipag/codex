Attribute VB_Name = "modContextMenu"
Option Explicit

Public Const WM_RBUTTONDOWN As Integer = &H204

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Sub OpenContextMenu(FormName As Form, MenuName As Menu)
    Call SendMessage(FormName.hwnd, WM_RBUTTONDOWN, 0, 0&)
    FormName.PopupMenu MenuName
End Sub

