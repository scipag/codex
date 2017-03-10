Attribute VB_Name = "modListView"
Option Explicit

Public Sub ListViewSort(ByRef ListViewControl As ListView, ByRef Column As ColumnHeader)
    If ListViewControl.SortKey <> Column.Index - 1 Then
        ListViewControl.SortKey = Column.Index - 1
        ListViewControl.SortOrder = lvwAscending
    Else
        If ListViewControl.SortOrder = lvwAscending Then
            ListViewControl.SortOrder = lvwDescending
        Else
            ListViewControl.SortOrder = lvwAscending
        End If
    End If
    
    ListViewControl.Sorted = -1
End Sub

