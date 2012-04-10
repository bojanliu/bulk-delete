Attribute VB_Name = "批量删除模块"
Option Explicit

'批量删除超链接
Sub del_hyperlink()
    Dim ws As Worksheet
    Dim hl As Hyperlink
    For Each ws In ThisWorkbook.Worksheets
        For Each hl In ws.Hyperlinks
            hl.Delete
        Next hl
    Next ws
End Sub

'批量删除批注
Sub del_comment()
    Dim ws As Worksheet
    Dim cm As Comment
    For Each ws In ThisWorkbook.Worksheets
        For Each cm In ws.Comments
            cm.Delete
        Next cm
    Next ws
End Sub

'批量删除名称
Sub del_name()
    Dim nm As Name
    For Each nm In ThisWorkbook.Names
        nm.Delete
    Next nm
End Sub

'批量清除条件格式
Sub del_formatcondition()
    Cells.FormatConditions.Delete
End Sub

'批量删除图表
Sub del_shape()
    Dim ws As Worksheet
    Dim sp As Shape
    For Each ws In ThisWorkbook.Worksheets
        For Each sp In ws.Shapes
            sp.Delete
        Next sp
    Next ws
End Sub
