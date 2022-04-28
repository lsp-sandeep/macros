Attribute VB_Name = "Module2"
Sub RemoveLinks()
'
' LinkPathCellToNameCell_V2 Macro
'

'
    last_row_num = Range("A1").CurrentRegion.Rows.Count
    For i = 2 To last_row_num
        Range("D" & i).Select
        Selection.Hyperlinks(1).Delete
    Next i
End Sub
Sub LinkPathCellToNameCell()
'
' LinkPathCellToNameCell_V2 Macro
'

'
    last_row_num = Range("A1").CurrentRegion.Rows.Count
    For i = 2 To last_row_num
        Range("D" & i).Select
        ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:=Range("E" & i).Value
    Next i
End Sub
Sub LinkFolderPathCellToNameCell()
'
' LinkPathCellToNameCell_V2 Macro
'

'
    last_row_num = Range("A1").CurrentRegion.Rows.Count
    For i = 2 To last_row_num
        Range("B" & i).Select
        ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:=Range("F" & i).Value
    Next i
End Sub
