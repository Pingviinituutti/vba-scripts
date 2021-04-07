Sub DeleteEmptyCellsShiftToLeft()
    ' DeleteEmptyCellsShiftToLeft Macro
    Selection.SpecialCells(xlCellTypeBlanks).Delete Shift:=xlToLeft
End Sub
