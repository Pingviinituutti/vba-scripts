Sub FillBlanksWithZero()
  ' FillBlanksWithZero Macro
  ' Fills zero in each selected cell that is blank

    For Each cell In Selection
      If IsEmpty(cell) Then
        cell.Value = 0
      End If
    Next
End Sub
