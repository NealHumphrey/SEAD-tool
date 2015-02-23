Option Explicit

Private Sub Worksheet_Activate()
    ActiveWindow.ScrollRow = ActiveCell.row
End Sub