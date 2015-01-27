Option Explicit

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Application.ScreenUpdating = False
    ' Clear the color of all the cells
    Cells.Interior.ColorIndex = 0
    ' Highlight the active cell
    Target.Interior.Color = RGB(124, 150, 178)
    Application.GoTo ActiveCell, True
    Application.ScreenUpdating = True
End Sub