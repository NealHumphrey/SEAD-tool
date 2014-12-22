Private Sub worksheet_change(ByVal Target As Range)
    Sheet13.EnableCalculation = True
    Sheet13.EnableCalculation = False
End Sub

Private Sub worksheet_activate()
    Sheet13.EnableCalculation = True
    Sheet13.EnableCalculation = False
End Sub
