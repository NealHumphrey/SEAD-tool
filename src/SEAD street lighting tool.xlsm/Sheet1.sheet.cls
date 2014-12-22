Option Explicit

Private Sub worksheet_change(ByVal Target As Range)
    ActiveSheet.Calculate
End Sub

Private Sub worksheet_activate()
    ActiveSheet.Calculate
End Sub