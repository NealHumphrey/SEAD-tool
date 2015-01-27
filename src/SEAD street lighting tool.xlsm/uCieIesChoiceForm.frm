VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uCieIesChoiceForm 
   Caption         =   "UserForm1"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3825
   OleObjectBlob   =   "uCieIesChoiceForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uCieIesChoiceForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CIE_Click()
Dim choice As String
choice = "CIE"
Unload Me
upgradePlot (choice)
End Sub

Private Sub IES_Click()
Dim choice As String
choice = "IES"
Unload Me
upgradePlot (choice)

End Sub

