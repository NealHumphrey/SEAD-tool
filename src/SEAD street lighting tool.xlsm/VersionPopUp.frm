VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VersionPopUp 
   Caption         =   "Update Tool"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8025
   OleObjectBlob   =   "VersionPopUp.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "VersionPopUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label2_Click()
    Link = "http://superefficient.org/en/Activities/Procurement/SEAD%20Street%20Lighting%20Evaluation%20Toolkit.aspx"
    On Error GoTo NoCanDo
    ActiveWorkbook.FollowHyperlink Address:=Link, NewWindow:=True
    Unload Me
    Exit Sub
NoCanDo:
    MsgBox "Cannot open " & Link
End Sub
