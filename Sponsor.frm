VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Sponsor 
   Caption         =   "Sponsor"
   ClientHeight    =   8430.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16140
   OleObjectBlob   =   "Sponsor.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Sponsor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdbtn_cancel_Click()
Unload Me
End Sub

Private Sub UserForm_Click()
Sponsor.Show
End Sub
