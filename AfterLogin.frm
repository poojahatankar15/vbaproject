VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AfterLogin 
   Caption         =   "Home Page"
   ClientHeight    =   10380
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   20130
   OleObjectBlob   =   "AfterLogin.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AfterLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
buy.Show
End Sub

Private Sub CommandButton2_Click()
borrow.Show
End Sub

Private Sub CommandButton3_Click()
Search.Show
End Sub

Private Sub home_2_Click()
About.Show
End Sub

Private Sub home_3_Click()
Disclaimer.Show
End Sub

Private Sub home_4_Click()
Sponsor.Show
End Sub

Private Sub home_5_Click()
Contact.Show
End Sub

Private Sub logout_Click()
End
End Sub

