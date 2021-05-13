VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Login 
   Caption         =   "Log-In"
   ClientHeight    =   9945.001
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   19605
   OleObjectBlob   =   "Login.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdbtn_cancel_Click()
Unload Me
End Sub

Private Sub cmdbtn_register_Click()
Registration.Show
End Sub


Private Sub cmdntn_login_Click()
If txt1.Value = "Admin" Then
If txt2.Value = "1234" Then
LoginFlag = True
AfterLogin.Show
Exit Sub
End If
End If
MsgBox "Sorry, Incorrect Login Details"
End Sub
