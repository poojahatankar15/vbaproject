VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Registration 
   Caption         =   "Registration"
   ClientHeight    =   10035
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   19215
   OleObjectBlob   =   "Registration.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Registration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub comd_cancel_Click()
Unload Me
End Sub

Private Sub CommandButton2_Click()
TextBox1.Text = " "
TextBox2.Text = " "
TextBox3.Text = " "
TextBox4.Text = " "
ComboBox1.Text = " "
ComboBox2.Text = " "
ComboBox3.Text = " "
ComboBox4.Text = " "
End Sub

Private Sub CommandButton3_Click()
Sheet1.Activate
emptyRow = WorksheetFunction.CountA(Range("A:A"))
Range("A1").Value = "Name"
Range("B1").Value = "Address"
Range("C1").Value = "Date of Birth"
Range("D1").Value = "Gender"
Range("E1").Value = "E-mail Id"
Range("F1").Value = "City"
Range("G1").Value = "Educational Role"

If btn_male.Value = True Then
Cells(emptyRow, 4).Value = "Male"
Else
Cells(emptyRow, 4).Value = "Female"
End If


Sheet1.Activate
emptyRow = WorksheetFunction.CountA(Range("A:A")) + 1
Cells(emptyRow, 1).Value = TextBox1.Value
Cells(emptyRow, 2).Value = TextBox2.Value
Cells(emptyRow, 3).Value = ComboBox1.Value + ComboBox2.Value + ComboBox3.Value
Cells(emptyRow, 5).Value = TextBox3.Value
Cells(emptyRow, 6).Value = TextBox4.Value
Cells(emptyRow, 7).Value = ComboBox4.Value

Done.Show
End Sub




Private Sub lib_copyright_Click()

End Sub

Private Sub UserForm_Initialize()
With ComboBox1
.AddItem ("1")
.AddItem ("2")
.AddItem ("3")
.AddItem ("4")
.AddItem ("5")
.AddItem ("6")
.AddItem ("7")
.AddItem ("8")
.AddItem ("9")
.AddItem ("10")
.AddItem ("11")
.AddItem ("12")
.AddItem ("13")
.AddItem ("14")
.AddItem ("15")
.AddItem ("16")
.AddItem ("17")
.AddItem ("18")
.AddItem ("19")
.AddItem ("20")
.AddItem ("21")
.AddItem ("22")
.AddItem ("23")
.AddItem ("24")
.AddItem ("25")
.AddItem ("26")
.AddItem ("27")
.AddItem ("28")
.AddItem ("29")
.AddItem ("30")
.AddItem ("31")
End With

With ComboBox2
.AddItem ("Jan")
.AddItem ("Feb")
.AddItem ("March")
.AddItem ("April")
.AddItem ("May")
.AddItem ("June")
.AddItem ("July")
.AddItem ("Aug")
.AddItem ("Sep")
.AddItem ("Oct")
.AddItem ("Nov")
.AddItem ("Dec")
End With

With ComboBox3
.AddItem ("1990")
.AddItem ("1991")
.AddItem ("1992")
.AddItem ("1993")
.AddItem ("1994")
.AddItem ("1995")
.AddItem ("1996")
.AddItem ("1997")
.AddItem ("1998")
.AddItem ("1999")
.AddItem ("2000")
.AddItem ("2001")
.AddItem ("2002")
.AddItem ("2003")
.AddItem ("2004")
.AddItem ("2005")
.AddItem ("2006")
.AddItem ("2007")
.AddItem ("2008")
.AddItem ("2009")
.AddItem ("2010")
.AddItem ("2011")
.AddItem ("2012")
.AddItem ("2013")
.AddItem ("2014")
.AddItem ("2015")
.AddItem ("2016")
.AddItem ("2017")
.AddItem ("2018")
.AddItem ("2019")

End With

With ComboBox4
.AddItem ("I to IV")
.AddItem ("V to VIII")
.AddItem ("IX to X")
.AddItem ("XI to XII")
.AddItem ("UG and PG")
.AddItem ("Other")
End With


End Sub
