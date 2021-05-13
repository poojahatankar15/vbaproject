VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} buy 
   Caption         =   "Purchase"
   ClientHeight    =   10800
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15870
   OleObjectBlob   =   "buy.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "buy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub comd_cancel_Click()
Unload Me
End Sub

Private Sub CommandButton3_Click()
Sheet4.Activate
emptyRow = WorksheetFunction.CountA(Range("A:A"))
Range("A1").Value = "Member Name"
Range("B1").Value = "Book Name"
Range("C1").Value = "Author"
Range("D1").Value = "Publisher"
Range("E1").Value = "Section"
Range("F1").Value = "Book No."
Range("G1").Value = "Price"
Range("H1").Value = "Copies of Book"
Range("I1").Value = "Method of Payment"

Sheet4.Activate
emptyRow = WorksheetFunction.CountA(Range("A:A")) + 1
Cells(emptyRow, 1).Value = txt9.Value
Cells(emptyRow, 2).Value = ComboBox1.Value
Cells(emptyRow, 3).Value = txt1.Value
Cells(emptyRow, 4).Value = txt2.Value
Cells(emptyRow, 5).Value = txt3.Value
Cells(emptyRow, 6).Value = txt4.Value
Cells(emptyRow, 7).Value = txt5.Value
Cells(emptyRow, 8).Value = txt6.Value

If ob_card.Value = True Then
Cells(emptyRow, 9).Value = "Card"
Else
Cells(emptyRow, 9).Value = "Cash"
End If
Done2.Show
End Sub

Private Sub ComboBox1_Change()
If ComboBox1.ListIndex = 0 Then
txt1.Text = "Stephen Hawking"
txt2.Text = "Bantam Books"
txt3.Text = "Non-Fiction"
txt4.Text = "B113"
txt5.Text = "359"
End If

If ComboBox1.ListIndex = 1 Then
txt1.Text = "Oliver Sacks"
txt2.Text = "Drovers"
txt3.Text = "Non-Fiction"
txt4.Text = "B114"
txt5.Text = "193"
End If

If ComboBox1.ListIndex = 2 Then
txt1.Text = "Lewis Carroll"
txt2.Text = "Macmillan"
txt3.Text = "Literature"
txt4.Text = "G128"
txt5.Text = "450"
End If

If ComboBox1.ListIndex = 3 Then
txt1.Text = "J.K Rowling"
txt2.Text = "Bloomsbury"
txt3.Text = "Fiction"
txt4.Text = "A100"
txt5.Text = "150"
End If

If ComboBox1.ListIndex = 4 Then
txt1.Text = "J.K Rowling"
txt2.Text = "Bloomsbury"
txt3.Text = "Fiction"
txt4.Text = "A101"
txt5.Text = "190"
End If

If ComboBox1.ListIndex = 5 Then
txt1.Text = "J.K Rowling"
txt2.Text = "Bloomsbury"
txt3.Text = "Fiction"
txt4.Text = "A102"
txt5.Text = "200"
End If

If ComboBox1.ListIndex = 6 Then
txt1.Text = "J.K Rowling"
txt2.Text = "Bloomsbury"
txt3.Text = "Fiction"
txt4.Text = "A103"
txt5.Text = "285"
End If

If ComboBox1.ListIndex = 7 Then
txt1.Text = "J.K Rowling"
txt2.Text = "Bloomsbury"
txt3.Text = "Fiction"
txt4.Text = "A104"
txt5.Text = "299"
End If

If ComboBox1.ListIndex = 8 Then
txt1.Text = "J.K Rowling"
txt2.Text = "Bloomsbury"
txt3.Text = "Fiction"
txt4.Text = "A105"
txt5.Text = "413"
End If

If ComboBox1.ListIndex = 9 Then
txt1.Text = "J.K Rowling"
txt2.Text = "Bloomsbury"
txt3.Text = "Fiction"
txt4.Text = "A106"
txt5.Text = "589"
End If

If ComboBox1.ListIndex = 10 Then
txt1.Text = "John Hersey"
txt2.Text = "Alfred A Knopf"
txt3.Text = "Non-Fiction"
txt4.Text = "B115"
txt5.Text = "248"
End If

If ComboBox1.ListIndex = 11 Then
txt1.Text = "Dale Carnegie"
txt2.Text = "Simon and Schuster"
txt3.Text = "Non-Fiction"
txt4.Text = "B116"
txt5.Text = "110"
End If

If ComboBox1.ListIndex = 12 Then
txt1.Text = "Steven Levy"
txt2.Text = "Anchor Press"
txt3.Text = "Computer Science"
txt4.Text = "C118"
txt5.Text = "200"
End If

If ComboBox1.ListIndex = 13 Then
txt1.Text = "Issac Asimov Gnome"
txt2.Text = "Gnome Press"
txt3.Text = "Technology"
txt4.Text = "F126"
txt5.Text = "379"
End If

If ComboBox1.ListIndex = 14 Then
txt1.Text = "Thomas H. Cormen"
txt2.Text = "MIT Press"
txt3.Text = "Computer Science"
txt4.Text = "C119"
txt5.Text = "700"
End If

If ComboBox1.ListIndex = 15 Then
txt1.Text = "Yann Martel"
txt2.Text = "Unknown"
txt3.Text = "Fiction"
txt4.Text = "A107"
txt5.Text = "140"
End If

If ComboBox1.ListIndex = 16 Then
txt1.Text = "Marcus Aurelius"
txt2.Text = "Unknown"
txt3.Text = "Philosopy and Psychology"
txt4.Text = "D121"
txt5.Text = "225"
End If

If ComboBox1.ListIndex = 17 Then
txt1.Text = "Victor Frankl"
txt2.Text = "Bencon Press"
txt3.Text = "Philosopy and Psychology"
txt4.Text = "D122"
txt5.Text = "114"
End If

If ComboBox1.ListIndex = 18 Then
txt1.Text = "Rick Riordan"
txt2.Text = "Miramax Books"
txt3.Text = "Non-Fiction"
txt4.Text = "A108"
txt5.Text = "254"
End If

If ComboBox1.ListIndex = 19 Then
txt1.Text = "Rick Riordan"
txt2.Text = "Miramax Books"
txt3.Text = "Non-Fiction"
txt4.Text = "A109"
txt5.Text = "350"
End If

If ComboBox1.ListIndex = 20 Then
txt1.Text = "Jane Austen"
txt2.Text = "Thomas Egerton"
txt3.Text = "Literature"
txt4.Text = "G127"
txt5.Text = "230"
End If

If ComboBox1.ListIndex = 21 Then
txt1.Text = "Suzanne Collins"
txt2.Text = "Scholastic Corporation"
txt3.Text = "Fiction"
txt4.Text = "A110"
txt5.Text = "499"
End If


If ComboBox1.ListIndex = 22 Then
txt1.Text = "Suzanne Collins"
txt2.Text = "Scholastic Corporation"
txt3.Text = "Fiction"
txt4.Text = "A111"
txt5.Text = "355"
End If


If ComboBox1.ListIndex = 23 Then
txt1.Text = "Tom Wolfe"
txt2.Text = "Starus"
txt3.Text = "Non-Fiction"
txt4.Text = "B117"
txt5.Text = "600"
End If


If ComboBox1.ListIndex = 24 Then
txt1.Text = "Tracy Kidder"
txt2.Text = "Little, Brown Company"
txt3.Text = "Computer-Science"
txt4.Text = "C120"
txt5.Text = "500"
End If

If ComboBox1.ListIndex = 25 Then
txt1.Text = "Charles Darwin"
txt2.Text = "Unknown"
txt3.Text = "Science"
txt4.Text = "E123"
txt5.Text = "200"
End If

If ComboBox1.ListIndex = 26 Then
txt1.Text = "James Watson"
txt2.Text = "Simon and Starus"
txt3.Text = "Science"
txt4.Text = "E124"
txt5.Text = "539"
End If


If ComboBox1.ListIndex = 27 Then
txt1.Text = "Primo Levi"
txt2.Text = "DK Children"
txt3.Text = "Science"
txt4.Text = "E125"
txt5.Text = "585"
End If
End Sub

Private Sub CommandButton2_Click()
ComboBox1.Text = " "
txt1.Text = " "
txt2.Text = " "
txt3.Text = " "
txt4.Text = " "
txt5.Text = " "
txt6.Text = " "
txt7.Text = " "
txt8.Text = " "
txt9.Text = " "
End Sub


Private Sub UserForm_Initialize()
With ComboBox1
.AddItem ("A Breif History of Time")
.AddItem ("Awakenings")
.AddItem ("Alice's Adventure in Wonderland")
.AddItem ("Harry Potter and The Philosopher's Stone")
.AddItem ("Harry Potter and The Chamber of Secret")
.AddItem ("Harry Potter and The Prisoner of Azkaban")
.AddItem ("Harry Potter and The Goblet of Fire")
.AddItem ("Harry Potter and The Order of the Phoenix")
.AddItem ("Harry Potter and The Half-Blood Prince")
.AddItem ("Harry Potter and The Deathly Hallows")
.AddItem ("Hiroshima")
.AddItem ("How to Win Friends and Influence People")
.AddItem ("Hackers")
.AddItem ("I, Robot")
.AddItem ("Introduction to Algorithms")
.AddItem ("Life of Pi")
.AddItem ("Meditation")
.AddItem ("Man's Search For Meaning")
.AddItem ("Percy Jackson and The Sea of Monsters")
.AddItem ("Percy Jackson and The Greek Gods")
.AddItem ("Pride and Prejudice")
.AddItem ("The Hunger Games")
.AddItem ("The Hunger Games: Catching Fire")
.AddItem ("The Right Stuff")
.AddItem ("The Soul of New Machine")
.AddItem ("The Origin of Species")
.AddItem ("The Double Helix")
.AddItem ("The Periodic Table")
End With
End Sub

