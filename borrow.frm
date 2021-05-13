VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} borrow 
   Caption         =   "Borrow Book"
   ClientHeight    =   10515
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18585
   OleObjectBlob   =   "borrow.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "borrow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton2_Click()
txt1.Text = " "
txt2.Text = " "
txt3.Text = " "
txt4.Text = " "
ComboBox1.Text = " "
ComboBox2.Text = " "
ComboBox3.Text = " "
ComboBox4.Text = " "
ComboBox5.Text = " "
ComboBox6.Text = " "
ComboBox7.Text = " "
ComboBox8.Text = " "
ComboBox9.Text = " "
ComboBox10.Text = " "
ComboBox11.Text = " "
ComboBox12.Text = " "
ComboBox13.Text = " "
ComboBox14.Text = " "
ComboBox15.Text = " "

End Sub

Private Sub CommandButton3_Click()
Sheet2.Activate
emptyRow = WorksheetFunction.CountA(Range("A:A"))
Range("A1").Value = "Member Name"
Range("B1").Value = "Book No."
Range("C1").Value = "Book Name"
Range("D1").Value = "Book Author"
Range("E1").Value = "Return Date"
Range("F1").Value = "Issued Date"
Range("G1").Value = "Book No."
Range("H1").Value = "Book Name"
Range("I1").Value = "Book Author"
Range("J1").Value = "Issued Date"
Range("K1").Value = "Return Date"


Sheet2.Activate
emptyRow = WorksheetFunction.CountA(Range("A:A")) + 1
Cells(emptyRow, 1).Value = ComboBox15.Value
Cells(emptyRow, 2).Value = ComboBox1.Value
Cells(emptyRow, 3).Value = txt1.Value
Cells(emptyRow, 4).Value = txt2.Value
Cells(emptyRow, 5).Value = txt_return.Value
Cells(emptyRow, 6).Value = txt_issue.Value
Cells(emptyRow, 7).Value = ComboBox8.Value
Cells(emptyRow, 8).Value = txt3.Value
Cells(emptyRow, 9).Value = txt4.Value
Cells(emptyRow, 10).Value = txt_issue2.Value
Cells(emptyRow, 11).Value = txt_return2.Value
Done3.Show

End Sub

Private Sub ComboBox1_Change()
If ComboBox1.ListIndex = 0 Then
txt1.Text = "Harry Potter and The Philospher's Stone"
txt2.Text = "J.K.Rowling"
End If

If ComboBox1.ListIndex = 1 Then
txt1.Text = "Harry Potter and The Chamber of Secrets. "
txt2.Text = "J.K.Rowling"
End If

If ComboBox1.ListIndex = 2 Then
txt1.Text = "Harry Potter and The Prisonner of Azkaban"
txt2.Text = "J.K.Rowling"
End If

If ComboBox1.ListIndex = 3 Then
txt1.Text = "Harry Potter and The Goblet of Fire"
txt2.Text = "J.K.Rowling"
End If

If ComboBox1.ListIndex = 4 Then
txt1.Text = "Harry Potter and The Order of Phoennix"
txt2.Text = "J.K.Rowling"
End If

If ComboBox1.ListIndex = 5 Then
txt1.Text = "Harry Potter and The Prince"
txt2.Text = "J.K.Rowling"
End If

If ComboBox1.ListIndex = 6 Then
txt1.Text = "Harry Potter and The Hsllows"
txt2.Text = "J.K.Rowling"
End If

If ComboBox1.ListIndex = 7 Then
txt1.Text = "Life of Pi"
txt2.Text = "Yann Martel"
End If

If ComboBox1.ListIndex = 8 Then
txt1.Text = "Percy jackson and The Sea of monsters"
txt2.Text = "Rick Riordon"
End If

If ComboBox1.ListIndex = 9 Then
txt1.Text = "Percy Jackson and The Greek Gods"
txt2.Text = "Rick Riordon"
End If

If ComboBox1.ListIndex = 10 Then
txt1.Text = "The Hunger Game"
txt2.Text = "Suzanne Collins"
End If

If ComboBox1.ListIndex = 11 Then
txt1.Text = "The Hunger Game: Catching Fire"
txt2.Text = "Suzanne Collins"
End If


If ComboBox1.ListIndex = 12 Then
txt1.Text = "A Breif History of Time"
txt2.Text = "Stephen Hawking"
End If

If ComboBox1.ListIndex = 13 Then
txt1.Text = "Awakenings"
txt2.Text = "Oliver Sacks"
End If

If ComboBox1.ListIndex = 14 Then
txt1.Text = "Hiroshima"
txt2.Text = "John Hersey"
End If

If ComboBox1.ListIndex = 15 Then
txt1.Text = "How to Win Friends and Influence People"
txt2.Text = " dale Carnegie"
End If

If ComboBox1.ListIndex = 16 Then
txt1.Text = "The Right Stuft"
txt2.Text = "Tom Wolfe"
End If

If ComboBox1.ListIndex = 17 Then
txt1.Text = "Hackers"
txt2.Text = "Steven Levy"
End If

If ComboBox1.ListIndex = 18 Then
txt1.Text = "Introduction to Algorithms"
txt2.Text = "Thomas H.Cormen"
End If

If ComboBox1.ListIndex = 19 Then
txt1.Text = "Meditation"
txt2.Text = "Marcus Aurelius"
End If

If ComboBox1.ListIndex = 20 Then
txt1.Text = "Maris Search For Meanning"
txt2.Text = "Victor Frankl"
End If


If ComboBox1.ListIndex = 21 Then
txt1.Text = "The Origin of Species"
txt2.Text = "Charles Darwin"
End If


If ComboBox1.ListIndex = 22 Then
txt1.Text = "The Double Felix"
txt2.Text = "James Watson"
End If


If ComboBox1.ListIndex = 23 Then
txt1.Text = "The Periodic Table"
txt2.Text = "Primo Levi"
End If

If ComboBox1.ListIndex = 24 Then
txt1.Text = "I,Robot"
txt2.Text = "Issac Asimov Gnome"
End If

If ComboBox1.ListIndex = 25 Then
txt1.Text = "Pride and Prejudice"
txt2.Text = "Jane Austen"
End If


If ComboBox1.ListIndex = 26 Then
txt1.Text = "Alice's Adventvre in Wonderand "
txt2.Text = "Lewis Carroll"
End If
End Sub

Private Sub ComboBox8_Change()
If ComboBox8.ListIndex = 0 Then
txt3.Text = "Harry Potter and The Philospher's Stone"
txt4.Text = "J.K.Rowling"
End If

If ComboBox8.ListIndex = 1 Then
txt3.Text = "Harry Potter and The Chamber of Secrets. "
txt4.Text = "J.K.Rowling"
End If

If ComboBox8.ListIndex = 2 Then
txt3.Text = "Harry Potter and The Prisonner of Azkaban"
txt4.Text = "J.K.Rowling"
End If

If ComboBox8.ListIndex = 3 Then
txt3.Text = "Harry Potter and The Goblet of Fire"
txt4.Text = "J.K.Rowling"
End If

If ComboBox8.ListIndex = 4 Then
txt3.Text = "Harry Potter and The Order of Phoennix"
txt4.Text = "J.K.Rowling"
End If

If ComboBox8.ListIndex = 5 Then
txt3.Text = "Harry Potter and The Prince"
txt4.Text = "J.K.Rowling"
End If

If ComboBox8.ListIndex = 6 Then
txt3.Text = "Harry Potter and The Hsllows"
txt4.Text = "J.K.Rowling"
End If

If ComboBox8.ListIndex = 7 Then
txt3.Text = "Life of Pi"
txt4.Text = "Yann Martel"
End If

If ComboBox8.ListIndex = 8 Then
txt3.Text = "Percy jackson and The Sea of monsters"
txt4.Text = "Rick Riordon"
End If

If ComboBox8.ListIndex = 9 Then
txt3.Text = "Percy Jackson and The Greek Gods"
txt4.Text = "Rick Riordon"
End If

If ComboBox8.ListIndex = 10 Then
txt3.Text = "The Hunger Game"
txt4.Text = "Suzanne Collins"
End If

If ComboBox8.ListIndex = 11 Then
txt3.Text = "The Hunger Game: Catching Fire"
txt4.Text = "Suzanne Collins"
End If


If ComboBox8.ListIndex = 12 Then
txt3.Text = "A Breif History of Time"
txt4.Text = "Stephen Hawking"
End If

If ComboBox8.ListIndex = 13 Then
txt3.Text = "Awakenings"
txt4.Text = "Oliver Sacks"
End If

If ComboBox8.ListIndex = 14 Then
txt3.Text = "Hiroshima"
txt4.Text = "John Hersey"
End If

If ComboBox8.ListIndex = 15 Then
txt3.Text = "How to Win Friends and Influence People"
txt4.Text = " dale Carnegie"
End If

If ComboBox8.ListIndex = 16 Then
txt3.Text = "The Right Stuft"
txt4.Text = "Tom Wolfe"
End If

If ComboBox8.ListIndex = 17 Then
txt3.Text = "Hackers"
txt4.Text = "Steven Levy"
End If

If ComboBox8.ListIndex = 18 Then
txt3.Text = "Introduction to Algorithms"
txt4.Text = "Thomas H.Cormen"
End If

If ComboBox8.ListIndex = 19 Then
txt3.Text = "Meditation"
txt4.Text = "Marcus Aurelius"
End If

If ComboBox8.ListIndex = 20 Then
txt3.Text = "Maris Search For Meanning"
txt4.Text = "Victor Frankl"
End If


If ComboBox8.ListIndex = 21 Then
txt3.Text = "The Origin of Species"
txt4.Text = "Charles Darwin"
End If


If ComboBox8.ListIndex = 22 Then
txt3.Text = "The Double Felix"
txt4.Text = "James Watson"
End If


If ComboBox8.ListIndex = 23 Then
txt3.Text = "The Periodic Table"
txt4.Text = "Primo Levi"
End If

If ComboBox8.ListIndex = 24 Then
txt3.Text = "I,Robot"
txt4.Text = "Issac Asimov Gnome"
End If

If ComboBox8.ListIndex = 25 Then
txt3.Text = "Pride and Prejudice"
txt4.Text = "Jane Austen"
End If


If ComboBox8.ListIndex = 26 Then
txt3.Text = "Alice's Adventvre in Wonderand "
txt4.Text = "Lewis Carroll"
End If
End Sub

Private Sub comd_cancel_Click()
Unload Me
End Sub


Private Sub UserForm_Initialize()
With ComboBox1
.AddItem ("A100")
.AddItem ("A101")
.AddItem ("A102")
.AddItem ("A103")
.AddItem ("A104")
.AddItem ("A105")
.AddItem ("A106")
.AddItem ("A107")
.AddItem ("A108")
.AddItem ("A109")
.AddItem ("A110")
.AddItem ("A111")
.AddItem ("B113")
.AddItem ("B114")
.AddItem ("B115")
.AddItem ("B116")
.AddItem ("B117")
.AddItem ("C118")
.AddItem ("C119")
.AddItem ("C120")
.AddItem ("D121")
.AddItem ("D122")
.AddItem ("E123")
.AddItem ("E124")
.AddItem ("E125")
.AddItem ("F126")
.AddItem ("G127")
.AddItem ("G128")
End With

With ComboBox8
.AddItem ("A100")
.AddItem ("A101")
.AddItem ("A102")
.AddItem ("A103")
.AddItem ("A104")
.AddItem ("A105")
.AddItem ("A106")
.AddItem ("A107")
.AddItem ("A108")
.AddItem ("A109")
.AddItem ("A110")
.AddItem ("A111")
.AddItem ("B113")
.AddItem ("B114")
.AddItem ("B115")
.AddItem ("B116")
.AddItem ("B117")
.AddItem ("C118")
.AddItem ("C119")
.AddItem ("C120")
.AddItem ("D121")
.AddItem ("D122")
.AddItem ("E123")
.AddItem ("E124")
.AddItem ("E125")
.AddItem ("F126")
.AddItem ("G127")
.AddItem ("G128")
End With


txt_return.Text = Date

txt_issue.Text = Date - 7

txt_return2.Text = Date + 7

txt_issue2.Text = Date

End Sub




