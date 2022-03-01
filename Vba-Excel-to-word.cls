
Option Explicit

Sub excel_to_word()
Dim wrd As New Word.Application
Dim newdocument As Word.Document
Set newdocument = wrd.Documents.Add()

Dim abc As String
abc = InputBox(Prompt:="Þube adýný girin")

Dim xyz As String
xyz = InputBox(Prompt:="Müþteri adýný giriniz")

Dim myrng As excel.Range
Set myrng = Me.Range("A1:O55")

Dim nwrng As excel.Range
Dim d As Integer
For d = 2 To 55
Set nwrng = myrng.Range(Me.Cells(d, 1), Me.Cells(d, 15))
If abc = Me.Cells(d, 1) Then
If xyz = Me.Cells(d, 3) Then

nwrng.Copy
Else
End If

Else
End If
                                Hello world

Set newdocument = wrd.Documents.Add()
newdocument.Content.Paste
Next
        
        input ("very good")
newdocument.Content.InsertBefore Text:="this" & abc & xyz
newdocument.SaveAs2 Filename:=abc & "_" & xyz
newdocument.Close

End Sub

