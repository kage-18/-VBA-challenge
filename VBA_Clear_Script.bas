Attribute VB_Name = "Module2"
Sub Button2_Click()
For i = 1 To Worksheets.Count
Dim ws As Worksheet
Set ws = Worksheets(i)
ws.Range("I1:M1000000").ClearContents
ws.Range("N1:P10").ClearContents
Next i
End Sub
