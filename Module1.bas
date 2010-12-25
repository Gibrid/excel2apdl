Attribute VB_Name = "Module1"
Sub showMessage()
MsgBox "Hello World!"
End Sub

Sub PopulateListBox()

    Dim MyArray As Variant
    Dim Ctr As Integer
    MyArray = Array("apples", "lolo", "lala", "kkaka")
    
    For Ctr = LBound(MyArray) To UBound(MyArray)
        UserForm.ListBox1.AddItem MyArray(Ctr)
    Next
    
End Sub

Sub ShowTime()
    Range("C1") = Now()
End Sub

Sub for_each_demo()
    For Each Cell In Selection
        MsgBox Cell.Value
    Next
End Sub

Sub showUF()
    Load UserForm
    UserForm.Show
End Sub


