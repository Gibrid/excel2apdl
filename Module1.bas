Attribute VB_Name = "Module1"
Sub showUF()
    Load UserForm
    UserForm.Show
End Sub

Sub setDir()
    Dim myFileDialog As FileDialog
    Set myFileDialog = Application.FileDialog(msoFileDialogFolderPicker)
    
    myFileDialog.InitialFileName = Cells(4, 9)
    
    If myFileDialog.Show = -1 Then
        Cells(4, 9) = myFileDialog.SelectedItems(1)
    End If
    
End Sub


