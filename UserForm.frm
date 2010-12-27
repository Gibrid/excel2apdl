VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm 
   Caption         =   "UserForm1"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9015.001
   OleObjectBlob   =   "UserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Button1_Click()
    'showUF1
    generalAction
End Sub


Private Sub CommandButtonClose_Click()
    Unload Me

End Sub

Private Sub CommandButtonCopyClip_Click()
    Dim outputData As New MSForms.DataObject
    outputData.SetText TextBox1.Value
    outputData.PutInClipboard

End Sub


Private Sub CommandButtonSaveFile_Click()
    Dim myFileDialog As FileDialog
                
    Set myFileDialog = Application.FileDialog(msoFileDialogSaveAs)
    With myFileDialog
    
    .InitialFileName = Cells(4, 9)
    .AllowMultiSelect = False
    
    '.Title = "Сохранить макрос"
    '.ButtonName = "Save"
    '.filters.Clear
    '.filters.Add "ANSYS APDL macros", "*.mac*"
    .FilterIndex = 18
    '.filters.Clear
    
    End With
    
    If myFileDialog.Show = -1 Then
        Call writeFile(myFileDialog.SelectedItems(1))
    End If
        
End Sub

Private Sub CommandButtonSetDir_Click()
    setDir
    Me.TextBoxPath.Value = Cells(4, 9)
End Sub

Private Sub UserForm_Initialize()
    Me.Caption = "excel2apdl converter (rough beta v0.92)"
    
    Me.TextBoxPath.Value = Cells(4, 9)
    Me.TextBox1.ScrollBars = fmScrollBarsVertical
    
End Sub
