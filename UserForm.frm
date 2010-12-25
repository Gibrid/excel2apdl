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

'Private Sub UserForm_Click()
    'Me.Height = Int(0.5 * 500)
    'Me.Width = Int(0.5 * 750)
    'showUF1
'End Sub

Private Sub Button1_Click()
    'showUF1
    generalAction
End Sub


Private Sub CommandButton2_Click()
    'loop through the items in the listbox
    For x = 0 To ListBox1.ListCount - 1
        'if the intem is selected
        If ListBox1.Selected(x) = True Then
            'display the selected item
            MsgBox ListBox1.List(x)
        End If
    Next x
End Sub

Private Sub CommandButton3_Click()
    PopulateListBox
End Sub

Private Sub CommandButtonClose_Click()
    Unload Me

End Sub

Private Sub CommandButtonCopyClip_Click()
    Dim outputData As New MSForms.DataObject
    outputData.SetText TextBox1.Value
    outputData.PutInClipboard

End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ListBox1.RowSource = "=Лист1!A1:A5"
End Sub

Private Sub CommandButtonSaveFile_Click()
    Dim myFileDialog As FileDialog
    
    Set myFileDialog = Application.FileDialog(msoFileDialogSaveAs)
    If myFileDialog.Show = -1 Then
        MsgBox myFileDialog.Item
    '    Call writeFile(myFileDialog.InitialFileName)
        
    End If
        
End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub UserForm_Initialize()
    Me.Caption = "excel2apdl converter (rough beta v0.91)"
    'Me.BackColor = RGB(10, 25, 100)
    
    Me.TextBox1.ScrollBars = fmScrollBarsVertical
    
End Sub
    

Private Sub UserForm_Resize()
    msg = "Width: " & Me.Width & Chr(10) & "Height: " & Me.Height
    MsgBox prompt:=msg, Title:="Resize Event"
End Sub

'Private Sub UserForm_Terminate()
    'msg = "Now Unloading " & Me.Caption
    'MsgBox prompt:=msg, Title:="TerminateEvent"
'End Sub
