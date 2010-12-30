Attribute VB_Name = "Module2"
    Dim lineCount As Integer
    
    Dim strOutput As String
    
    Dim alphaPart As String
    Dim tempPart As String
    Dim coordPart As String
    
    Dim alphaArrayName As String
    Dim tempArrayName As String
    Dim nsName As String
    Dim csysName As String
    
    Dim entityX As String, rangeX As Integer
    Dim entityY As String, rangeY As Integer
    Dim entityZ As String, rangeZ As Integer
    
    Dim coordRow As Integer
    

Sub getLineCount()
    Range("E4").Select
            If IsEmpty(ActiveCell) Then Exit Sub
top:
            lineCount = lineCount + 1
            ActiveCell.Offset(1, 0).Select
            If Not IsEmpty(ActiveCell) Then GoTo top
    UserForm.TextBox2.Text = lineCount
End Sub

Public Sub coordsCheck()
    'MsgBox "coordsCheck"
    Select Case UserForm.CheckBoxX.Value
        Case True
            entityX = "X"
            rangeX = lineCount
        Case False
            entityX = ""
            rangeX = 1
    End Select
    
    Select Case UserForm.CheckBoxY.Value
        Case True
            entityY = "Y"
            rangeY = lineCount
        Case False
            entityY = ""
            rangeY = 1
    End Select
    
    Select Case UserForm.CheckBoxZ.Value
        Case True
            entityZ = "Z"
            rangeZ = lineCount
        Case False
            entityZ = ""
            rangeZ = 1
    End Select
    
    ' x y z selected
    If UserForm.CheckBoxX.Value = True And _
        UserForm.CheckBoxY.Value = True And _
        UserForm.CheckBoxZ.Value = True Then
        
        coordPart = coordPart _
        & "," & rangeX _
        & "," & rangeY _
        & "," & rangeZ _
        & "," & entityX _
        & "," & entityY _
        & "," & entityZ
        
    ' y z selected
    ElseIf UserForm.CheckBoxX.Value = False And _
        UserForm.CheckBoxY.Value = True And _
        UserForm.CheckBoxZ.Value = True Then
        
        coordPart = coordPart _
        & "," & rangeY _
        & "," & rangeZ _
        & "," & rangeX _
        & "," & entityY _
        & "," & entityZ _
        & "," & entityX
        
    ' x z selected
    ElseIf UserForm.CheckBoxX.Value = True And _
        UserForm.CheckBoxY.Value = False And _
        UserForm.CheckBoxZ.Value = True Then
        
        coordPart = coordPart _
        & "," & rangeX _
        & "," & rangeZ _
        & "," & rangeY _
        & "," & entityX _
        & "," & entityZ _
        & "," & entityY

    ' x y selected
    ElseIf UserForm.CheckBoxX.Value = True And _
        UserForm.CheckBoxY.Value = True And _
        UserForm.CheckBoxZ.Value = False Then
        
        coordPart = coordPart _
        & "," & rangeX _
        & "," & rangeY _
        & "," & rangeZ _
        & "," & entityX _
        & "," & entityY _
        & "," & entityZ

    ' x selected
    ElseIf UserForm.CheckBoxX.Value = True And _
        UserForm.CheckBoxY.Value = False And _
        UserForm.CheckBoxZ.Value = False Then
        
        coordPart = coordPart _
        & "," & rangeX _
        & "," & rangeY _
        & "," & rangeZ _
        & "," & entityX _
        & "," & entityY _
        & "," & entityZ
        
        coordRow = 2
        
    ' y selected
    ElseIf UserForm.CheckBoxX.Value = False And _
        UserForm.CheckBoxY.Value = True And _
        UserForm.CheckBoxZ.Value = False Then
        
        coordPart = coordPart _
        & "," & rangeY _
        & "," & rangeX _
        & "," & rangeZ _
        & "," & entityY _
        & "," & entityX _
        & "," & entityZ
        
        coordRow = 3

    ' z selected
    ElseIf UserForm.CheckBoxX.Value = False And _
        UserForm.CheckBoxY.Value = False And _
        UserForm.CheckBoxZ.Value = True Then
        
        coordPart = coordPart _
        & "," & rangeZ _
        & "," & rangeX _
        & "," & rangeY _
        & "," & entityZ _
        & "," & entityX _
        & "," & entityY
        
        coordRow = 4
        
    End If
       
End Sub

Sub alphaLoop()
    alphaPart = ""
    For i = 1 To lineCount
        
        alphaPart = alphaPart & "*SET," _
        & alphaArrayName & "(" _
        & i & ",0)," _
        & Replace(Cells(i + 3, coordRow), ",", ".") & vbLf
    Next i
    
    alphaPart = alphaPart & vbLf
    For i = 1 To lineCount
        
        alphaPart = alphaPart & "*SET," _
        & alphaArrayName & "(" _
        & i & ",1)," _
        & Replace(Cells(i + 3, 5), ",", ".") & vbLf
    Next i
    
End Sub

Sub tempLoop()
    tempPart = ""
    For i = 1 To lineCount
        
        tempPart = tempPart & "*SET," _
        & tempArrayName & "(" _
        & i & ",0)," _
        & Replace(Cells(i + 3, coordRow), ",", ".") & vbLf
    Next i
    
    tempPart = tempPart & vbLf
    For i = 1 To lineCount
        
        tempPart = tempPart & "*SET," _
        & tempArrayName & "(" _
        & i & ",1)," _
        & Replace(Cells(i + 3, 6), ",", ".") & vbLf
    Next i
    
End Sub

Sub guTypeCheck()
    Select Case UserForm.CheckBoxAlpha.Value
        Case True
            alphaLoop
        Case False
    End Select
    
    Select Case UserForm.CheckBoxTemp.Value
        Case True
            tempLoop
        Case False
    End Select
    
End Sub

Sub generalAction()
    
    lineCount = 0
    
    alphaPart = ""
    tempPart = ""
    coordPart = ""
    
    alphaArrayName = Range("E2").Value
    tempArrayName = Range("F2").Value
    csysName = Range("G4").Value
    nsName = Range("H4").Value
    
    getLineCount
    coordsCheck
    guTypeCheck
    
    strOutput = "/prep7" & vbLf & vbLf
    strOutput = strOutput & "*DIM," & alphaArrayName _
        & ",TABLE" _
        & coordPart _
        & "," & csysName & vbLf
    
    strOutput = strOutput & alphaPart & vbLf & vbLf
    
    strOutput = strOutput & "*DIM," & tempArrayName _
        & ",TABLE" _
        & coordPart _
        & "," & csysName & vbLf
    
    strOutput = strOutput & tempPart & vbLf
    
    
    strOutput = strOutput & "cmsel,s," & nsName
    
    strOutput = strOutput & vbLf & vbLf & "SF,all," _
    & "CONV," _
    & "%" & alphaArrayName & "%," _
    & "%" & tempArrayName & "%" _
    & vbLf
    
    strOutput = strOutput & "allsel,all,all"
    
    strOutput = strOutput & vbLf & vbLf & "/solu"
    
    UserForm.TextBox1.Text = strOutput
    
    UserForm.TextBox1.SetFocus
    
    'MsgBox CDbl(Cells(4, 2))
    'MsgBox UserForm.TextBox1.Value
    
End Sub

Sub writeFile(ByRef filename)
    'MsgBox filename
    'MsgBox c
    f = FreeFile
    Open Replace(filename, ".txt", ".mac") For Output As #f
    Print #f, UserForm.TextBox1.Value
    Close #f
    'MsgBox filename
    

End Sub

