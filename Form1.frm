Dim C As New Connection
Dim R As New Recordset
Dim S As String

Private Sub cmdAdd_Click()
    txtRno.Text = ""
    txtSname.Text = ""
    txtClass.Text = ""
    txtDOB.Text = ""
    txtRno.SetFocus
End Sub

Private Sub cmdNext_Click()
    R.MoveNext
    If Not R.EOF Then
        txtRno.Text = R.Fields(0).Value
        txtSname.Text = R.Fields(1).Value
        txtClass.Text = R.Fields(2).Value
        txtDOB.Text = R.Fields(3).Value
    Else
        MsgBox "No More Records!", vbInformation, "Student"
    End If
End Sub

Private Sub cmdPrev_Click()
    R.MovePrevious
    If Not R.BOF Then
        txtRno.Text = R.Fields(0).Value
        txtSname.Text = R.Fields(1).Value
        txtClass.Text = R.Fields(2).Value
        txtDOB.Text = R.Fields(3).Value
    Else
        MsgBox "No More Records!", vbInformation, "Student"
    End If
End Sub

Private Sub cmdSave_Click()
    R.Close
    S = "Insert Into studData Values(" & Val(txtRno.Text) & ",'" & txtSname.Text & "','" & txtClass.Text & "','" & txtDOB.Text & "')"
    R.Open S, C, adOpenDynamic, adLockOptimistic

    S = "Select * From studData"
    R.Open S, C, adOpenDynamic, adLockOptimistic

    If Not R.BOF And Not R.EOF Then
        R.MoveFirst
        txtRno.Text = R.Fields(0).Value
        txtSname.Text = R.Fields(1).Value
        txtClass.Text = R.Fields(2).Value
        txtDOB.Text = R.Fields(3).Value
    End If
    MsgBox "Student record Added Successfully!", vbInformation, "Student"
End Sub

Private Sub Form_Load()
    S = "Select * From studData"
    C.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\VBSlipSol\Slip07\Ques2\stud.mdb;Persist Security Info=False"
    R.Open S, C, adOpenDynamic, adLockOptimistic

    If Not R.BOF And Not R.EOF Then
        R.MoveFirst
        txtRno.Text = R.Fields(0).Value
    End If
End Sub
