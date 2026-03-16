Option Explicit

Sub ShowMultiSelect(ByVal cell As Range)
    Dim subjects As Variant
    subjects = Array("Maths", "Science", "Social Science", "Biology", "Physics", "Chemistry", "Computer Science")
    
    Dim frmName As String
    frmName = "frmSubjectPicker"
    
    ' Remove existing form if any
    On Error Resume Next
    ThisWorkbook.VBProject.VBComponents.Remove _
        ThisWorkbook.VBProject.VBComponents(frmName)
    On Error GoTo 0
    
    ' Create UserForm
    Dim frm As Object
    Set frm = ThisWorkbook.VBProject.VBComponents.Add(3)
    frm.Name = frmName
    With frm.Properties
        .Item("Caption") = "Select Subjects — Row " & cell.Row
        .Item("Width")   = 230
        .Item("Height")  = 310
        .Item("StartUpPosition") = 1
        .Item("BackColor") = RGB(238, 242, 249)
    End With
    
    ' Title label
    Dim lbl As Object
    Set lbl = frm.Designer.Controls.Add("Forms.Label.1")
    With lbl
        .Caption = "Select one or more subjects:"
        .Left = 10: .Top = 8: .Width = 200: .Height = 16
        .Font.Bold = True: .Font.Size = 9
        .ForeColor = RGB(26, 60, 110)
    End With
    
    ' ListBox
    Dim lb As Object
    Set lb = frm.Designer.Controls.Add("Forms.ListBox.1")
    lb.Name = "lstSubjects"
    lb.MultiSelect  = 1
    lb.Left = 10: lb.Top = 28: lb.Width = 200: lb.Height = 200
    lb.Font.Size = 10
    
    ' OK button
    Dim btnOK As Object
    Set btnOK = frm.Designer.Controls.Add("Forms.CommandButton.1")
    btnOK.Name = "btnOK"
    btnOK.Caption = "OK": btnOK.Left = 10: btnOK.Top = 238
    btnOK.Width = 95: btnOK.Height = 28
    btnOK.BackColor = RGB(26, 60, 110)
    btnOK.ForeColor = RGB(255, 255, 255)
    
    ' Cancel button
    Dim btnCancel As Object
    Set btnCancel = frm.Designer.Controls.Add("Forms.CommandButton.1")
    btnCancel.Name = "btnCancel"
    btnCancel.Caption = "Cancel": btnCancel.Left = 115: btnCancel.Top = 238
    btnCancel.Width = 95: btnCancel.Height = 28
    
    ' Current selections
    Dim current As String
    current = Trim(cell.Value)
    
    ' Add items and pre-select existing
    Dim i As Integer
    For i = 0 To UBound(subjects)
        lb.AddItem subjects(i)
        If Len(current) > 0 Then
            If InStr(1, "," & current & ",", "," & Trim(subjects(i)) & ",", vbTextCompare) > 0 Then
                lb.Selected(i) = True
            End If
        End If
    Next i
    
    ' Form code
    Dim code As String
    code = "Private Sub btnOK_Click()" & Chr(13) & Chr(10) & _
           "    Dim res As String: Dim i As Integer" & Chr(13) & Chr(10) & _
           "    For i = 0 To lstSubjects.ListCount - 1" & Chr(13) & Chr(10) & _
           "        If lstSubjects.Selected(i) Then" & Chr(13) & Chr(10) & _
           "            If Len(res) > 0 Then res = res & "", """ & Chr(13) & Chr(10) & _
           "            res = res & lstSubjects.List(i)" & Chr(13) & Chr(10) & _
           "        End If" & Chr(13) & Chr(10) & _
           "    Next i" & Chr(13) & Chr(10) & _
           "    Me.Tag = res: Me.Hide" & Chr(13) & Chr(10) & _
           "End Sub" & Chr(13) & Chr(10) & _
           "Private Sub btnCancel_Click()" & Chr(13) & Chr(10) & _
           "    Me.Tag = ""CANCEL"": Me.Hide" & Chr(13) & Chr(10) & _
           "End Sub"
    frm.CodeModule.AddFromString code
    
    ' Show form
    Dim result As String
    VBA.UserForms.Add(frm.Name).Show
    result = VBA.UserForms(VBA.UserForms.Count - 1).Tag
    Unload VBA.UserForms(VBA.UserForms.Count - 1)
    
    ' Clean up
    ThisWorkbook.VBProject.VBComponents.Remove _
        ThisWorkbook.VBProject.VBComponents(frmName)
    
    If result <> "CANCEL" Then
        cell.Value = result
    End If
End Sub
