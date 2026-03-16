Option Explicit

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    If Sh.Name <> "Student_Info" Then Exit Sub
    If Target.Column <> 10 Then Exit Sub
    If Target.Row < 7 Or Target.Row > 106 Then Exit Sub
    If Target.Cells.Count > 1 Then Exit Sub
    Application.EnableEvents = False
    Call ShowMultiSelect(Target)
    Application.EnableEvents = True
End Sub
