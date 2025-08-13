Option Explicit

Public Sub ArrangeTokens_Simple()
    Dim rng As Range
    Dim cell As Range
    Dim uf As Object
    Dim arr As Variant

    On Error Resume Next
    Set rng = Application.InputBox("A列の範囲を選択してください", Type:=8)
    On Error GoTo 0
    If rng Is Nothing Then Exit Sub

    For Each cell In rng.Columns(1).Cells
        If Trim(CStr(cell.Value)) <> "" Then
            On Error Resume Next
            Set uf = VBA.UserForms.Add("UF_TokenArrange")
            If uf Is Nothing Then
                Set uf = VBA.UserForms.Add("UserForm1")
            End If
            On Error GoTo 0
            If Not uf Is Nothing Then
                arr = OneRowArray(cell.Row)
                CallByName uf, "InitBatch", VbMethod, cell.Worksheet.Name, arr, 0
                uf.Show vbModal
                Unload uf
                Set uf = Nothing
            End If
        End If
    Next cell

    MsgBox "完了しました"
End Sub

Public Function OneRowArray(ByVal rowNum As Long) As Variant
    Dim arr(0 To 0) As Long
    arr(0) = rowNum
    OneRowArray = arr
End Function
