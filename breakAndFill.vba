Private Sub CommandButton1_Click()

    Dim youneed As String
    
    '유저폼에서 입력한 값을 변수에 저장
    youneed = Me.TextBox1.value

    '프로시저 실행
    Module1.breaksandfill youneed

    '유저폼 종료
    Unload Me

End Sub

'sheet 등의 함수에 저장
Function IsFormLoaded(ByVal FormName As String) As Boolean
    Dim frm As Object
    For Each frm In VBA.UserForms
        If frm.Name = FormName Then
            IsFormLoaded = True
            Exit Function
        End If
    Next frm
    IsFormLoaded = False
End Function
Sub breaksandfill()
    '사용자 정의 폼 띄우기
    If Not IsFormLoaded("UserForm1") Then
        UserForm1.Show
    End If
End Sub



'sheet 등의 함수에 저장
Function IsFormLoaded(ByVal FormName As String) As Boolean
    Dim frm As Object
    For Each frm In VBA.UserForms
        If frm.Name = FormName Then
            IsFormLoaded = True
            Exit Function
        End If
    Next frm
    IsFormLoaded = False
End Function

Sub breaksandfill(youneed)

    Dim cell As Range

    '사용자 정의 폼 띄우기
    If Not IsFormLoaded("UserForm1") Then
        UserForm1.Show
    End If

    For Each cell In Selection.Cells
        Dim value As Variant
        value = cell.value
        If InStr(1, value, vbLf) > 0 Then
            value = Replace(value, vbLf, youneed)
            cell.value = value
        ElseIf InStr(1, value, vbCr) > 0 Then
            value = Replace(value, vbCr, youneed)
            cell.value = value
        End If
    Next cell

    For Each cell In Selection.Cells
        If cell.MergeCells Then
            Dim mergeRange As Range
            Set mergeRange = cell.MergeArea
            cell.UnMerge
            value = cell.value
            value = Replace(value, vbCrLf, youneed)
            mergeRange.value = value
        End If
    Next cell
    
End Sub
