Sub select_result()
'
' select_result Macro
'
    Range("H1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
End Sub

Sub NameSplit()

Dim txt As String
Dim i As Integer
Dim FullName As Variant
Dim x As String, cell As Range

For Each cell In ActiveSheet.Range("A1:A20")
     txt = cell.Value

     FullName = Split(txt, " ")

     For i = 0 To UBound(FullName)

            If i = 0 Then
                cell.Offset(0, 1).Value = FullName(0)
            ElseIf i = 1 Then
                cell.Offset(0, 7).Value = FullName(i)
            ElseIf i = 2 Then
                cell.Offset(0, 8).Value = FullName(i)
            ElseIf i = 3 Then
                cell.Offset(0, 9).Value = "AM"
                cell.Offset(0, 10).Value = FullName(i)
            End If
     Next i

Next cell

End Sub
Sub SlashSplit()

Dim txt As String
Dim i As Integer
Dim FullName As Variant
Dim x As String, cell As Range

For Each cell In ActiveSheet.Range("B1:B20")
     txt_for_slash = cell.Value
     
     FullName_slash = Split(txt_for_slash, "\")
        max_Size_slash = UBound(FullName_slash)
        path_ = ""
        For Z = 0 To max_Size_slash
            If Z = max_Size_slash Then
                cell.Offset(0, 11).Value = path_
                cell.Offset(0, 10).Value = FullName_slash(Z)
                path_ = ""
            Else
                path_ = path_ + FullName_slash(Z) + "\"
            End If
        Next Z


Next cell

End Sub
Sub RemoveSpaces_3()
'Remove multiple spaces from a range
Dim r1 As Range
Set r1 = ActiveSheet.Range("A1:A20") 'change this line to match your range
r1.Replace _
      What:=Space(2), _
      Replacement:=" ", _
      SearchOrder:=xlByColumns, _
      MatchCase:=True
Set r1 = r1.Find(What:=Space(2))
If Not r1 Is Nothing Then
   Call RemoveSpaces_3
End If
End Sub
Sub date_set()
    For Each cell In ActiveSheet.Range("H1:H20")
     Content = cell.Value
     If Content <> "" Then
        first = Mid(Content, 5, 2)
        center = Mid(Content, 7, 2)
        last = Mid(Content, 1, 4)
        cell.Value = first + "/" + center + "/" + last
     End If
    Next cell
End Sub
Sub main_macro()
Call RemoveSpaces_3
Call NameSplit
Call SlashSplit
Call set_col_number
Call date_set
Call select_result
'Call delete_col
End Sub
Sub set_col_number()
    Columns("H:H").Select
    Selection.NumberFormat = "@"
    Columns("K:K").Select
    Selection.NumberFormat = "@"
End Sub
Sub delete_col()
    Selection.Delete Shift:=xlToLeft
End Sub


Private Sub CommandButton1_Click()
Call main_macro
End Sub
