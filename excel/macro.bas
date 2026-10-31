' 貼り付け先: 「ボード」シートのコードウィンドウ（シート名を右クリック→コードの表示）
Option Explicit

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim c As Range, st As Range

    If Target.Cells.Count <> 1 Then Exit Sub
    Set c = Target.Cells(1)

    ' I1 = 運用中／編集中 の切り替え
    If c.Address = "$I$1" Then
        If c.Value = "編集中" Then c.Value = "運用中" Else c.Value = "編集中"
        Call ResetSelection
        Exit Sub
    End If

    ' 編集中は何もしない（名前の変更やセルの移動ができる）
    If Me.Range("I1").Value = "編集中" Then Exit Sub

    ' 名前セル（1つ下が状態セル）または状態セルをタッチしたら切り替え
    If c.Value = "出勤" Or c.Value = "退勤" Then
        Set st = c
    ElseIf c.Offset(1, 0).Value = "出勤" Or c.Offset(1, 0).Value = "退勤" Then
        Set st = c.Offset(1, 0)
    Else
        Exit Sub
    End If

    Application.EnableEvents = False
    If st.Value = "出勤" Then st.Value = "退勤" Else st.Value = "出勤"
    Application.EnableEvents = True

    ' 同じ場所をもう一度タッチしても反応するよう、選択を余白セルに逃がす
    Call ResetSelection
End Sub

Private Sub ResetSelection()
    Application.EnableEvents = False
    Me.Range("A1").Select
    Application.EnableEvents = True
End Sub
