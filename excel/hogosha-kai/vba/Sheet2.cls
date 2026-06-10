Attribute VB_Name = "Sheet2"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

'==============================================================
' 希望入力シート（調整シートはこのコードごとコピーされる）
'   ・入力モード中の行ハイライト
'   ・右クリックでマーク消去
'   ・入れ替えモード（ChangeColor）での確定操作
'==============================================================

Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
    Dim c As Range
    Dim m As Long
    If UserForm3.Visible = False Then Exit Sub
    EnsureDims
    With Selection
        If .Row <= Range("FIRST").Row Or .Row > ss + Range("FIRST").Row Or _
            .Column < Range("FIRST").Column Or .Column > Range("FIRST").Column + kk * dd - 1 Then Exit Sub
        For Each c In Selection
            If CStr(c.Value) <> "" Then
                m = MsgBox("選択セルを消去しますか。", vbOKCancel)
                If m = vbOK Then .Value = ""
                Cancel = True
                Exit Sub
            End If
        Next
        Cancel = True
    End With
End Sub

Private Sub Worksheet_Deactivate()
    If UserForm3.Visible = True And ActiveSheet.Name <> nowSh Then
        MsgBox "シートを切り替えるにはメニューを閉じてください。"
        Sheets(nowSh).Select
    End If
    On Error Resume Next
    Me.Range("BL3:BL52").ClearContents
    ChoseiSagyo = ""
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim r0 As Long, c0 As Long
    Dim r As Long, c As Long
    Dim ownCol As Long, otherRow As Long
    Dim prevC As Variant, prevC2 As Variant
    Dim m As Long

    If UserForm3.Visible = True Then
        '--- 入力モード: 選択行のハイライト --------------------
        If Cells(Range("RowNow").Row - 1, Range("RowNow").Column) = 1 Or _
            Cells(Target.Row, Range("RowNow").Column) = 1 Then Exit Sub
        Range("RowNow").ClearContents
        EnsureDims
        If Target.Row <= Range("FIRST").Row Or Target.Row > Range("FIRST").Row + ss Then Exit Sub
        Range("RowNow").Cells(Target.Row - Range("FIRST").Row, 1) = 1

    ElseIf ChoseiSagyo = "ChangeColor" Then
        '--- 入れ替えモード ------------------------------------
        EnsureDims
        If ss = 0 Then Exit Sub
        r0 = Range("FIRST").Row
        c0 = Range("FIRST").Column
        With Selection
            If .Row <= r0 Or .Row > ss + r0 Or _
                .Column < c0 Or .Column > c0 + kk * dd - 1 Then Exit Sub
            If .Cells.Count > 1 Then Exit Sub
            If .Interior.ColorIndex = xlNone Or .Interior.ColorIndex = kColor Then
                ActiveSheet.Unprotect

                '自分（同じ行）の確定コマを捜す
                ownCol = 0
                For c = c0 To c0 + dd * kk - 1
                    If Cells(.Row, c).Interior.ColorIndex = mColor Or _
                       Cells(.Row, c).Interior.ColorIndex = pColor Then
                        ownCol = c
                        Exit For
                    End If
                Next
                '相手（同じ列）の確定コマを捜す
                otherRow = 0
                For r = r0 + 1 To r0 + ss
                    If Cells(r, .Column).Interior.ColorIndex = mColor Or _
                       Cells(r, .Column).Interior.ColorIndex = pColor Then
                        otherRow = r
                        Exit For
                    End If
                Next

                If ownCol = 0 And otherRow > 0 Then
                    '自分の確定が無いのに枠は他の生徒が使用中 → 不可
                    prevC2 = Cells(otherRow, .Column).Interior.ColorIndex
                    Cells(otherRow, .Column).Interior.ColorIndex = dColor
                    MsgBox "この枠は別の生徒が確定済みです。" & vbLf & _
                        "（入れ替えは、自分の確定コマがある生徒同士でできます）"
                    Cells(otherRow, .Column).Interior.ColorIndex = prevC2
                    ActiveSheet.Protect
                    Exit Sub
                End If

                prevC = .Interior.ColorIndex
                If ownCol > 0 And otherRow > 0 Then
                    Cells(otherRow, ownCol).Interior.ColorIndex = tColor
                End If
                .Interior.ColorIndex = tColor

                m = MsgBox("入れ替え・移動・指定を実行しますか?", vbOKCancel)
                If m = vbOK Then
                    .Interior.ColorIndex = mColor
                    If CStr(.Value) <> "◎" Then .Value = "○"
                    Range("RowFail").Cells(.Row - r0, 1) = ""
                    If ownCol > 0 And otherRow > 0 Then
                        '入れ替え: 相手は自分の旧コマへ
                        Cells(otherRow, ownCol).Interior.ColorIndex = mColor
                        If CStr(Cells(otherRow, ownCol).Value) <> "◎" Then
                            Cells(otherRow, ownCol).Value = "○"
                        End If
                        Cells(otherRow, .Column).Interior.ColorIndex = MyPrev(.Column)
                        Cells(.Row, ownCol).Interior.ColorIndex = MyPrev(ownCol)
                        Cells(otherRow, STATE_COL) = ownCol - c0 + 1
                    ElseIf ownCol > 0 Then
                        '移動: 旧コマを解除
                        Cells(.Row, ownCol).Interior.ColorIndex = MyPrev(ownCol)
                    End If
                    Cells(.Row, STATE_COL) = .Column - c0 + 1
                    Application.CalculateFull
                Else
                    If ownCol > 0 And otherRow > 0 Then
                        Cells(otherRow, ownCol).Interior.ColorIndex = MyPrev(ownCol)
                    End If
                    .Interior.ColorIndex = prevC
                End If
                ActiveSheet.Protect
            End If
        End With
    End If
End Sub
