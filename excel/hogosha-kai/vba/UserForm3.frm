Attribute VB_Name = "UserForm3"
Attribute VB_Base = "0{57F01445-7368-4648-B2E9-32CF6009A3A8}{E47D60E3-94AA-43E3-94A3-5A7E785ACE3C}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

'==============================================================
' 希望入力・休憩設定パレット
'   CommandButton1: ○  CommandButton2: ◎  CommandButton3: △
'   CommandButton5: 休憩設定  CommandButton6: 休憩解除
'   CommandButton4: 終了（メニューへ戻る）
'   CheckBox1: 他の日にもコピー  CheckBox2: 同時刻を全日に設定
'==============================================================

Private Sub CommandButton1_Click()
    With Selection
        If .Row <= Range("FIRST").Row Or .Row > ss + Range("FIRST").Row Or _
            .Column < Range("FIRST").Column Or .Column > Range("FIRST").Column + kk * dd - 1 Then Exit Sub
        .Value = "○"
        If Range("RowFail").Cells(.Row - Range("FIRST").Row, 1) = 1 Then
            Range("RowFail").Cells(.Row - Range("FIRST").Row, 1) = ""
        End If
    End With
    If CheckBox1.Value = True Then CopyTo
End Sub

Private Sub CommandButton2_Click()
    Dim fr As Long, fc As Long
    fr = Range("FIRST").Row
    fc = Range("FIRST").Column
    With Selection
        If .Row <= fr Or .Row > ss + fr Or .Column < fc Or .Column > fc + kk * dd - 1 Then Exit Sub
        If .Columns.Count > 1 Or .Rows.Count > 1 Or _
                WorksheetFunction.CountIf(Range(Cells(.Row, fc), Cells(.Row, fc + kk * dd - 1)), "◎") > 0 Or _
                WorksheetFunction.CountIf(Range(Cells(fr + 1, .Column), Cells(fr + ss, .Column)), "◎") > 0 Then
            MsgBox "優先設定は同一時間・同一生徒ともに一カ所のみです。"
            Exit Sub
        End If
        .Value = "◎"
        If Range("RowFail").Cells(.Row - fr, 1) = 1 Then
            Range("RowFail").Cells(.Row - fr, 1) = ""
        End If
    End With
End Sub

Private Sub CommandButton3_Click()
    With Selection
        If .Row <= Range("FIRST").Row Or .Row > ss + Range("FIRST").Row Or _
            .Column < Range("FIRST").Column Or .Column > Range("FIRST").Column + kk * dd - 1 Then Exit Sub
        .Value = "△"
    End With
    If CheckBox1.Value = True Then CopyTo
End Sub

Private Sub CommandButton4_Click()
    Unload Me
    Range("RowNow").ClearContents
    ActiveSheet.Protect AllowFormattingCells:=False
    UserForm2.Show
End Sub

Private Sub CommandButton5_Click()
    Dim r As Long, c As Long, cf As Long
    Dim m As Long, i As Long
    r = Selection.Row
    c = Selection.Column
    cf = Range("FIRST").Column
    If r <> Range("FIRST").Row Or c < cf Or c > cf + kk * dd - 1 Then Exit Sub
    If Selection.Interior.ColorIndex = kColor Then Exit Sub
    If ss + kNumber >= dd * kk Then
        m = MsgBox("コマ数の下限を下回りますが、設定しますか。（実施しない生徒がいる場合等）", vbYesNo)
        If m = vbNo Then Exit Sub
    End If

    If CheckBox2.Value = True Then
        c = (c - cf + 1) Mod kk
        If c = 0 Then c = kk
        For i = 1 To dd
            Range(Cells(r, cf + c - 1 + kk * (i - 1)), _
                  Cells(r + ss, cf + c - 1 + kk * (i - 1))).Interior.ColorIndex = kColor
        Next
        CheckBox2.Value = False
    Else
        Range(Cells(r, c), Cells(r + ss, c)).Interior.ColorIndex = kColor
    End If

    RecountBreaks
End Sub

Private Sub CommandButton6_Click()
    Dim r As Long, c As Long, cf As Long
    Dim i As Long
    r = Selection.Row
    c = Selection.Column
    cf = Range("FIRST").Column
    If r <> Range("FIRST").Row Or c < cf Or c > cf + kk * dd - 1 Then Exit Sub
    If Selection.Interior.ColorIndex = kColor Then
        If CheckBox2.Value = True Then
            c = (c - cf + 1) Mod kk
            If c = 0 Then c = kk
            For i = 1 To dd
                Range(Cells(r, cf + c - 1 + kk * (i - 1)), _
                      Cells(r + ss, cf + c - 1 + kk * (i - 1))).Interior.ColorIndex = xlNone
            Next
            CheckBox2.Value = False
        Else
            Range(Cells(r, c), Cells(r + ss, c)).Interior.ColorIndex = xlNone
        End If
    End If

    RecountBreaks
End Sub

Private Sub RecountBreaks()
    Dim r As Long, c As Long, i As Long
    kNumber = 0
    r = Range("FIRST").Row
    c = Range("FIRST").Column
    For i = 0 To dd * kk - 1
        If Cells(r, c + i).Interior.ColorIndex = kColor Then
            kNumber = kNumber + 1
        End If
    Next
End Sub

Private Sub UserForm_Activate()
    If fail = False Then KomaCheck
    ActiveSheet.Protect AllowFormattingCells:=True
    nowSh = ActiveSheet.Name
End Sub

Private Sub CopyTo()
    Dim i As Long
    Selection.Copy
    For i = 2 To dd
        If Selection.Column + Selection.Columns.Count > kk * (dd - 1) + Range("FIRST").Column Then
            Exit For
        End If
        Cells(Selection.Row, Selection.Column + kk).PasteSpecial xlPasteValues
    Next
    Application.CutCopyMode = False
    CheckBox1.Value = False
End Sub

Private Sub UserForm_Terminate()
    Range("RowNow").ClearContents
    ActiveSheet.Protect AllowFormattingCells:=False
End Sub
