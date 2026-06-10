Attribute VB_Name = "UserForm3"
Attribute VB_Base = "0{57F01445-7368-4648-B2E9-32CF6009A3A8}{E47D60E3-94AA-43E3-94A3-5A7E785ACE3C}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Private Sub CommandButton1_Click()

    With Selection
        If .Row <= Range("FIRST").Row Or .Row > ss + Range("FIRST").Row Or _
            .Column < Range("FIRST").Column Or .Column > Range("FIRST").Column + kk * dd - 1 Then Exit Sub
        Selection.Value = "○"
        If Range("RowFail").Cells(.Row - Range("FIRST").Row, 1) = 1 Then
           Range("RowFail").Cells(.Row - Range("FIRST").Row, 1) = ""
        End If
    End With
    
    If CheckBox1 = True Then                    '他日にコピー
        CopyTo
    End If
    

End Sub

Private Sub CommandButton2_Click()
    
    fr = Range("FIRST").Row
    fc = Range("FIRST").Column
    With Selection
        If .Row <= fr Or .Row > ss + fr Or .Column < fc Or .Column > fc + kk * dd - 1 Then Exit Sub
        If .Columns.Count > 1 Or _
                WorksheetFunction.CountIf(Range(Cells(.Row, fc), Cells(.Row, fc + 59)), "◎") > 0 Or _
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
        Selection.Value = "△"
    End With
    
    If CheckBox1 = True Then                    '他日にコピー
        CopyTo
    End If
    
End Sub

Private Sub CommandButton4_Click()

    Unload Me
    Range("RowNow").ClearContents
    
    ActiveSheet.Protect AllowFormattingCells:=False

    UserForm2.Show

End Sub

Private Sub CommandButton5_Click()

    r = Selection.Row
    c = Selection.Column
    cf = Range("FIRST").Column
    If r <> Range("FIRST").Row Or c < cf Or c > cf + kk * dd - 1 Then Exit Sub
    If Selection.Interior.ColorIndex = kColor Then Exit Sub
    If Val(ss + kNumber) = Val(dd * kk) Then
        m = MsgBox("コマ数の下限を下回りますが、設定しますか。(実施しない生徒がいる場合等)", vbYesNo)
        If m = vbNo Then Exit Sub
    End If
    
    If CheckBox2 Then
        c = (c - cf + 1) Mod kk
        If c = 0 Then c = kk
        For i = 1 To dd
            Range(Cells(r, cf + c - 1 + kk * (i - 1)), _
                  Cells(r + ss, cf + c - 1 + kk * (i - 1))).Interior.ColorIndex = kColor
        Next
        CheckBox2 = False
    Else
        Range(Cells(r, c), Cells(r + ss, c)).Interior.ColorIndex = kColor
    End If
    
    kNumber = 0
    r = Range("FIRST").Row
    c = Range("FIRST").Column
    For i = 0 To 59
        If Cells(r, c + i).Interior.ColorIndex = kColor Then
            kNumber = kNumber + 1
        End If
    Next
    
End Sub

Private Sub CommandButton6_Click()
    
    r = Selection.Row
    c = Selection.Column
    cf = Range("FIRST").Column

    If r <> Range("FIRST").Row Or c < cf Or c > cf + kk * dd - 1 Then Exit Sub
    If Selection.Interior.ColorIndex = kColor Then
        If CheckBox2 Then
            c = (c - cf + 1) Mod kk
            If c = 0 Then c = kk
            For i = 1 To dd
                Range(Cells(r, cf + c - 1 + kk * (i - 1)), _
                      Cells(r + 50, cf + c - 1 + kk * (i - 1))).Interior.ColorIndex = xlNone
            Next
            CheckBox2 = False
        Else
            Range(Cells(r, c), Cells(r + 50, c)).Interior.Color = xlNone
        End If
    End If
    
    kNumber = 0
    r = Range("FIRST").Row
    c = Range("FIRST").Column
    For i = 0 To 59
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

        Selection.Copy
        For i = 2 To dd
            If Val(Selection.Column + Selection.Columns.Count) > Val(kk * (dd - 1) + Range("FIRST").Column) Then
                Exit For
            End If
            Cells(Selection.Row, Selection.Column + kk).PasteSpecial xlPasteValues
        Next
        Application.CutCopyMode = False
        CheckBox1 = False

End Sub


Private Sub UserForm_Terminate()

    Range("RowNow").ClearContents
    ActiveSheet.Protect AllowFormattingCells:=False

End Sub
