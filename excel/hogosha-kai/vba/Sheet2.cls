Attribute VB_Name = "Sheet2"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Attribute VB_Control = "CommandButton1, 2, 0, MSForms, CommandButton"

Private Sub CommandButton1_Click()

    ss = Sheets("Start!").Range("生徒数")
    If ss = 0 Then
        MsgBox "名簿を貼り付けてください。"
        Exit Sub
    End If
    UserForm2.Show

End Sub


Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)

    If UserForm3.Visible = False Then Exit Sub
    With Selection
        If .Row <= Range("FIRST").Row Or .Row > ss + Range("FIRST").Row Or _
            .Column < Range("FIRST").Column Or .Column > Range("FIRST").Column + kk * dd - 1 Then Exit Sub
        Dim c As Range
        For Each c In Selection
            If c <> "" Then
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
    
    Range("RowNow").ClearContents
    ChoseiSagyo = ""
    
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    If UserForm3.Visible = True Then
        
        If Cells(Range("RowNow").Row - 1, Range("RowNow").Column) = 1 Or _
            Cells(Target.Row, Range("RowNow").Column) = 1 Then Exit Sub
    
        Range("RowNow").ClearContents
        ss = Sheets("Start!").Range("生徒数")
    
        If Target.Row <= Range("FIRST").Row Or Target.Row > Range("FIRST").Row + ss Then Exit Sub

        Range("RowNow").Cells(Target.Row - Range("FIRST").Row, 1) = 1
    
    Else
        
        If ChoseiSagyo = "ChangeColor" Then
            With Selection
                If .Row <= Range("FIRST").Row Or .Row > ss + Range("FIRST").Row Or _
                    .Column < Range("FIRST").Column Or .Column > Range("FIRST").Column + kk * dd - 1 Then Exit Sub
        
                If .Interior.ColorIndex = xlNone Or .Interior.ColorIndex = kColor Then
                    ActiveSheet.Unprotect
                    s = Range("FIRST").Column
                    For c = s To dd * kk + s - 1           '自分の確定コマを捜す
                        If Cells(Target.Row, c).Interior.ColorIndex = mColor Or _
                           Cells(Target.Row, c).Interior.ColorIndex = pColor Or _
                           Cells(Target.Row, c).Interior.ColorIndex = dColor Then Exit For
                    Next
                    If c = dd * kk + s Then jibun = 1
                    
                    For r = Range("FIRST").Row + 1 To Range("FIRST").Row + ss   '相手の確定コマを捜す
                        If Cells(r, Target.Column).Interior.ColorIndex = mColor Or _
                           Cells(r, Target.Column).Interior.ColorIndex = pColor Then Exit For
                    Next
                    If r = Range("First").Row + ss + 1 Then aite = 1
                    
                    prevC = .Interior.ColorIndex
                    If jibun <> 1 And aite <> 1 Then
                        Cells(r, c).Interior.ColorIndex = tColor
                    ElseIf jibun = 1 And aite <> 1 Then
                        prevC2 = Cells(r, Target.Column).Interior.ColorIndex
                        Cells(r, Target.Column).Interior.ColorIndex = dColor
                        MsgBox "重複するコマがあります。"
                        Cells(r, Target.Column).Interior.ColorIndex = prevC2
                        Exit Sub
                    End If
                    .Interior.ColorIndex = tColor
                    
                    m = MsgBox("入れ替え・移動・指定を実行しますか?", vbOKCancel)
                    
                    If m = vbOK Then
                        .Interior.ColorIndex = mColor
                        If .Value <> "◎" Then .Value = "○"
                        Cells(.Row, Range("RowFail").Column) = ""
                        If jibun <> 1 And aite <> 1 Then
                            Cells(r, c).Interior.ColorIndex = mColor
                            Cells(r, Target.Column).Interior.ColorIndex = MyPrev(Target.Column)
                            Cells(Target.Row, c).Interior.ColorIndex = MyPrev(c)
                        ElseIf jibun <> 1 Then
                            Cells(Target.Row, c).Interior.ColorIndex = MyPrev(c)
                        End If
                        Application.ScreenUpdating = False
                        Range("LastRow").Copy           'シート５７行目のユーザー関数再計算のため
                        Range("LastRow").Select
                        Selection.PasteSpecial xlValue
                        Application.ScreenUpdating = True
                    Else
                        If aite <> 1 Then
                            Cells(r, c).Interior.ColorIndex = MyPrev(c)
                        End If
                        .Interior.ColorIndex = prevC
                    End If
                    ActiveSheet.Protect
                End If
            End With
        End If
    End If
    
End Sub
