Attribute VB_Name = "UserForm2"
Attribute VB_Base = "0{529E8CA9-86ED-43DF-869B-1F79433DBDBD}{BC0C48F9-A831-4EA1-81D5-256FA9AF5C81}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub CommandButton1_Click()

    Unload Me
    UserForm3.Show

End Sub

Private Sub CommandButton2_Click()

    fail = False
    KomaCheck
    If fail = True Then
        MsgBox "設定を変更してください。"
        Unload Me
        UserForm3.Show
        Exit Sub
    End If
    
    For i = 1 To ss
        If Cells(Range("SEGMENTS").Row + i, Range("SEGMENTS").Column) = 0 Then
            m = MsgBox(i & " 番が未入力です。続行しますか。", vbYesNo)
            If m = vbNo Then Exit Sub
        End If
    Next
    
    For Each Sh In Sheets
        If Left(Sh.Name, 2) = "調整" Then f = True
    Next
    
    Select Case ActiveSheet.Name
    Case "希望入力"
        If f = False Then
            m = MsgBox("希望入力シートをコピーし、調整を開始します。", vbOKCancel)
            If m = vbOK Then
                Unload Me
                SheetCopy
                Chosei
            Else
                Exit Sub
            End If
        Else
            m = MsgBox("新規に調整を開始します。すでにある調整シートを利用する場合は、" & _
                    "Cancelをして、そのシート上で作業を開始してください。", vbOKCancel)
            If m = vbOK Then
                Unload Me
                SheetCopy
                Chosei
            Else
                Exit Sub
            End If
        End If
    Case Else
        m = MsgBox("確定コマの入れ替えなどをする場合は「はい」" & Chr(13) & _
                    "現在の確定コマをリセットして調整し直す場合は「いいえ」", vbYesNoCancel)
        If m = vbCancel Then Exit Sub
        If m = vbYes Then
            MsgBox "新しく確定したい場所をクリックしてください。" & _
                    Chr(13) & Chr(13) & _
                    "修正が終わったら、メニューより予定表を作成して下さい。" & _
                    Chr(13) & Chr(13) & _
                    "***他のシートを表示するまでこのモードは続きます***"
            ChoseiSagyo = "ChangeColor"
            Unload Me
            
            
            Exit Sub
        End If
'        Range("RowFail").ClearContents
        Unload Me
        scrUpdate = True
        KomaReset
        Chosei
    End Select
    
End Sub

Private Sub CommandButton3_Click()

    m = MsgBox("現在の調整シートをもとに予定表を作成します。", vbOKCancel)
    If m = vbCancel Then Exit Sub
    
    n = 1
Again:
    For Each Sh In Sheets
        If Sh.Name = "予定表" & CStr(n) Then
            n = n + 1
            GoTo Again
        End If
    Next
    
    Application.ScreenUpdating = False
    Sh = ActiveSheet.Name                       '元_予定表をコピー
    Sheets("元_予定表").Copy After:=Sheets(Sh)
    Sheets("元_予定表 (2)").Visible = True
    Sheets("元_予定表 (2)").Activate
    ActiveSheet.Name = "予定表" & CStr(n)
    ActiveSheet.Unprotect
    
    dd = Val(Range("日数"))
    kk = Range("コマ数")
    ss = Range("生徒数")
    rd = Range("表日付").Row
    cd = Range("表日付").Column
    rt = Range("表時間").Row
    ct = Range("表時間").Column
    
    If dd < 5 Then                              '不要な列削除
        cs = dd * 3
        cc = 2 + 3 * (4 - dd)
        Range(Cells(rd, cd + cs), Cells(rd, cd + cs + cc)).EntireColumn.Delete
    End If
                                                
    If kk < 30 Then                             '不要な行削除
        Range(Cells(rt + kk, ct), Cells(rt + kk + (29 - kk), ct)).EntireRow.Delete
    End If
    
    rr = Range("FIRST").Row
    cc = Range("FIRST").Column
    For i = 1 To kk                             '時間転記
        Cells(rt + i - 1, ct) = Sheets(Sh).Cells(rr, cc + i - 1)
    Next i
    For i = 0 To dd - 1                           '日付転記
        Cells(rd, cd + 3 * i) = Sheets(Sh).Cells(rr - 1, cc + kk * i)
    Next
     
    rn = Range("表番号").Row                    '生徒番号転記
    cn = Range("表番号").Column
    For d = 1 To dd
        For k = 1 To kk
            For s = 1 To ss
                If Sheets(Sh).Cells(rr + s, cc - 1 + k + (kk * (d - 1))).Interior.ColorIndex = mColor Or _
                    Sheets(Sh).Cells(rr + s, cc - 1 + k + (kk * (d - 1))).Interior.ColorIndex = pColor Then
                    Cells(rn + k, cn + 3 * (d - 1)) = s
                End If
            Next
        Next
    Next
    
    ActiveSheet.Protect
    Unload Me
    Application.ScreenUpdating = True
    MsgBox "このシートを自由に加工し、印刷してください。" & Chr(13) & Chr(13) & _
            "(生徒番号を変更すれば、生徒名も変わります。)"
    
    
End Sub

Private Sub CommandButton4_Click()

    Unload Me

End Sub

Private Sub UserForm_Activate()
    
    If ActiveSheet.Name = "希望入力" Then
        Me.CommandButton3.Enabled = False
    End If

End Sub

