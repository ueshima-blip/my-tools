Attribute VB_Name = "UserForm2"
Attribute VB_Base = "0{529E8CA9-86ED-43DF-869B-1F79433DBDBD}{BC0C48F9-A831-4EA1-81D5-256FA9AF5C81}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

'==============================================================
' メニューフォーム
'   CommandButton1: 希望入力・休憩設定（パレットを開く）
'   CommandButton2: 調整作業
'   CommandButton3: 予定表・通知票の作成
'   CommandButton4: 閉じる
'==============================================================

Private Sub CommandButton1_Click()
    Unload Me
    UserForm3.Show
End Sub

Private Sub CommandButton2_Click()
    Dim i As Long, m As Long
    Dim nBlank As Long
    Dim blanks As String
    Dim q As String

    On Error GoTo Trouble
    fail = False
    KomaCheck
    If fail = True Then
        MsgBox "赤色表示の行を修正してください。"
        Unload Me
        UserForm3.Show
        Exit Sub
    End If

    If ActiveSheet.Name = "希望入力" Then
        nBlank = 0
        blanks = ""
        For i = 1 To ss
            If Val(CStr(Cells(Range("SEGMENTS").Row + i, Range("SEGMENTS").Column).Value)) = 0 Then
                nBlank = nBlank + 1
                If nBlank <= 10 Then
                    blanks = blanks & CStr(Val(CStr(Sheets("名簿").Cells(1 + i, 1).Value))) & " "
                End If
            End If
        Next
        q = "希望入力シートをコピーして、調整を開始します。"
        If nBlank > 0 Then
            q = "未入力の生徒が " & nBlank & " 名います（番号: " & Trim(blanks) & _
                IIf(nBlank > 10, " …", "") & "）。" & vbLf & _
                "未入力の生徒は飛ばして調整します。" & vbLf & vbLf & q
        End If
        m = MsgBox(q, vbOKCancel)
        If m = vbCancel Then Exit Sub
        Unload Me
        SheetCopy
        RunChosei
    Else
        m = MsgBox("確定コマの入れ替え・手動指定をする場合は「はい」" & vbLf & _
            "現在の確定をリセットして調整し直す場合は「いいえ」", vbYesNoCancel)
        If m = vbCancel Then Exit Sub
        If m = vbYes Then
            MsgBox "新しく確定したい場所をクリックしてください。" & vbLf & vbLf & _
                "・空きセルをクリック → そこに確定（入れ替え/移動も自動）" & vbLf & _
                "・修正が終わったら、メニューから予定表を作成してください。" & vbLf & vbLf & _
                "*** 他のシートを表示するまでこのモードは続きます ***"
            ChoseiSagyo = "ChangeColor"
            Unload Me
            Exit Sub
        End If
        Unload Me
        KomaReset
        RunChosei
    End If
    Exit Sub
Trouble:
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました: " & Err.Description, vbExclamation
End Sub

Private Sub CommandButton3_Click()
    Dim src As String
    Dim schName As String
    Dim ntfName As String
    Dim done As String

    On Error GoTo Trouble
    If Left(ActiveSheet.Name, 2) <> "調整" Then
        MsgBox "調整シート上で実行してください。"
        Exit Sub
    End If
    src = ActiveSheet.Name
    If MsgBox("現在の調整シートをもとに予定表を作成します。", vbOKCancel) = vbCancel Then Exit Sub
    Unload Me

    schName = MakeScheduleCore()
    ntfName = ""
    If MsgBox("保護者へ配る『個別通知票』も作成しますか?" & vbLf & _
        "（一人ずつの切り取り式お知らせ・A4 1枚に10名分）", vbYesNo) = vbYes Then
        ntfName = MakeNotifySheet(src)
    End If

    done = "作成しました: " & schName
    If ntfName <> "" Then done = done & " と " & ntfName
    done = done & vbLf & "このシートは自由に加工して印刷できます。" & vbLf & _
        "（予定表の生徒番号を変更すれば、氏名も変わります）"
    MsgBox done

    If MsgBox("PDF でも保存しますか?", vbYesNo) = vbYes Then
        ExportSheetPDF schName
        If ntfName <> "" Then ExportSheetPDF ntfName
    End If
    Exit Sub
Trouble:
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました: " & Err.Description, vbExclamation
End Sub

Private Sub CommandButton4_Click()
    Unload Me
End Sub

Private Sub UserForm_Activate()
    If Left(ActiveSheet.Name, 2) = "調整" Then
        Me.CommandButton3.Enabled = True
    Else
        Me.CommandButton3.Enabled = False
    End If
End Sub
