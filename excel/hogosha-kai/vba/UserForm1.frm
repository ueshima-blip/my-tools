Attribute VB_Name = "UserForm1"
Attribute VB_Base = "0{6E23E6A0-CD43-4D1E-B441-C0E3AE4A6961}{6D9B0881-9FA7-4015-9B85-423D3CADD24C}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

'==============================================================
' 初期設定フォーム（日数・面談時間・開始/最終時刻）
'==============================================================

Private Sub CommandButton1_Click()
    Dim k As Double
    Dim sTime As Date, lTime As Date
    Dim kpd As Long
    Dim d As Long, t As Long
    Dim r0 As Long, c0 As Long
    Dim m As Long
    Dim ws As Worksheet

    On Error GoTo Trouble

    Sheets("Start!").Unprotect
    Sheets("Start!").Range("日数") = ComboBox1.Value
    Sheets("Start!").Range("時間") = ComboBox2.Value
    Sheets("Start!").Range("開始時") = ComboBox3.Value
    Sheets("Start!").Range("開始分") = ComboBox4.Value
    Sheets("Start!").Range("最終時") = ComboBox5.Value
    Sheets("Start!").Range("最終分") = ComboBox6.Value
    Sheets("Start!").Range("協力度") = ""
    Sheets("Start!").Protect

    sTime = TimeSerial(CLng(Val(ComboBox3.Value)), CLng(Val(ComboBox4.Value)), 0)
    lTime = TimeSerial(CLng(Val(ComboBox5.Value)), CLng(Val(ComboBox6.Value)), 0)
    If lTime <= sTime Then
        MsgBox "最終開始時間が開始時間より前になっています。"
        Exit Sub
    End If
    k = (lTime - sTime) / TimeSerial(0, CLng(Val(ComboBox2.Value)), 0)
    If Abs(k - CLng(k)) > 0.000001 Then
        MsgBox "時間設定が正しくありません。面談時間と(最終)開始時間を確認してください。"
        Exit Sub
    End If
    kpd = CLng(k) + 1
    If CLng(Val(ComboBox1.Value)) * kpd > 60 Then
        MsgBox "コマ数が多すぎるため、設定できません。（最大 計60コマ）"
        Exit Sub
    End If

    '生徒数に対する過不足を確認画面に表示（40人学級などで一目で分かるように）
    Dim totalSlots As Long, nStu As Long
    Dim capMsg As String
    totalSlots = CLng(Val(ComboBox1.Value)) * kpd
    nStu = CLng(Val(CStr(Sheets("Start!").Range("生徒数").Value)))
    capMsg = "希望入力シートを初期化し、" & Val(ComboBox1.Value) & " 日 × " & kpd & _
        " コマ（合計 " & totalSlots & " コマ）で設定します。"
    If nStu > 0 Then
        If totalSlots < nStu Then
            capMsg = capMsg & vbLf & vbLf & "※注意: 生徒 " & nStu & " 名に対してコマ数が足りません。" & vbLf & _
                "　日数を増やすか、面談時間・開始/最終時間を見直してください。"
        Else
            capMsg = capMsg & vbLf & "（生徒 " & nStu & " 名 ・ 休憩などに使える余裕 " & _
                (totalSlots - nStu) & " コマ）"
        End If
    End If
    m = MsgBox(capMsg, vbOKCancel)
    If m = vbCancel Then Exit Sub

    Application.ScreenUpdating = False
    Set ws = Sheets("希望入力")
    ws.Activate
    ws.Unprotect

    r0 = Range("FIRST").Row
    c0 = Range("FIRST").Column

    With ws.Range(ws.Cells(r0 - 1, c0), ws.Cells(r0 + 50, c0 + 59))
        .ClearContents
        .Interior.ColorIndex = xlNone
        .MergeCells = False
        .HorizontalAlignment = xlCenter
        .Borders(xlInsideVertical).Weight = xlHairline
    End With

    For d = 1 To CLng(Val(ComboBox1.Value))
        For t = 0 To kpd - 1
            ws.Cells(r0, c0 + t + kpd * (d - 1)) = _
                Format(DateAdd("n", CLng(Val(ComboBox2.Value)) * t, sTime), "hh:mm")
        Next
        '日の右境界線
        ws.Range(ws.Cells(r0 - 1, c0 + kpd * d), ws.Cells(r0 + 50, c0 + kpd * d)) _
            .Borders(xlEdgeLeft).Weight = xlThin
        '日付セル（結合 + 灰色）
        With ws.Range(ws.Cells(r0 - 1, c0 + kpd * (d - 1)), ws.Cells(r0 - 1, c0 + kpd * d - 1))
            .Merge
            .Interior.Color = RGB(216, 216, 216)
        End With
    Next

    '時刻ヘッダ: 秒なし表示 + 列幅に合わせて自動縮小（####防止）
    With ws.Range(ws.Cells(r0, c0), ws.Cells(r0, c0 + 59))
        .NumberFormat = "h:mm"
        .ShrinkToFit = True
    End With

    Range("RowFail").ClearContents
    ws.Range(ws.Cells(r0, STATE_COL), ws.Cells(r0 + 50, STATE_COL)).ClearContents
    ws.Protect
    Application.ScreenUpdating = True

    Unload Me
    ws.Range("D1").Select
    MsgBox "灰色の日付セルに日付（例: 6/20）を入力してください。" & vbLf & vbLf & _
        "希望の入力方法は3つ:" & vbLf & _
        "・Start! の「希望を取り込む」: フォームやOMRのファイルから自動入力" & vbLf & _
        "・Start! の「クイック入力」: 紙を見ながらキーボードで連続入力" & vbLf & _
        "・希望入力シートの「メニュー」: セルを選んで○◎△を手入力"
    Exit Sub
Trouble:
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました: " & Err.Description, vbExclamation
End Sub

Private Sub CommandButton2_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    On Error Resume Next
    With ComboBox1
        .AddItem "2"
        .AddItem "3"
        .AddItem "4"
        .AddItem "5"
        .Value = CStr(Sheets("Start!").Range("日数").Value)
    End With
    With ComboBox2
        .AddItem "10"
        .AddItem "15"
        .AddItem "20"
        .AddItem "25"
        .AddItem "30"
        .Value = CStr(Sheets("Start!").Range("時間").Value)
    End With
    With ComboBox3
        .AddItem "09"
        .AddItem "10"
        .AddItem "11"
        .AddItem "12"
        .AddItem "13"
        .AddItem "14"
        .AddItem "15"
        .AddItem "16"
        .AddItem "17"
        .Value = CStr(Sheets("Start!").Range("開始時").Value)
    End With
    With ComboBox4
        .AddItem "00"
        .AddItem "05"
        .AddItem "10"
        .AddItem "15"
        .AddItem "20"
        .AddItem "25"
        .AddItem "30"
        .AddItem "35"
        .AddItem "40"
        .AddItem "45"
        .AddItem "50"
        .AddItem "55"
        .Value = CStr(Sheets("Start!").Range("開始分").Value)
    End With
    With ComboBox5
        .AddItem "11"
        .AddItem "12"
        .AddItem "13"
        .AddItem "14"
        .AddItem "15"
        .AddItem "16"
        .AddItem "17"
        .AddItem "18"
        .AddItem "19"
        .Value = CStr(Sheets("Start!").Range("最終時").Value)
    End With
    With ComboBox6
        .AddItem "00"
        .AddItem "05"
        .AddItem "10"
        .AddItem "15"
        .AddItem "20"
        .AddItem "25"
        .AddItem "30"
        .AddItem "35"
        .AddItem "40"
        .AddItem "45"
        .AddItem "50"
        .AddItem "55"
        .Value = CStr(Sheets("Start!").Range("最終分").Value)
    End With
End Sub
