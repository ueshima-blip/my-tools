Attribute VB_Name = "UserForm1"
Attribute VB_Base = "0{6E23E6A0-CD43-4D1E-B441-C0E3AE4A6961}{6D9B0881-9FA7-4015-9B85-423D3CADD24C}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Private Sub CommandButton1_Click()
    
    Dim k As Single
    
    ActiveSheet.Unprotect
    Range("日数") = ComboBox1
    Range("時間") = ComboBox2
    Range("開始時") = ComboBox3
    Range("開始分") = ComboBox4
    Range("最終時") = ComboBox5
    Range("最終分") = ComboBox6
    Range("協力度") = ""
    ActiveSheet.Protect
    
    s = TimeSerial(ComboBox3, ComboBox4, 0)
    l = TimeSerial(ComboBox5, ComboBox6, 0)
    k = (l - s) / TimeSerial(0, ComboBox2, 0)
    If k <> Int(k) Then
        MsgBox ("時間設定が正しくありません。面談時間と(最終)開始時間を確認してください。 ")
        Exit Sub
    ElseIf Val(ComboBox1) * (k + 1) > 60 Then
        MsgBox ("コマ数が多すぎるため、設定できません。")
        Exit Sub
    End If
    m = MsgBox("希望入力シートを初期化し、" & Val(ComboBox1) & "日 ×" & k + 1 & " コマで設定します。", vbOKCancel)
    If m = vbCancel Then Exit Sub
        
    Application.ScreenUpdating = False
    Sheets("希望入力").Select
    ActiveSheet.Unprotect
    
    Range(Cells(Range("FIRST").Row - 1, Range("FIRST").Column), _
        Cells(Range("FIRST").Row + 50, Range("FIRST").Column + 59)).Select
    With Selection
        .ClearContents
        .Interior.Color = xlNone
        .MergeCells = False
        .HorizontalAlignment = xlCenter
        .Borders(xlInsideVertical).Weight = xlHairline
    End With
    For d = 1 To ComboBox1
        For t = 0 To k
            Cells(Range("FIRST").Row, Range("FIRST").Column + t + (k + 1) * (d - 1)) = _
                Format(DateAdd("n", ComboBox2 * t, s), "hh:mm")
        Next
        Range(Cells(Range("FIRST").Row - 1, Range("FIRST").Column + t * d), _
            Cells(Range("FIRST").Row + 50, Range("FIRST").Column + t * d)).Select
        Selection.Borders(xlEdgeLeft).Weight = xlThin
        Range(Cells(Range("FIRST").Row - 1, Range("FIRST").Column + t * d - k - 1), _
            Cells(Range("FIRST").Row - 1, Range("FIRST").Column + t * d - 1)).Select
        Selection.Merge
        Selection.Interior.Color = RGB(216, 216, 216)
    Next
    Range("RowFail").ClearContents
    ActiveSheet.Protect
    
    Unload Me
    Application.ScreenUpdating = True
    Range("D1").Select
    MsgBox "日付を入力後、希望入力をしてください。"

End Sub

Private Sub CommandButton2_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()

    With ComboBox1
        .AddItem "2"
        .AddItem "3"
        .AddItem "4"
        .AddItem "5"
        .Value = Range("日数")
    End With
    With ComboBox2
        .AddItem "10"
        .AddItem "15"
        .AddItem "20"
        .AddItem "25"
        .AddItem "30"
        .Value = Range("時間")
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
        .Value = Range("開始時")
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
        .Value = Range("開始分")
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
        .Value = Range("最終時")
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
        .Value = Range("最終分")
    End With


End Sub
