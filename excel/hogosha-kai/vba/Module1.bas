Attribute VB_Name = "Module1"
Option Explicit

'==============================================================
' 保護者会調整 ver2.0  メインモジュール
'   ・調整エンジン（増加路法: 解があれば必ず見つかる）
'   ・△（第2希望）対応 / 失敗時の原因診断
'   ・予定表・保護者向け通知票・PDF出力
'   ・シート上ボタンの入口（BtnXXX）
'==============================================================

'=== 背景色（ver1.11 から継承） ===
Public Const kColor = 24                '休憩指定の背景色
Public Const mColor = 33                '面談確定の背景色
Public Const pColor = 47                '優先(◎)確定の背景色
Public Const tColor = 22                '一時表示の背景色
Public Const dColor = 46                '重複表示の背景色

'=== 上限・補助列 ===
Public Const MAXS As Long = 50          '生徒数上限
Public Const MAXSLOT As Long = 60       '総コマ数上限
Public Const STATE_COL As Long = 71     'BS列: 確定コマ番号（裏データ）
Public Const STATE_VER As String = "STATE2"

'=== 共有状態 ===
Public dd As Long                       '日数
Public kk As Long                       '1日のコマ数
Public ss As Long                       '生徒数（名簿の最大番号）
Public kNumber As Long                  '休憩コマ数
Public fail As Boolean                  'KomaCheck の結果
Public ChoseiSagyo As String            '"ChangeColor"=入れ替えモード
Public nowSh As String                  '入力モード中のシート名

'--------------------------------------------------------------
' 設定値の読み込み（どこからでも呼んでよい）
'--------------------------------------------------------------
Public Sub EnsureDims()
    On Error Resume Next
    dd = CLng(Val(CStr(Sheets("Start!").Range("日数").Value)))
    kk = CLng(Val(CStr(Sheets("Start!").Range("コマ数").Value)) + 0.0000001)
    ss = CLng(Val(CStr(Sheets("Start!").Range("生徒数").Value)))
    On Error GoTo 0
    If dd < 0 Then dd = 0
    If kk < 0 Then kk = 0
    If ss < 0 Then ss = 0
    If ss > MAXS Then ss = MAXS
    If dd > 0 Then
        If dd * kk > MAXSLOT Then kk = MAXSLOT \ dd
    End If
End Sub

Private Sub CleanFail()
    Application.ScreenUpdating = True
    Application.CutCopyMode = False
    MsgBox "エラーが発生しました。" & vbLf & "(" & Err.Number & ") " & Err.Description, vbExclamation
End Sub

'--------------------------------------------------------------
' シート上のボタン入口
'--------------------------------------------------------------
Public Sub BtnSettings()
    On Error GoTo Trouble
    Sheets("Start!").Activate
    UserForm1.Show
    Exit Sub
Trouble:
    CleanFail
End Sub

Public Sub BtnMenu()
    On Error GoTo Trouble
    EnsureDims
    If ss = 0 Then
        MsgBox "先に「名簿」シートへ名簿を貼り付けてください。"
        Exit Sub
    End If
    UserForm2.Show
    Exit Sub
Trouble:
    CleanFail
End Sub

Public Sub BtnImport()
    On Error GoTo Trouble
    ImportWishes
    Exit Sub
Trouble:
    CleanFail
End Sub

Public Sub BtnQuickInput()
    On Error GoTo Trouble
    QuickInput
    Exit Sub
Trouble:
    CleanFail
End Sub

Public Sub BtnSurveyPrint()
    On Error GoTo Trouble
    MakeSurveySheets
    Exit Sub
Trouble:
    CleanFail
End Sub

'--------------------------------------------------------------
' 希望入力シートを「調整N」としてコピー
'--------------------------------------------------------------
Public Sub SheetCopy()
    Dim n As Long
    Dim shObj As Object
    n = 1
Again:
    For Each shObj In ThisWorkbook.Sheets
        If shObj.Name = "調整" & CStr(n) Then
            n = n + 1
            GoTo Again
        End If
    Next

    Sheets("希望入力").Copy After:=Sheets("希望入力")
    ActiveSheet.Name = "調整" & CStr(n)

    ActiveSheet.Protect AllowFormattingCells:=True
    If ss < MAXS Then
        ActiveSheet.Range(ActiveSheet.Cells(Range("NAMES").Row + ss + 1, 1), _
            ActiveSheet.Cells(Range("NAMES").Row + MAXS, 3)).Interior.ColorIndex = xlNone
    End If
End Sub

'--------------------------------------------------------------
' 確定（青・紫）だけを消す
'--------------------------------------------------------------
Public Sub KomaReset()
    Dim r0 As Long, c0 As Long, r As Long, c As Long
    On Error GoTo Trouble
    EnsureDims
    Application.ScreenUpdating = False
    ActiveSheet.Unprotect
    r0 = Range("FIRST").Row
    c0 = Range("FIRST").Column
    For r = r0 + 1 To r0 + ss
        For c = c0 To c0 + dd * kk - 1
            With ActiveSheet.Cells(r, c).Interior
                If .ColorIndex = mColor Or .ColorIndex = pColor Then .ColorIndex = xlNone
            End With
        Next
        ActiveSheet.Cells(r, STATE_COL).ClearContents
    Next
    ActiveSheet.Cells(r0, STATE_COL).ClearContents
    ActiveSheet.Protect AllowFormattingCells:=True
    Application.ScreenUpdating = True
    Exit Sub
Trouble:
    CleanFail
End Sub

'--------------------------------------------------------------
' 入力内容の事前チェック
'   RowFail: 1=未入力(灰), 2=問題あり(赤)
'--------------------------------------------------------------
Public Sub KomaCheck()
    Dim r0 As Long, c0 As Long, i As Long, j As Long
    Dim isBreak(1 To MAXSLOT) As Boolean
    Dim v As String, msg As String
    Dim nMark As Long, nBreakMaru As Long, nijuInBreak As Boolean

    On Error GoTo Trouble
    EnsureDims
    fail = False
    If ss = 0 Or dd * kk = 0 Then Exit Sub
    r0 = Range("FIRST").Row
    c0 = Range("FIRST").Column

    Application.ScreenUpdating = False
    Range("RowFail").ClearContents

    kNumber = 0
    For j = 1 To dd * kk
        isBreak(j) = (ActiveSheet.Cells(r0, c0 + j - 1).Interior.ColorIndex = kColor)
        If isBreak(j) Then kNumber = kNumber + 1
    Next

    For i = 1 To ss
        nMark = 0
        nBreakMaru = 0
        nijuInBreak = False
        For j = 1 To dd * kk
            v = CStr(ActiveSheet.Cells(r0 + i, c0 + j - 1).Value)
            If isBreak(j) Then
                If v = "◎" Then nijuInBreak = True
                If v = "○" Then nBreakMaru = nBreakMaru + 1
            Else
                If v = "○" Or v = "◎" Or v = "△" Then nMark = nMark + 1
            End If
        Next
        If nijuInBreak Then
            If InStr(msg, "休憩時間帯に優先マーク") = 0 Then _
                msg = msg & "休憩時間帯に優先マーク(◎)があります。" & vbLf
            Range("RowFail").Cells(i, 1) = 2
            fail = True
        ElseIf nMark = 0 Then
            If nBreakMaru > 0 Then
                If InStr(msg, "休憩時間帯にしか") = 0 Then _
                    msg = msg & "休憩時間帯にしか○のない生徒がいます。" & vbLf
                Range("RowFail").Cells(i, 1) = 2
                fail = True
            Else
                Range("RowFail").Cells(i, 1) = 1        '未入力（実施しない扱い）
            End If
        End If
    Next

    Application.ScreenUpdating = True
    If msg <> "" Then MsgBox msg & "変更が必要です。"
    If ss + kNumber > dd * kk Then
        MsgBox "コマ数が足りません（生徒 " & ss & " 名 + 休憩 " & kNumber & _
            " コマ > 全 " & dd * kk & " コマ）。" & vbLf & _
            "休憩を減らすか、初期設定でコマ数を増やしてください。"
    End If
    Exit Sub
Trouble:
    CleanFail
End Sub

'--------------------------------------------------------------
' 調整エンジン本体（増加路法による最大マッチング）
'   フェーズ1: ○のみ
'   フェーズ2: ○で入れなかった生徒だけ自分の△も使う
'   フェーズ3: 全員の△を使う（理論上の最大マッチング）
'   失敗時 : 原因グループを特定して表示
'--------------------------------------------------------------
Public Sub RunChosei()
    Dim r0 As Long, c0 As Long
    Dim wish(1 To MAXS, 1 To MAXSLOT) As Byte
    Dim blocked(1 To MAXSLOT) As Boolean
    Dim matchStu(1 To MAXSLOT) As Long
    Dim matchSlot(1 To MAXS) As Long
    Dim nijuAt(1 To MAXS) As Long
    Dim visited(1 To MAXSLOT) As Boolean
    Dim order(1 To MAXS) As Long
    Dim optCnt(1 To MAXS) As Long
    Dim nSlots As Long, np As Long
    Dim i As Long, j As Long, s As Long
    Dim v As String, cn As Long
    Dim phase As Long
    Dim total As Long, nSub As Long, nSkip As Long
    Dim anyW As Boolean
    Dim dummy As Boolean

    On Error GoTo Trouble
    EnsureDims
    If dd < 1 Or kk < 1 Or ss < 1 Or dd * kk > MAXSLOT Then
        MsgBox "初期設定が正しくありません。Start! シートで設定してください。"
        Exit Sub
    End If
    nSlots = dd * kk
    r0 = Range("FIRST").Row
    c0 = Range("FIRST").Column

    Application.ScreenUpdating = False
    ActiveSheet.Unprotect

    '--- 読み取り ----------------------------------------------
    For j = 1 To nSlots
        blocked(j) = (ActiveSheet.Cells(r0, c0 + j - 1).Interior.ColorIndex = kColor)
    Next
    cn = 0
    For s = 1 To ss
        For j = 1 To nSlots
            If Not blocked(j) Then
                v = CStr(ActiveSheet.Cells(r0 + s, c0 + j - 1).Value)
                If v = "○" Then
                    wish(s, j) = 1
                    cn = cn + 1
                ElseIf v = "△" Then
                    wish(s, j) = 2
                ElseIf v = "◎" Then
                    wish(s, j) = 1
                    cn = cn + 1
                    If nijuAt(s) = 0 Then nijuAt(s) = j
                End If
            End If
        Next
    Next

    '協力度（Start! の表示セル）
    On Error Resume Next
    Sheets("Start!").Range("協力度") = Format(cn / (dd * kk * ss) * 100, "##.0")
    On Error GoTo Trouble

    '--- ◎は先に確定 -------------------------------------------
    For s = 1 To ss
        j = nijuAt(s)
        If j > 0 Then
            If matchStu(j) = 0 Then
                matchStu(j) = s
                matchSlot(s) = j
            End If
        End If
    Next
    For j = 1 To nSlots
        If matchStu(j) > 0 Then blocked(j) = True
    Next

    '--- 不参加（無印）の生徒を除外 -----------------------------
    '（希望はあるが全部 ◎ に塞がれている生徒は除外せず、診断対象に残す）
    nSkip = 0
    For s = 1 To ss
        If matchSlot(s) = 0 Then
            anyW = False
            For j = 1 To nSlots
                If wish(s, j) > 0 Then
                    anyW = True
                    Exit For
                End If
            Next
            If Not anyW Then
                matchSlot(s) = -1
                nSkip = nSkip + 1
            End If
        End If
    Next

    '--- 希望の少ない生徒から処理（安定挿入ソート） -------------
    np = 0
    For s = 1 To ss
        If matchSlot(s) = 0 Then
            np = np + 1
            order(np) = s
            optCnt(s) = 0
            For j = 1 To nSlots
                If wish(s, j) = 1 And Not blocked(j) Then optCnt(s) = optCnt(s) + 1
            Next
        End If
    Next
    For i = 2 To np
        s = order(i)
        j = i - 1
        Do While j >= 1
            If optCnt(order(j)) > optCnt(s) Then
                order(j + 1) = order(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        order(j + 1) = s
    Next

    '--- 増加路マッチング（3段階） ------------------------------
    For phase = 1 To 3
        For i = 1 To np
            s = order(i)
            If matchSlot(s) = 0 Then
                Erase visited
                dummy = TryAugment(s, s, phase, wish, blocked, matchStu, matchSlot, visited, nSlots)
            End If
        Next
    Next

    '--- 結果の描画と裏データ書き込み ---------------------------
    total = 0
    nSub = 0
    For s = 1 To ss
        If matchSlot(s) > 0 Then
            j = matchSlot(s)
            If j = nijuAt(s) Then
                ActiveSheet.Cells(r0 + s, c0 + j - 1).Interior.ColorIndex = pColor
            Else
                ActiveSheet.Cells(r0 + s, c0 + j - 1).Interior.ColorIndex = mColor
            End If
            ActiveSheet.Cells(r0 + s, STATE_COL) = j
            If wish(s, j) = 2 Then nSub = nSub + 1
            total = total + 1
        Else
            ActiveSheet.Cells(r0 + s, STATE_COL).ClearContents
        End If
    Next
    ActiveSheet.Cells(r0, STATE_COL) = STATE_VER
    Application.CalculateFull

    '--- 結果表示 ----------------------------------------------
    If total = ss - nSkip Then
        v = total & " / " & (ss - nSkip) & " 名 全員確定しました!"
        If nSub > 0 Then v = v & vbLf & "（うち " & nSub & " 名は△の枠を使用しています）"
        If nSkip > 0 Then v = v & vbLf & "（未入力の " & nSkip & " 名は飛ばしました）"
        v = v & vbLf & vbLf & "これでよければ、メニューの作成ボタンから予定表・通知票を作成してください。" & _
            vbLf & "微調整はメニューの調整作業（入れ替えモード）でできます。"
        ActiveSheet.Protect AllowFormattingCells:=False
        Application.ScreenUpdating = True
        MsgBox v
    Else
        Diagnose wish, blocked, matchStu, matchSlot, order, np, nSlots, total, nSkip
        ActiveSheet.Protect AllowFormattingCells:=False
        Application.ScreenUpdating = True
    End If
    Exit Sub
Trouble:
    CleanFail
    On Error Resume Next
    ActiveSheet.Protect AllowFormattingCells:=False
End Sub

'--- 増加路を1本探す（再帰） -----------------------------------
Private Function TryAugment(ByVal s As Long, ByVal root As Long, ByVal phase As Long, _
        wish() As Byte, blocked() As Boolean, matchStu() As Long, matchSlot() As Long, _
        visited() As Boolean, ByVal nSlots As Long) As Boolean
    Dim j As Long
    Dim edgeOk As Boolean
    TryAugment = False
    For j = 1 To nSlots
        If Not blocked(j) And Not visited(j) Then
            edgeOk = (wish(s, j) = 1)
            If Not edgeOk Then
                If wish(s, j) = 2 Then
                    If phase = 3 Then edgeOk = True
                    If phase = 2 And s = root Then edgeOk = True
                End If
            End If
            If edgeOk Then
                visited(j) = True
                If matchStu(j) = 0 Then
                    matchStu(j) = s
                    matchSlot(s) = j
                    TryAugment = True
                    Exit Function
                ElseIf TryAugment(matchStu(j), root, phase, wish, blocked, matchStu, matchSlot, visited, nSlots) Then
                    matchStu(j) = s
                    matchSlot(s) = j
                    TryAugment = True
                    Exit Function
                End If
            End If
        End If
    Next
End Function

'--- 失敗の原因グループを特定して表示 ---------------------------
Private Sub Diagnose(wish() As Byte, blocked() As Boolean, matchStu() As Long, _
        matchSlot() As Long, order() As Long, ByVal np As Long, ByVal nSlots As Long, _
        ByVal total As Long, ByVal nSkip As Long)
    Dim grouped(1 To MAXS) As Boolean
    Dim inStu(1 To MAXS) As Boolean
    Dim inSlot(1 To MAXSLOT) As Boolean
    Dim stack(1 To MAXS) As Long
    Dim i As Long, j As Long, s As Long, t As Long, top As Long
    Dim nStu As Long, nSlot As Long
    Dim nms As String, msg As String

    msg = total & " / " & (ss - nSkip) & " 名しか確定できませんでした。" & vbLf & vbLf & _
        "次のグループは、希望枠の数が人数より少ないため全員は入れません:" & vbLf

    For i = 1 To np
        s = order(i)
        If matchSlot(s) = 0 And Not grouped(s) Then
            Erase inStu
            Erase inSlot
            top = 1
            stack(1) = s
            inStu(s) = True
            nStu = 1
            nSlot = 0
            Do While top > 0
                t = stack(top)
                top = top - 1
                For j = 1 To nSlots
                    If Not blocked(j) And Not inSlot(j) Then
                        If wish(t, j) > 0 Then
                            inSlot(j) = True
                            nSlot = nSlot + 1
                            If matchStu(j) > 0 Then
                                If Not inStu(matchStu(j)) Then
                                    inStu(matchStu(j)) = True
                                    nStu = nStu + 1
                                    top = top + 1
                                    stack(top) = matchStu(j)
                                End If
                            End If
                        End If
                    End If
                Next
            Loop
            nms = ""
            For t = 1 To ss
                If inStu(t) Then
                    grouped(t) = True
                    If nms <> "" Then nms = nms & ", "
                    nms = nms & CStr(Val(CStr(Sheets("名簿").Cells(1 + t, 1).Value)))
                    If matchSlot(t) = 0 Then Range("RowFail").Cells(t, 1) = 2
                End If
            Next
            msg = msg & "・出席番号 [" & nms & "] の " & nStu & " 名 → 使える枠は " & _
                nSlot & " コマ" & vbLf
        End If
    Next

    msg = msg & vbLf & "赤色表示の生徒の希望(○/△)を増やしてもらうか、" & vbLf & _
        "このグループの希望枠と重なる休憩・◎を見直して、再度調整してください。"
    MsgBox msg, vbExclamation
End Sub

'--------------------------------------------------------------
' 裏データと色の整合チェック / 修復
'--------------------------------------------------------------
Public Function StateIsConsistent() As Boolean
    Dim r0 As Long, c0 As Long, s As Long, j As Long
    Dim stateJ As Long, colorJ As Long, nCol As Long
    EnsureDims
    r0 = Range("FIRST").Row
    c0 = Range("FIRST").Column
    StateIsConsistent = True
    If CStr(ActiveSheet.Cells(r0, STATE_COL).Value) <> STATE_VER Then Exit Function
    For s = 1 To ss
        stateJ = CLng(Val(CStr(ActiveSheet.Cells(r0 + s, STATE_COL).Value)))
        colorJ = 0
        nCol = 0
        For j = 1 To dd * kk
            With ActiveSheet.Cells(r0 + s, c0 + j - 1).Interior
                If .ColorIndex = mColor Or .ColorIndex = pColor Then
                    nCol = nCol + 1
                    colorJ = j
                End If
            End With
        Next
        If nCol > 1 Or colorJ <> stateJ Then
            StateIsConsistent = False
            Exit Function
        End If
    Next
End Function

Public Sub RepaintFromState()
    Dim r0 As Long, c0 As Long, s As Long, j As Long, stateJ As Long
    On Error GoTo Trouble
    EnsureDims
    Application.ScreenUpdating = False
    ActiveSheet.Unprotect
    r0 = Range("FIRST").Row
    c0 = Range("FIRST").Column
    For s = 1 To ss
        For j = 1 To dd * kk
            With ActiveSheet.Cells(r0 + s, c0 + j - 1).Interior
                If .ColorIndex = mColor Or .ColorIndex = pColor Then .ColorIndex = xlNone
            End With
        Next
        stateJ = CLng(Val(CStr(ActiveSheet.Cells(r0 + s, STATE_COL).Value)))
        If stateJ >= 1 And stateJ <= dd * kk Then
            If CStr(ActiveSheet.Cells(r0 + s, c0 + stateJ - 1).Value) = "◎" Then
                ActiveSheet.Cells(r0 + s, c0 + stateJ - 1).Interior.ColorIndex = pColor
            Else
                ActiveSheet.Cells(r0 + s, c0 + stateJ - 1).Interior.ColorIndex = mColor
            End If
        End If
    Next
    Application.CalculateFull
    ActiveSheet.Protect AllowFormattingCells:=False
    Application.ScreenUpdating = True
    Exit Sub
Trouble:
    CleanFail
End Sub

'--------------------------------------------------------------
' 予定表シートの作成（旧 UserForm2 から移設・整理）
'   呼び出し時は調整シートがアクティブであること
'--------------------------------------------------------------
Public Function MakeScheduleCore() As String
    Dim n As Long
    Dim shObj As Object
    Dim src As String
    Dim rd As Long, cd As Long, rt As Long, ct As Long
    Dim rr As Long, cc As Long, rn As Long, cn2 As Long
    Dim d As Long, k As Long, s As Long
    Dim cs As Long, cextra As Long

    src = ActiveSheet.Name
    EnsureDims

    If CStr(ActiveSheet.Cells(Range("FIRST").Row, STATE_COL).Value) = STATE_VER Then
        If Not StateIsConsistent() Then
            If MsgBox("確定の色と内部データが一致していません（色の手修正など）。" & vbLf & _
                "内部データに合わせて塗り直してから作成しますか?", vbYesNo) = vbYes Then
                RepaintFromState
            End If
        End If
    End If

    n = 1
Again:
    For Each shObj In ThisWorkbook.Sheets
        If shObj.Name = "予定表" & CStr(n) Then
            n = n + 1
            GoTo Again
        End If
    Next

    Application.ScreenUpdating = False
    Sheets("元_予定表").Copy After:=Sheets(src)
    Sheets("元_予定表 (2)").Visible = True
    Sheets("元_予定表 (2)").Activate
    ActiveSheet.Name = "予定表" & CStr(n)
    ActiveSheet.Unprotect

    rd = Range("表日付").Row
    cd = Range("表日付").Column
    rt = Range("表時間").Row
    ct = Range("表時間").Column

    If dd < 5 Then                              '不要な列削除
        cs = dd * 3
        cextra = 2 + 3 * (4 - dd)
        ActiveSheet.Range(ActiveSheet.Cells(rd, cd + cs), _
            ActiveSheet.Cells(rd, cd + cs + cextra)).EntireColumn.Delete
    End If
    If kk < 30 Then                             '不要な行削除
        ActiveSheet.Range(ActiveSheet.Cells(rt + kk, ct), _
            ActiveSheet.Cells(rt + kk + (29 - kk), ct)).EntireRow.Delete
    End If

    rr = Range("FIRST").Row
    cc = Range("FIRST").Column
    For k = 1 To kk                             '時間転記
        ActiveSheet.Cells(rt + k - 1, ct) = Sheets(src).Cells(rr, cc + k - 1)
    Next
    For d = 0 To dd - 1                         '日付転記
        ActiveSheet.Cells(rd, cd + 3 * d) = Sheets(src).Cells(rr - 1, cc + kk * d)
    Next

    rn = Range("表番号").Row                    '生徒番号転記
    cn2 = Range("表番号").Column
    For d = 1 To dd
        For k = 1 To kk
            For s = 1 To ss
                With Sheets(src).Cells(rr + s, cc - 1 + k + kk * (d - 1)).Interior
                    If .ColorIndex = mColor Or .ColorIndex = pColor Then
                        ActiveSheet.Cells(rn + k, cn2 + 3 * (d - 1)) = s
                    End If
                End With
            Next
        Next
    Next

    ActiveSheet.Protect
    Application.ScreenUpdating = True
    MakeScheduleCore = "予定表" & CStr(n)
End Function

'--------------------------------------------------------------
' 保護者向け 個別通知票の作成
'   srcName: 調整シート名（確定済み）
'--------------------------------------------------------------
Public Function MakeNotifySheet(ByVal srcName As String) As String
    Dim n As Long
    Dim shObj As Object
    Dim ws As Worksheet, srcWs As Worksheet
    Dim r0 As Long, c0 As Long
    Dim s As Long, j As Long, d As Long, k As Long
    Dim slotOf(1 To MAXS) As Long
    Dim mins As Long
    Dim br As Long, bc As Long, idx As Long
    Dim dtv As Variant, tmv As Variant
    Dim dateTxt As String, timeTxt As String
    Dim nm As String, bn As String
    Dim blockRows As Long
    Dim perPage As Long

    Set srcWs = Sheets(srcName)
    EnsureDims
    r0 = Range("FIRST").Row
    c0 = Range("FIRST").Column
    mins = CLng(Val(CStr(Sheets("Start!").Range("時間").Value)))

    '確定コマの収集（色ベース）
    For s = 1 To ss
        slotOf(s) = 0
        For j = 1 To dd * kk
            With srcWs.Cells(r0 + s, c0 + j - 1).Interior
                If .ColorIndex = mColor Or .ColorIndex = pColor Then
                    slotOf(s) = j
                    Exit For
                End If
            End With
        Next
    Next

    n = 1
Again:
    For Each shObj In ThisWorkbook.Sheets
        If shObj.Name = "通知票" & CStr(n) Then
            n = n + 1
            GoTo Again
        End If
    Next

    Application.ScreenUpdating = False
    Set ws = ThisWorkbook.Worksheets.Add(After:=Sheets(srcName))
    ws.Name = "通知票" & CStr(n)

    'レイアウト: 2列 × 5段 = 10枚/ページ
    ws.Cells.Font.Size = 10
    ws.Columns(1).ColumnWidth = 1.5
    ws.Columns(2).ColumnWidth = 11
    ws.Columns(3).ColumnWidth = 13
    ws.Columns(4).ColumnWidth = 13
    ws.Columns(5).ColumnWidth = 9
    ws.Columns(6).ColumnWidth = 3
    ws.Columns(7).ColumnWidth = 11
    ws.Columns(8).ColumnWidth = 13
    ws.Columns(9).ColumnWidth = 13
    ws.Columns(10).ColumnWidth = 9
    ws.Columns(11).ColumnWidth = 1.5

    blockRows = 7
    perPage = 5

    For idx = 0 To ss - 1
        s = idx + 1
        br = 2 + (idx \ 2) * (blockRows + 1)
        If (idx \ 2) > 0 And ((idx \ 2) Mod perPage) = 0 And (idx Mod 2) = 0 Then
            ws.HPageBreaks.Add Before:=ws.Rows(br)
        End If
        bc = 2 + (idx Mod 2) * 5

        bn = CStr(Val(CStr(Sheets("名簿").Cells(1 + s, 1).Value)))
        nm = CStr(Sheets("名簿").Cells(1 + s, 2).Value) & "　" & CStr(Sheets("名簿").Cells(1 + s, 3).Value)

        ws.Rows(br).RowHeight = 20
        ws.Rows(br + 1).RowHeight = 6
        ws.Rows(br + 2).RowHeight = 18
        ws.Rows(br + 3).RowHeight = 22
        ws.Rows(br + 4).RowHeight = 18
        ws.Rows(br + 5).RowHeight = 16
        ws.Rows(br + 6).RowHeight = 14
        ws.Rows(br + blockRows).RowHeight = 12

        With ws.Range(ws.Cells(br, bc), ws.Cells(br, bc + 3))
            .Merge
            .Value = "保護者会のお知らせ"
            .HorizontalAlignment = xlCenter
            .Font.Bold = True
            .Font.Size = 12
        End With

        ws.Cells(br + 2, bc).Value = "番号・氏名"
        With ws.Range(ws.Cells(br + 2, bc + 1), ws.Cells(br + 2, bc + 3))
            .Merge
            .Value = bn & "番　" & nm & "　さん"
        End With

        ws.Cells(br + 3, bc).Value = "日　時"
        With ws.Range(ws.Cells(br + 3, bc + 1), ws.Cells(br + 3, bc + 3))
            .Merge
            j = slotOf(s)
            If j > 0 Then
                d = (j - 1) \ kk + 1
                k = (j - 1) Mod kk + 1
                dtv = srcWs.Cells(r0 - 1, c0 + kk * (d - 1)).Value
                tmv = srcWs.Cells(r0, c0 + j - 1).Value
                If IsDate(dtv) Then
                    dateTxt = Format(dtv, "m月d日(aaa)")
                Else
                    dateTxt = CStr(d) & "日目"
                End If
                If IsNumeric(tmv) Or IsDate(tmv) Then
                    timeTxt = Format(tmv, "h:mm") & " ～ " & Format(CDbl(tmv) + mins / 1440#, "h:mm")
                Else
                    timeTxt = CStr(tmv)
                End If
                .Value = dateTxt & "　" & timeTxt
                .Font.Bold = True
                .Font.Size = 12
            Else
                .Value = "未定（別途ご連絡します）"
            End If
        End With

        ws.Cells(br + 4, bc).Value = "場　所"
        With ws.Range(ws.Cells(br + 4, bc + 1), ws.Cells(br + 4, bc + 3))
            .Merge
            .Value = "教室（　　　　　　　　）"
        End With

        With ws.Range(ws.Cells(br + 5, bc), ws.Cells(br + 5, bc + 3))
            .Merge
            .Value = "※ご都合が悪い場合は担任までご連絡ください。"
            .Font.Size = 8
        End With

        With ws.Range(ws.Cells(br, bc), ws.Cells(br + 5, bc + 3))
            .BorderAround Weight:=xlThin
        End With
        With ws.Range(ws.Cells(br + 6, bc), ws.Cells(br + 6, bc + 3))
            .Merge
            .Value = "・・・・・・・・・・（キリトリ）・・・・・・・・・・"
            .HorizontalAlignment = xlCenter
            .Font.Size = 8
            .Font.ColorIndex = 16
        End With
    Next

    With ws.PageSetup
        .PaperSize = xlPaperA4
        .Orientation = xlPortrait
        .Zoom = 100
        .LeftMargin = Application.CentimetersToPoints(1)
        .RightMargin = Application.CentimetersToPoints(1)
        .TopMargin = Application.CentimetersToPoints(1)
        .BottomMargin = Application.CentimetersToPoints(1)
    End With

    ws.Activate
    Application.ScreenUpdating = True
    MakeNotifySheet = ws.Name
End Function

'--------------------------------------------------------------
' シートを PDF に保存
'--------------------------------------------------------------
Public Sub ExportSheetPDF(ByVal wsName As String)
    Dim p As String, f As String
    On Error GoTo Trouble
    p = ThisWorkbook.Path
    If p = "" Or LCase(Left(p, 4)) = "http" Then
        p = Environ("USERPROFILE") & "\Desktop"
    End If
    f = p & Application.PathSeparator & wsName & "_" & Format(Now, "yyyymmdd_hhmm") & ".pdf"
    Sheets(wsName).ExportAsFixedFormat Type:=xlTypePDF, Filename:=f, _
        Quality:=xlQualityStandard, OpenAfterPublish:=True
    MsgBox "PDF を保存しました:" & vbLf & f
    Exit Sub
Trouble:
    MsgBox "PDF の保存に失敗しました。" & vbLf & Err.Description, vbExclamation
End Sub

'--------------------------------------------------------------
' ユーザー定義関数（シート57行目の式から使用・ver1.11 互換）
'--------------------------------------------------------------
Function CellColor(ByVal rng As Range)
    Dim cc As Range
    CellColor = 0
    For Each cc In rng
        With cc.Interior
            If .ColorIndex = mColor Or .ColorIndex = pColor Then
                CellColor = CellColor + 1
            ElseIf .ColorIndex = kColor Then
                CellColor = 51
                Exit For
            End If
        End With
    Next
End Function

Function MyPrev(ByVal c As Variant)
    '休憩時間帯なら kColor、そうでなければ無色を返す
    If ActiveSheet.Cells(Range("FIRST").Row, c).Interior.ColorIndex = kColor Then
        MyPrev = kColor
    Else
        MyPrev = xlNone
    End If
End Function
