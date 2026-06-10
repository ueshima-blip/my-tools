Attribute VB_Name = "Module2"
Option Explicit

'==============================================================
' 保護者会調整 ver2.0  取り込み・調査票・クイック入力
'   ・ImportWishes    : CSV / Excel（Googleフォーム・Microsoft Forms・
'                       OMR読み取りアプリ出力）から希望を取り込む
'   ・MakeSurveySheets: 紙配布用のマークシート式調査票を作成
'                       （OMRアプリ用の定義ファイルも書き出す）
'   ・QuickInput      : 紙回収分をキーボードで高速入力
'==============================================================

'--------------------------------------------------------------
' 全角数字などを半角へ（日付・時刻の解析用）
'--------------------------------------------------------------
Private Function Z2H(ByVal s As String) As String
    Dim i As Long, ch As String, p As Long
    Const ZEN As String = "０１２３４５６７８９：／－〜（）［］"
    Const HAN As String = "0123456789:/-~()[]"
    Dim out As String
    out = ""
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        p = InStr(ZEN, ch)
        If p > 0 Then ch = Mid(HAN, p, 1)
        If ch <> " " And ch <> "　" Then out = out & ch
    Next
    Z2H = out
End Function

'--------------------------------------------------------------
' 文字列から「月/日」と「時:分」を取り出す
'   "6/20 15:30" "６月２０日(金) １５時３０分" "希望 [6/20 15:30]" など
'--------------------------------------------------------------
Private Function ParseSlotLabel(ByVal s As String, ByRef mo As Long, _
        ByRef dy As Long, ByRef hh As Long, ByRef mm As Long) As Boolean
    Dim t As String
    Dim i As Long, p As Long
    Dim n1 As Long, n2 As Long
    Dim sep As String
    Dim got2 As Boolean

    ParseSlotLabel = False
    t = Z2H(s)
    '角かっこがあれば中身だけを見る（Googleフォームのグリッド形式）
    p = InStr(t, "[")
    If p > 0 Then
        i = InStr(p, t, "]")
        If i > p Then t = Mid(t, p + 1, i - p - 1)
    End If
    '曜日かっこを除去
    t = Replace(t, "(月)", ""): t = Replace(t, "(火)", "")
    t = Replace(t, "(水)", ""): t = Replace(t, "(木)", "")
    t = Replace(t, "(金)", ""): t = Replace(t, "(土)", "")
    t = Replace(t, "(日)", "")
    '表記ゆれを正規化
    t = Replace(t, "月", "/")
    t = Replace(t, "時半", ":30")
    t = Replace(t, "時", ":")
    t = Replace(t, "分", "")
    t = Replace(t, "日", "")

    mo = 0: dy = 0: hh = -1: mm = -1
    i = 1
    Do While i <= Len(t)
        If Mid(t, i, 1) >= "0" And Mid(t, i, 1) <= "9" Then
            n1 = 0
            Do While i <= Len(t)
                If Mid(t, i, 1) >= "0" And Mid(t, i, 1) <= "9" Then
                    n1 = n1 * 10 + CLng(Mid(t, i, 1))
                    i = i + 1
                Else
                    Exit Do
                End If
            Loop
            If i <= Len(t) Then
                If Mid(t, i, 1) = "/" Or Mid(t, i, 1) = ":" Then
                    sep = Mid(t, i, 1)
                    i = i + 1
                    n2 = 0
                    got2 = False
                    Do While i <= Len(t)
                        If Mid(t, i, 1) >= "0" And Mid(t, i, 1) <= "9" Then
                            n2 = n2 * 10 + CLng(Mid(t, i, 1))
                            i = i + 1
                            got2 = True
                        Else
                            Exit Do
                        End If
                    Loop
                    If sep = "/" And mo = 0 And got2 Then
                        If n1 >= 1 And n1 <= 12 And n2 >= 1 And n2 <= 31 Then
                            mo = n1
                            dy = n2
                        End If
                    ElseIf sep = ":" And hh < 0 Then
                        If n1 >= 0 And n1 <= 23 And got2 And n2 >= 0 And n2 <= 59 Then
                            hh = n1
                            mm = n2
                        ElseIf n1 >= 0 And n1 <= 23 And Not got2 Then
                            hh = n1
                            mm = 0
                        End If
                    End If
                End If
            End If
        Else
            i = i + 1
        End If
    Loop
    If mo > 0 And hh >= 0 Then ParseSlotLabel = True
End Function

'--------------------------------------------------------------
' 希望入力シートのヘッダから全コマの (月,日,時,分,ラベル) を作る
'   戻り値: 有効コマ数（0=日付や時間が未設定）
'--------------------------------------------------------------
Private Function BuildSlots(slotMo() As Long, slotDy() As Long, _
        slotHH() As Long, slotMM() As Long, slotLabel() As String) As Long
    Dim r0 As Long, c0 As Long
    Dim d As Long, k As Long, j As Long
    Dim dtv As Variant, tmv As Variant
    Dim ws As Worksheet

    Set ws = Sheets("希望入力")
    r0 = ws.Range("D2").Row
    c0 = ws.Range("D2").Column
    BuildSlots = 0

    For d = 1 To dd
        dtv = ws.Cells(r0 - 1, c0 + kk * (d - 1)).Value
        If Not IsDate(dtv) Then Exit Function
        For k = 1 To kk
            j = (d - 1) * kk + k
            tmv = ws.Cells(r0, c0 + j - 1).Value
            If Not (IsNumeric(tmv) Or IsDate(tmv)) Then Exit Function
            slotMo(j) = Month(CDate(dtv))
            slotDy(j) = Day(CDate(dtv))
            slotHH(j) = Hour(CDbl(tmv))
            slotMM(j) = Minute(CDbl(tmv))
            slotLabel(j) = slotMo(j) & "/" & slotDy(j) & " " & _
                Format(slotHH(j), "0") & ":" & Format(slotMM(j), "00")
        Next
    Next
    BuildSlots = dd * kk
End Function

Private Function FindSlot(ByVal mo As Long, ByVal dy As Long, ByVal hh As Long, _
        ByVal mm As Long, slotMo() As Long, slotDy() As Long, _
        slotHH() As Long, slotMM() As Long) As Long
    Dim j As Long
    FindSlot = 0
    For j = 1 To dd * kk
        If slotMo(j) = mo And slotDy(j) = dy And slotHH(j) = hh And slotMM(j) = mm Then
            FindSlot = j
            Exit Function
        End If
    Next
End Function

'--------------------------------------------------------------
' 名簿から生徒を探す（出席番号優先・なければ氏名）
'--------------------------------------------------------------
Private Function FindStudentByNo(ByVal v As Variant) As Long
    Dim s As Long, n As Long
    FindStudentByNo = 0
    n = CLng(Val(CStr(v)))
    If n < 1 Then Exit Function
    For s = 1 To ss
        If CLng(Val(CStr(Sheets("名簿").Cells(1 + s, 1).Value))) = n Then
            FindStudentByNo = s
            Exit Function
        End If
    Next
End Function

Private Function NameKey(ByVal s As String) As String
    NameKey = Replace(Replace(CStr(s), " ", ""), "　", "")
End Function

Private Function FindStudentByName(ByVal v As String) As Long
    Dim s As Long, key As String
    FindStudentByName = 0
    key = NameKey(v)
    If key = "" Then Exit Function
    For s = 1 To ss
        If NameKey(CStr(Sheets("名簿").Cells(1 + s, 2).Value) & _
                CStr(Sheets("名簿").Cells(1 + s, 3).Value)) = key Then
            FindStudentByName = s
            Exit Function
        End If
    Next
End Function

'--------------------------------------------------------------
' テキストファイルを文字コード自動判定で読む（UTF-8 / Shift_JIS）
'--------------------------------------------------------------
Private Function ReadTextFile(ByVal path As String) As String
    Dim stm As Object
    Dim bytes() As Byte
    Dim cs As String

    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1                'binary
    stm.Open
    stm.LoadFromFile path
    bytes = stm.Read
    stm.Close
    Set stm = Nothing

    cs = "shift-jis"
    If UBound(bytes) >= 2 Then
        If bytes(0) = &HEF And bytes(1) = &HBB And bytes(2) = &HBF Then
            cs = "utf-8"
        ElseIf LooksUtf8(bytes) Then
            cs = "utf-8"
        End If
    End If

    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2                'text
    stm.Charset = cs
    stm.Open
    stm.LoadFromFile path
    ReadTextFile = stm.ReadText(-1)
    stm.Close
    Set stm = Nothing

    'BOM が文字として残った場合は除去
    If Len(ReadTextFile) > 0 Then
        If AscW(Left(ReadTextFile, 1)) = -257 Or AscW(Left(ReadTextFile, 1)) = 65279 Then
            ReadTextFile = Mid(ReadTextFile, 2)
        End If
    End If
End Function

Private Function LooksUtf8(bytes() As Byte) As Boolean
    'UTF-8 として正しく、かつマルチバイト文字を含むか
    Dim i As Long, b As Long, follow As Long, hasMulti As Boolean
    LooksUtf8 = False
    hasMulti = False
    i = 0
    Do While i <= UBound(bytes)
        b = bytes(i)
        If b < &H80 Then
            follow = 0
        ElseIf b >= &HC2 And b <= &HDF Then
            follow = 1
        ElseIf b >= &HE0 And b <= &HEF Then
            follow = 2
        ElseIf b >= &HF0 And b <= &HF4 Then
            follow = 3
        Else
            Exit Function
        End If
        If follow > 0 Then
            hasMulti = True
            Do While follow > 0
                i = i + 1
                If i > UBound(bytes) Then Exit Function
                If bytes(i) < &H80 Or bytes(i) > &HBF Then Exit Function
                follow = follow - 1
            Loop
        End If
        i = i + 1
    Loop
    LooksUtf8 = hasMulti
End Function

'--------------------------------------------------------------
' CSV パーサ（RFC4180: 引用符・改行・カンマ対応）
'--------------------------------------------------------------
Private Function ParseCsv(ByVal text As String) As Collection
    Dim rows As New Collection
    Dim fields(1 To 256) As String
    Dim nf As Long
    Dim cur As String
    Dim inQ As Boolean
    Dim i As Long, ch As String

    nf = 0
    cur = ""
    inQ = False
    i = 1
    Do While i <= Len(text)
        ch = Mid(text, i, 1)
        If inQ Then
            If ch = """" Then
                If Mid(text, i + 1, 1) = """" Then
                    cur = cur & """"
                    i = i + 1
                Else
                    inQ = False
                End If
            Else
                cur = cur & ch
            End If
        Else
            Select Case ch
            Case """"
                inQ = True
            Case ","
                If nf < 256 Then nf = nf + 1
                fields(nf) = cur
                cur = ""
            Case vbCr
                If Mid(text, i + 1, 1) = vbLf Then i = i + 1
                If nf < 256 Then nf = nf + 1
                fields(nf) = cur
                cur = ""
                rows.Add CopyFields(fields, nf)
                nf = 0
            Case vbLf
                If nf < 256 Then nf = nf + 1
                fields(nf) = cur
                cur = ""
                rows.Add CopyFields(fields, nf)
                nf = 0
            Case Else
                cur = cur & ch
            End Select
        End If
        i = i + 1
    Loop
    If cur <> "" Or nf > 0 Then
        If nf < 256 Then nf = nf + 1
        fields(nf) = cur
        rows.Add CopyFields(fields, nf)
    End If
    Set ParseCsv = rows
End Function

Private Function CopyFields(fields() As String, ByVal nf As Long) As Variant
    Dim out() As String
    Dim i As Long
    If nf < 1 Then nf = 1
    ReDim out(1 To nf)
    For i = 1 To nf
        out(i) = fields(i)
    Next
    CopyFields = out
End Function

'--------------------------------------------------------------
' 取り込みファイルを 2次元配列 (1..nr, 1..nc) に読み込む
'--------------------------------------------------------------
Private Function LoadTable(ByVal path As String, ByRef nr As Long, ByRef nc As Long) As Variant
    Dim arr() As Variant
    Dim i As Long, j As Long

    nr = 0
    nc = 0
    If LCase(Right(path, 4)) = ".csv" Or LCase(Right(path, 4)) = ".txt" Then
        Dim rows As Collection
        Dim f As Variant
        Set rows = ParseCsv(ReadTextFile(path))
        nr = rows.Count
        For i = 1 To nr
            f = rows(i)
            If UBound(f) > nc Then nc = UBound(f)
        Next
        If nr = 0 Or nc = 0 Then Exit Function
        ReDim arr(1 To nr, 1 To nc)
        For i = 1 To nr
            f = rows(i)
            For j = 1 To UBound(f)
                arr(i, j) = f(j)
            Next
        Next
    Else
        Dim wb As Workbook
        Dim v As Variant
        Application.ScreenUpdating = False
        Set wb = Workbooks.Open(path, ReadOnly:=True, UpdateLinks:=0)
        v = wb.Sheets(1).UsedRange.Value
        wb.Close SaveChanges:=False
        Application.ScreenUpdating = True
        If IsArray(v) Then
            nr = UBound(v, 1)
            nc = UBound(v, 2)
            ReDim arr(1 To nr, 1 To nc)
            For i = 1 To nr
                For j = 1 To nc
                    arr(i, j) = v(i, j)
                Next
            Next
        Else
            nr = 1
            nc = 1
            ReDim arr(1 To 1, 1 To 1)
            arr(1, 1) = v
        End If
    End If
    LoadTable = arr
End Function

'--------------------------------------------------------------
' 希望の取り込み（メイン）
'--------------------------------------------------------------
Public Sub ImportWishes()
    Dim f As Variant
    Dim arr As Variant
    Dim nr As Long, nc As Long
    Dim slotMo(1 To MAXSLOT) As Long, slotDy(1 To MAXSLOT) As Long
    Dim slotHH(1 To MAXSLOT) As Long, slotMM(1 To MAXSLOT) As Long
    Dim slotLabel(1 To MAXSLOT) As String
    Dim colSlot(1 To 256) As Long       '列→コマ番号（マトリクス形式）
    Dim colSub(1 To 256) As Boolean     'その列は△扱いか
    Dim idCol As Long, nameCol As Long
    Dim multiCols(1 To 16) As Long      '複数選択形式の列
    Dim multiSub(1 To 16) As Boolean
    Dim nMulti As Long
    Dim r As Long, c As Long, j As Long, s As Long
    Dim h As String, v As String
    Dim mo As Long, dy As Long, hh As Long, mm As Long
    Dim marks(1 To MAXS, 1 To MAXSLOT) As Byte    '1=○ 2=△ 3=◎
    Dim hit(1 To MAXS) As Boolean
    Dim nHit As Long, nUnknown As Long
    Dim unknownTxt As String, badLabelTxt As String
    Dim headerRow As Long
    Dim ws As Worksheet
    Dim r0 As Long, c0 As Long
    Dim mode As Long
    Dim cnt As Long
    Dim allEmpty As Boolean
    Dim toks() As String
    Dim tk As Long, m As Long
    Dim clearFirst As Boolean
    Dim sampleEnd As Long

    EnsureDims
    If ss = 0 Then
        MsgBox "先に「名簿」シートへ名簿を貼り付けてください。"
        Exit Sub
    End If
    If dd * kk = 0 Then
        MsgBox "先に Start! シートで初期設定をしてください。"
        Exit Sub
    End If
    If BuildSlots(slotMo, slotDy, slotHH, slotMM, slotLabel) = 0 Then
        MsgBox "希望入力シートの日付・時間が設定されていません。" & vbLf & _
            "初期設定をして、日付（例: 6/20）を入力してから取り込んでください。"
        Exit Sub
    End If

    f = Application.GetOpenFilename( _
        "取込ファイル (*.csv;*.xlsx;*.xlsm;*.txt),*.csv;*.xlsx;*.xlsm;*.txt", _
        , "希望ファイルの選択（Googleフォーム/Forms/OMR読み取りアプリのファイル）")
    If VarType(f) = vbBoolean Then Exit Sub

    arr = LoadTable(CStr(f), nr, nc)
    If nr < 2 Or nc < 2 Then
        MsgBox "データが見つかりませんでした。"
        Exit Sub
    End If
    If nc > 256 Then nc = 256

    'ヘッダ行: 最初の非空行
    headerRow = 1
    For r = 1 To nr
        For c = 1 To nc
            If Trim(CStr(arr(r, c))) <> "" Then
                headerRow = r
                GoTo HeaderFound
            End If
        Next
    Next
HeaderFound:

    '--- 列の分類 ----------------------------------------------
    idCol = 0
    nameCol = 0
    nMulti = 0
    For c = 1 To nc
        h = CStr(arr(headerRow, c))
        colSlot(c) = 0
        colSub(c) = False
        If ParseSlotLabel(h, mo, dy, hh, mm) Then
            j = FindSlot(mo, dy, hh, mm, slotMo, slotDy, slotHH, slotMM)
            If j > 0 Then
                colSlot(c) = j
                If InStr(h, "第2") > 0 Or InStr(h, "第２") > 0 Or InStr(h, "△") > 0 Then
                    colSub(c) = True
                End If
            Else
                If Len(badLabelTxt) < 400 Then badLabelTxt = badLabelTxt & "・" & h & vbLf
            End If
        ElseIf InStr(h, "番号") > 0 And InStr(h, "電話") = 0 And idCol = 0 Then
            idCol = c
        ElseIf (InStr(h, "氏名") > 0 Or InStr(h, "名前") > 0 Or InStr(h, "なまえ") > 0) _
                And nameCol = 0 Then
            nameCol = c
        End If
    Next

    '複数選択（セル内にコマのラベルが入る）形式の列を探す
    sampleEnd = nr
    If sampleEnd > headerRow + 20 Then sampleEnd = headerRow + 20
    For c = 1 To nc
        If colSlot(c) = 0 And c <> idCol And c <> nameCol Then
            For r = headerRow + 1 To sampleEnd
                v = CStr(arr(r, c))
                If ParseSlotLabel(v, mo, dy, hh, mm) Then
                    If nMulti < 16 Then
                        nMulti = nMulti + 1
                        multiCols(nMulti) = c
                        h = CStr(arr(headerRow, c))
                        multiSub(nMulti) = (InStr(h, "第2") > 0 Or InStr(h, "第２") > 0 _
                            Or InStr(h, "△") > 0 Or InStr(h, "どうしても") > 0)
                    End If
                    Exit For
                End If
            Next
        End If
    Next

    mode = 0
    For c = 1 To nc
        If colSlot(c) > 0 Then mode = 1     'マトリクス形式
    Next
    If mode = 0 And nMulti > 0 Then mode = 2
    If mode = 0 Then
        MsgBox "コマ（日付+時刻）に対応する列が見つかりませんでした。" & vbLf & _
            "ヘッダまたは回答に「6/20 15:30」のような表記が必要です。" & vbLf & _
            "（フォームの選択肢名は『月/日 時:分』形式にしてください）"
        Exit Sub
    End If
    If idCol = 0 And nameCol = 0 Then
        MsgBox "生徒を特定する列（出席番号 または 氏名）が見つかりませんでした。"
        Exit Sub
    End If

    '--- 各行を生徒へ割り当て ----------------------------------
    nHit = 0
    nUnknown = 0
    For r = headerRow + 1 To nr
        allEmpty = True
        For c = 1 To nc
            If Trim(CStr(arr(r, c))) <> "" Then
                allEmpty = False
                Exit For
            End If
        Next
        If allEmpty Then GoTo NextRow

        s = 0
        If idCol > 0 Then s = FindStudentByNo(arr(r, idCol))
        If s = 0 And nameCol > 0 Then s = FindStudentByName(CStr(arr(r, nameCol)))
        If s = 0 Then
            nUnknown = nUnknown + 1
            If nUnknown <= 5 Then
                v = ""
                If idCol > 0 Then v = v & CStr(arr(r, idCol)) & " "
                If nameCol > 0 Then v = v & CStr(arr(r, nameCol))
                unknownTxt = unknownTxt & "・" & r & "行目: " & v & vbLf
            End If
            GoTo NextRow
        End If
        If Not hit(s) Then
            hit(s) = True
            nHit = nHit + 1
        Else
            '同じ生徒の2回目以降の回答 → 上書き（最新を優先）
            For j = 1 To dd * kk
                marks(s, j) = 0
            Next
        End If

        If mode = 1 Then
            For c = 1 To nc
                If colSlot(c) > 0 Then
                    v = Trim(CStr(arr(r, c)))
                    If v <> "" Then
                        j = colSlot(c)
                        If v = "◎" Then
                            marks(s, j) = 3
                        ElseIf v = "△" Or colSub(c) Then
                            If marks(s, j) = 0 Then marks(s, j) = 2
                        ElseIf v = "○" Or v = "TRUE" Or v = "True" Or v = "true" _
                                Or v = "はい" Or v = "1" Or v = "可" Or v = "Yes" Then
                            marks(s, j) = 1
                        End If
                    End If
                End If
            Next
        Else
            For m = 1 To nMulti
                v = CStr(arr(r, multiCols(m)))
                v = Replace(v, "、", ",")
                v = Replace(v, ";", ",")
                v = Replace(v, vbLf, ",")
                toks = Split(v, ",")
                For tk = 0 To UBound(toks)
                    If Trim(toks(tk)) <> "" Then
                        If ParseSlotLabel(toks(tk), mo, dy, hh, mm) Then
                            j = FindSlot(mo, dy, hh, mm, slotMo, slotDy, slotHH, slotMM)
                            If j > 0 Then
                                If multiSub(m) Then
                                    If marks(s, j) = 0 Then marks(s, j) = 2
                                Else
                                    marks(s, j) = 1
                                End If
                            ElseIf InStr(badLabelTxt, Trim(toks(tk))) = 0 Then
                                If Len(badLabelTxt) < 400 Then _
                                    badLabelTxt = badLabelTxt & "・" & Trim(toks(tk)) & vbLf
                            End If
                        End If
                    End If
                Next
            Next
        End If
NextRow:
    Next

    If nHit = 0 Then
        MsgBox "取り込める生徒が1人も見つかりませんでした。" & vbLf & unknownTxt
        Exit Sub
    End If

    '--- 確認して書き込み --------------------------------------
    v = nHit & " 名分の希望を読み取りました。"
    If nUnknown > 0 Then
        v = v & vbLf & "※名簿と一致しない行が " & nUnknown & " 行ありました:" & vbLf & unknownTxt
    End If
    If badLabelTxt <> "" Then
        v = v & vbLf & "※コマに対応しないラベル:" & vbLf & badLabelTxt
    End If
    v = v & vbLf & "希望入力シートへ書き込みますか?"
    If MsgBox(v, vbOKCancel) = vbCancel Then Exit Sub

    clearFirst = (MsgBox("既存の ○◎△ を消してから取り込みますか?" & vbLf & _
        "はい = 取り込んだ生徒の行を置き換え / いいえ = 追記（既存マークを残す）", _
        vbYesNo) = vbYes)

    Set ws = Sheets("希望入力")
    r0 = ws.Range("D2").Row
    c0 = ws.Range("D2").Column
    Application.ScreenUpdating = False
    ws.Unprotect
    cnt = 0
    For s = 1 To ss
        If hit(s) Then
            If clearFirst Then
                For j = 1 To dd * kk
                    v = CStr(ws.Cells(r0 + s, c0 + j - 1).Value)
                    If v = "○" Or v = "◎" Or v = "△" Then
                        ws.Cells(r0 + s, c0 + j - 1).ClearContents
                    End If
                Next
            End If
            For j = 1 To dd * kk
                If marks(s, j) > 0 Then
                    v = CStr(ws.Cells(r0 + s, c0 + j - 1).Value)
                    If v = "◎" Then
                        '手入力の◎を尊重して何もしない
                    ElseIf marks(s, j) = 3 Then
                        ws.Cells(r0 + s, c0 + j - 1) = "◎"
                        cnt = cnt + 1
                    ElseIf marks(s, j) = 2 Then
                        If v = "" Then
                            ws.Cells(r0 + s, c0 + j - 1) = "△"
                            cnt = cnt + 1
                        End If
                    Else
                        If v <> "○" Then cnt = cnt + 1
                        ws.Cells(r0 + s, c0 + j - 1) = "○"
                    End If
                End If
            Next
        End If
    Next
    ws.Protect AllowFormattingCells:=False
    ws.Activate
    Application.ScreenUpdating = True

    KomaCheck
    MsgBox "取り込みました（" & nHit & " 名 / " & cnt & " マーク）。" & vbLf & _
        "内容を確認して、メニューから調整を実行してください。"
End Sub

'--------------------------------------------------------------
' マークシート式 調査票の作成 + OMR用定義ファイルの書き出し
'
' 1ページ（=1人）の行構成（top = ページ先頭行）:
'   top+0 : 上マーカー行（B列・右列に黒四角）
'   top+1 : タイトル + 氏名
'   top+2 : 説明文
'   top+3 : 日付見出し
'   top+4 .. top+3+kk : マーク格子（時刻ラベル + 日列）
'   top+4+kk : 下マーカー行
'   top+5+kk : 注意書き
'   top+6+kk : 空白（ページ区切り）
' 出席番号は B列の top+1..top+8 に 8ビット縦帯
' （b6..b0 + 奇数パリティ）で印字して機械読み取りする。
'--------------------------------------------------------------
Public Sub MakeSurveySheets()
    Dim ws As Worksheet
    Dim slotMo(1 To MAXSLOT) As Long, slotDy(1 To MAXSLOT) As Long
    Dim slotHH(1 To MAXSLOT) As Long, slotMM(1 To MAXSLOT) As Long
    Dim slotLabel(1 To MAXSLOT) As String
    Dim s As Long, d As Long, k As Long, j As Long
    Dim pageRows As Long, top As Long
    Dim cM1 As Long, cT As Long, cD0 As Long, cM2 As Long
    Dim r As Long
    Dim num As Long, parity As Long
    Dim nm As String
    Dim gridH As Double, dayW As Double

    EnsureDims
    If ss = 0 Or dd * kk = 0 Then
        MsgBox "名簿と初期設定を先に済ませてください。"
        Exit Sub
    End If
    If kk < 5 Then
        MsgBox "コマ数が少なすぎます（1日5コマ以上で利用できます）。"
        Exit Sub
    End If
    If BuildSlots(slotMo, slotDy, slotHH, slotMM, slotLabel) = 0 Then
        MsgBox "希望入力シートに日付が入っていません。" & vbLf & _
            "初期設定 → 日付入力（例: 6/20）のあとに作成してください。"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error Resume Next
    Sheets("調査票印刷").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set ws = ThisWorkbook.Worksheets.Add(After:=Sheets("Start!"))
    ws.Name = "調査票印刷"

    '列構成: A=余白 B=左マーカー/ID帯 C=時刻 D..=日列 右端=マーカー
    cM1 = 2
    cT = 3
    cD0 = 4
    cM2 = cD0 + dd

    dayW = 340 / dd
    If dayW > 110 Then dayW = 110

    ws.Cells.Font.Size = 9
    ws.Columns(1).ColumnWidth = 0.8
    SetColWidthPts ws, cM1, 18
    SetColWidthPts ws, cT, 52
    For d = 0 To dd - 1
        SetColWidthPts ws, cD0 + d, dayW
    Next
    SetColWidthPts ws, cM2, 18

    gridH = 540 / kk
    If gridH > 20 Then gridH = 20
    If gridH < 13 Then gridH = 13

    pageRows = kk + 7

    For s = 1 To ss
        top = 1 + (s - 1) * pageRows
        num = CLng(Val(CStr(Sheets("名簿").Cells(1 + s, 1).Value)))
        If num < 1 Or num > 99 Then num = s
        nm = CStr(Sheets("名簿").Cells(1 + s, 2).Value) & "　" & _
             CStr(Sheets("名簿").Cells(1 + s, 3).Value)

        '上マーカー
        ws.Rows(top).RowHeight = 16
        ws.Cells(top, cM1).Interior.Color = RGB(0, 0, 0)
        ws.Cells(top, cM2).Interior.Color = RGB(0, 0, 0)

        'タイトル + 氏名
        ws.Rows(top + 1).RowHeight = 20
        With ws.Range(ws.Cells(top + 1, cT), ws.Cells(top + 1, cM2 - 1))
            .Merge
            .Value = "保護者会 希望調査票　　" & num & "番　" & nm & "　様"
            .Font.Size = 11
            .Font.Bold = True
        End With

        '説明
        ws.Rows(top + 2).RowHeight = 15
        With ws.Range(ws.Cells(top + 2, cT), ws.Cells(top + 2, cM2 - 1))
            .Merge
            .Value = "ご都合のよい時間枠を ■ のように黒く塗りつぶしてください（できるだけ多く）。"
            .Font.Size = 9
        End With

        '日付見出し
        ws.Rows(top + 3).RowHeight = 15
        For d = 1 To dd
            With ws.Cells(top + 3, cD0 + d - 1)
                .Value = slotMo((d - 1) * kk + 1) & "/" & slotDy((d - 1) * kk + 1)
                .HorizontalAlignment = xlCenter
                .Font.Bold = True
            End With
        Next
        With ws.Cells(top + 3, cT)
            .Value = "時刻"
            .HorizontalAlignment = xlCenter
        End With

        'マーク格子
        For k = 1 To kk
            r = top + 3 + k
            ws.Rows(r).RowHeight = gridH
            With ws.Cells(r, cT)
                .Value = "(" & k & ") " & Format(slotHH(k), "0") & ":" & Format(slotMM(k), "00")
                .HorizontalAlignment = xlRight
                .Font.Size = 9
            End With
        Next
        With ws.Range(ws.Cells(top + 4, cD0), ws.Cells(top + 3 + kk, cD0 + dd - 1))
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With

        'IDビット帯（B列 top+1..top+8: b6..b0 + 奇数パリティ）
        parity = 0
        For j = 6 To 0 Step -1
            If (num And CLng(2 ^ j)) <> 0 Then
                ws.Cells(top + 1 + (6 - j), cM1).Interior.Color = RGB(0, 0, 0)
                parity = parity + 1
            End If
        Next
        If (parity Mod 2) = 0 Then
            ws.Cells(top + 8, cM1).Interior.Color = RGB(0, 0, 0)
        End If

        '下マーカー
        r = top + 4 + kk
        ws.Rows(r).RowHeight = 16
        ws.Cells(r, cM1).Interior.Color = RGB(0, 0, 0)
        ws.Cells(r, cM2).Interior.Color = RGB(0, 0, 0)

        '注意書き
        ws.Rows(r + 1).RowHeight = 13
        With ws.Range(ws.Cells(r + 1, cT), ws.Cells(r + 1, cM2 - 1))
            .Merge
            .Value = "※濃い鉛筆かペンで枠内を塗る。黒い四角と左端の帯には書き込まないでください。"
            .Font.Size = 8
        End With
        ws.Rows(r + 2).RowHeight = 14

        If s < ss Then ws.HPageBreaks.Add Before:=ws.Rows(top + pageRows)
    Next

    With ws.PageSetup
        .PaperSize = xlPaperA4
        .Orientation = xlPortrait
        .Zoom = 100
        .LeftMargin = Application.CentimetersToPoints(1.2)
        .RightMargin = Application.CentimetersToPoints(1.2)
        .TopMargin = Application.CentimetersToPoints(1.2)
        .BottomMargin = Application.CentimetersToPoints(1.2)
        .HeaderMargin = Application.CentimetersToPoints(0.5)
        .FooterMargin = Application.CentimetersToPoints(0.5)
        .CenterHorizontally = False
        .PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(ss * pageRows, cM2)).Address
    End With

    WriteOmrDef ws, slotLabel, cM1, cT, cD0, cM2

    ws.Activate
    Application.ScreenUpdating = True
    MsgBox "調査票（" & ss & " 名分・1人1枚）と読み取り定義ファイル『調査票定義.csv』を作成しました。" & vbLf & vbLf & _
        "・印刷は必ず【100%（実際のサイズ）】で。「ページに合わせる」は使わないでください。" & vbLf & _
        "・回収後はスキャン（PDF または JPG）し、OMR読み取りアプリで CSV にして、" & vbLf & _
        "　Start! の「希望を取り込む」ボタンで取り込みます。"
End Sub

'列幅をポイント指定へ近づける（3回の補正で十分収束する）
Private Sub SetColWidthPts(ws As Worksheet, ByVal col As Long, ByVal pts As Double)
    Dim i As Long
    ws.Columns(col).ColumnWidth = 8
    For i = 1 To 3
        If ws.Columns(col).Width > 0 Then
            ws.Columns(col).ColumnWidth = ws.Columns(col).ColumnWidth * pts / ws.Columns(col).Width
        End If
    Next
End Sub

'OMR定義ファイル: 1ページ目の実測ジオメトリ（左上マーカー中心が原点・pt）
Private Sub WriteOmrDef(ws As Worksheet, slotLabel() As String, _
        ByVal cM1 As Long, ByVal cT As Long, ByVal cD0 As Long, ByVal cM2 As Long)
    Dim p As String, fpath As String
    Dim oX As Double, oY As Double
    Dim mW As Double, mH As Double
    Dim txt As String
    Dim d As Long, k As Long, s As Long
    Dim stm As Object
    Dim x0 As Double, y0 As Double
    Dim pitchX As Double, pitchY As Double
    Dim cellW As Double, cellH As Double

    'マーカー中心（1ページ目: 上=行1, 下=行 4+kk+... → 行 kk+5）
    oX = ws.Cells(1, cM1).Left + ws.Cells(1, cM1).Width / 2
    oY = ws.Cells(1, cM1).Top + ws.Rows(1).Height / 2
    mW = (ws.Cells(1, cM2).Left + ws.Cells(1, cM2).Width / 2) - oX
    mH = (ws.Cells(5 + kk, cM1).Top + ws.Rows(5 + kk).Height / 2) - oY

    'グリッド先頭セル（行5 = top+4, 列 cD0）
    x0 = ws.Cells(5, cD0).Left + ws.Cells(5, cD0).Width / 2 - oX
    y0 = ws.Cells(5, cD0).Top + ws.Rows(5).Height / 2 - oY
    cellW = ws.Cells(5, cD0).Width
    cellH = ws.Rows(5).Height
    If dd > 1 Then
        pitchX = (ws.Cells(5, cD0 + 1).Left + ws.Cells(5, cD0 + 1).Width / 2 - oX) - x0
    Else
        pitchX = cellW
    End If
    If kk > 1 Then
        pitchY = (ws.Cells(6, cD0).Top + ws.Rows(6).Height / 2 - oY) - y0
    Else
        pitchY = cellH
    End If

    txt = "PTMOMR,2" & vbCrLf
    txt = txt & "size," & Format(mW, "0.00") & "," & Format(mH, "0.00") & vbCrLf
    txt = txt & "days," & dd & ",slots," & kk & vbCrLf
    txt = txt & "grid," & Format(x0, "0.00") & "," & Format(y0, "0.00") & "," & _
        Format(pitchX, "0.00") & "," & Format(pitchY, "0.00") & "," & _
        Format(cellW, "0.00") & "," & Format(cellH, "0.00") & vbCrLf

    'IDビット帯（B列 行2..9 の中心）
    txt = txt & "idstrip," & Format(ws.Cells(2, cM1).Left + ws.Cells(2, cM1).Width / 2 - oX, "0.00") & _
        "," & Format(ws.Cells(2, cM1).Width, "0.00")
    For k = 2 To 9
        txt = txt & "," & Format(ws.Cells(k, cM1).Top + ws.Rows(k).Height / 2 - oY, "0.00") & _
            "," & Format(ws.Rows(k).Height, "0.00")
    Next
    txt = txt & vbCrLf

    txt = txt & "labels"
    For d = 1 To dd
        For k = 1 To kk
            txt = txt & "," & slotLabel((d - 1) * kk + k)
        Next
    Next
    txt = txt & vbCrLf

    For s = 1 To ss
        txt = txt & "student," & CStr(Val(CStr(Sheets("名簿").Cells(1 + s, 1).Value))) & "," & _
            CStr(Sheets("名簿").Cells(1 + s, 2).Value) & " " & _
            CStr(Sheets("名簿").Cells(1 + s, 3).Value) & vbCrLf
    Next

    p = ThisWorkbook.Path
    If p = "" Or LCase(Left(p, 4)) = "http" Then p = Environ("USERPROFILE") & "\Desktop"
    fpath = p & Application.PathSeparator & "調査票定義.csv"

    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText txt
    stm.SaveToFile fpath, 2
    stm.Close
    Set stm = Nothing
End Sub

'--------------------------------------------------------------
' クイック入力（紙回収分のキーボード入力）
'--------------------------------------------------------------
Public Sub QuickInput()
    Dim ws As Worksheet
    Dim r0 As Long, c0 As Long
    Dim s As Long, j As Long, d As Long, k As Long
    Dim ans As String
    Dim tok As Variant
    Dim tt As String
    Dim guide As String
    Dim mark As String
    Dim n As Long, total As Long
    Dim slotMo(1 To MAXSLOT) As Long, slotDy(1 To MAXSLOT) As Long
    Dim slotHH(1 To MAXSLOT) As Long, slotMM(1 To MAXSLOT) As Long
    Dim slotLabel(1 To MAXSLOT) As String
    Dim p As Long

    EnsureDims
    If ss = 0 Or dd * kk = 0 Then
        MsgBox "名簿と初期設定を先に済ませてください。"
        Exit Sub
    End If
    If BuildSlots(slotMo, slotDy, slotHH, slotMM, slotLabel) = 0 Then
        MsgBox "希望入力シートに日付・時間が設定されていません。"
        Exit Sub
    End If

    Set ws = Sheets("希望入力")
    r0 = ws.Range("D2").Row
    c0 = ws.Range("D2").Column
    ws.Activate

    guide = "コマ番号表:" & vbLf
    For d = 1 To dd
        guide = guide & "  " & d & "日目 = " & slotMo((d - 1) * kk + 1) & "/" & _
            slotDy((d - 1) * kk + 1) & vbLf
    Next
    guide = guide & "  時刻: "
    For k = 1 To kk
        guide = guide & "(" & k & ")" & Format(slotHH(k), "0") & ":" & Format(slotMM(k), "00") & " "
        If k Mod 6 = 0 Then guide = guide & vbLf & "        "
    Next

    total = 0
    Do
        ans = InputBox("出席番号を入力してください（終了 = 空欄のまま OK）", "クイック入力")
        If ans = "" Then Exit Do
        s = 0
        For j = 1 To ss
            If CLng(Val(CStr(Sheets("名簿").Cells(1 + j, 1).Value))) = Val(ans) Then
                s = j
                Exit For
            End If
        Next
        If s = 0 Then
            MsgBox "出席番号 " & ans & " は名簿にありません。"
        Else
            On Error Resume Next
            ws.Cells(r0 + s, 1).Select
            On Error GoTo 0
            ans = InputBox(Val(ans) & "番 " & _
                CStr(Sheets("名簿").Cells(1 + s, 2).Value) & " さんの希望を入力:" & vbLf & _
                "  例: 1-3 1-5 2-10 （日-コマ番号をスペース区切り）" & vbLf & _
                "  ◎1-3 で優先、△2-4 で第2希望、c で行クリア" & vbLf & vbLf & guide, _
                "クイック入力")
            If ans <> "" Then
                ws.Unprotect
                If LCase(Trim(ans)) = "c" Then
                    For j = 1 To dd * kk
                        mark = CStr(ws.Cells(r0 + s, c0 + j - 1).Value)
                        If mark = "○" Or mark = "◎" Or mark = "△" Then
                            ws.Cells(r0 + s, c0 + j - 1).ClearContents
                        End If
                    Next
                Else
                    n = 0
                    ans = Replace(ans, "　", " ")
                    ans = Replace(ans, ",", " ")
                    For Each tok In Split(ans, " ")
                        tt = Trim(CStr(tok))
                        If tt <> "" Then
                            mark = "○"
                            If Left(tt, 1) = "◎" Then
                                mark = "◎"
                                tt = Mid(tt, 2)
                            ElseIf Left(tt, 1) = "△" Then
                                mark = "△"
                                tt = Mid(tt, 2)
                            End If
                            p = InStr(tt, "-")
                            If p > 0 Then
                                d = CLng(Val(Left(tt, p - 1)))
                                k = CLng(Val(Mid(tt, p + 1)))
                                If d >= 1 And d <= dd And k >= 1 And k <= kk Then
                                    j = (d - 1) * kk + k
                                    ws.Cells(r0 + s, c0 + j - 1) = mark
                                    n = n + 1
                                End If
                            End If
                        End If
                    Next
                    total = total + n
                End If
                ws.Protect AllowFormattingCells:=False
            End If
        End If
    Loop
    If total > 0 Then KomaCheck
End Sub
