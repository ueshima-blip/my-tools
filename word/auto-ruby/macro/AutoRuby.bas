Attribute VB_Name = "AutoRuby"
'==================================================================
'  自動ルビ振り（ふりがな）マクロ  AutoRuby
'  - Word内蔵の「ふりがなガイド」を使って漢字にルビを自動で振ります。
'  - このファイルは「VBA画面にドラッグ＆ドロップ」または
'    「ファイル → ファイルのインポート」で丸ごと取り込めます（コピペ不要）。
'  - 取り込むと、表示 → マクロ に下の4つが並びます。
'==================================================================
Option Explicit

'ルビを振った漢字を格納するArray
Public kanjiArray(9999) As String
'KanjiArrayのインデックス
Public KI As Long

'選択した範囲内の文字列にルビ設定
Public Sub MakeRubiPartial()
  SetPhoneticRange Selection.Range, False
End Sub

'文書全体にルビ設定
Public Sub MakeRubiAll()
  SetPhoneticRange ActiveDocument.Range, False
End Sub

'選択した範囲内の文字列にルビ設定（最初の漢字のみ）
Public Sub MakeFirstRubiPartial()
  SetPhoneticRange Selection.Range, True
End Sub

'文書全体にルビ設定（最初の漢字のみ）
Public Sub MakeFirstRubiAll()
  SetPhoneticRange ActiveDocument.Range, True
End Sub

Private Sub SetPhoneticRange(ByVal rng As Word.Range, ByVal FirstFlag As Boolean)
  Dim r As Word.Range
  Dim s As Word.Range
  Dim i As Long
  Dim dFlag As Boolean

  ' kanjiArrayのインデックスの初期化
  KI = 0

  '単語単位で処理
  For Each r In rng.Words
    'ルビが振られていないか最初にフィールド数で判定
    If r.Fields.Count < 1 Then
      ' 漢字が含まれているか判定
      If ChkKanjiRange2(r) = True Then

        ' 全部が漢字か判定
        If ChkKanjiRange(r) = True Then

          If FirstFlag = False Then
            ' 全ての漢字にルビをふる
            r.Select
            Application.Dialogs(wdDialogPhoneticGuide).Show 1
          Else
            ' 最初に出てきた漢字にだけルビをふる
            If inKanjiArray(r.Text) = False Then
              addKanjiArray (r.Text)
              r.Select
              Application.Dialogs(wdDialogPhoneticGuide).Show 1
            End If
          End If

        Else
          '文字単位で処理
          i = 1
          For Each s In r.Characters
            ' 漢字か判定
            If ChkKanjiRange(s) = True Then
              ' 次の文字が漢字か判定
              dFlag = False
              If i < Len(r.Text) And Len(Mid(r.Text, i + 1, 1)) > 0 Then
                If isKanji(Mid(r.Text, i + 1, 1)) = True Then
                  ' 漢字が２文字続きの場合、一緒にルビを振る
                  s.End = s.End + 1
                  dFlag = True
                End If
              End If

              If FirstFlag = False Then
                ' 全ての漢字にルビをふる
                s.Select
                Application.Dialogs(wdDialogPhoneticGuide).Show 1
              Else
                ' 最初に出てきた漢字にだけルビをふる
                If inKanjiArray(s.Text) = False Then
                  If dFlag = True Then
                    addKanjiArray (Mid(r.Text, i, 1))
                    addKanjiArray (Mid(r.Text, i + 1, 1))
                  End If
                  addKanjiArray (s.Text)
                  s.Select
                  Application.Dialogs(wdDialogPhoneticGuide).Show 1
                End If
              End If

            End If
            i = i + 1
          Next
        End If

      End If
    End If
  Next
End Sub

Private Function ChkKanjiRange(ByVal rng As Word.Range) As Boolean
'指定したRangeが全部漢字だったらTrue
  Dim ret As Boolean
  Dim i As Long

  ret = True
  For i = 1 To Len(rng.Text)
    If isKanji(Mid(rng.Text, i, 1)) = False Then
      ret = False
      Exit For
    End If
  Next
  ChkKanjiRange = ret
End Function

Private Function ChkKanjiRange2(ByVal rng As Word.Range) As Boolean
'指定したRangeに漢字が１文字でも含まれていたらTrue
  Dim ret As Boolean
  Dim i As Long

  ret = False
  For i = 1 To Len(rng.Text)
    If isKanji(Mid(rng.Text, i, 1)) = True Then
      ret = True
      Exit For
    End If
  Next
  ChkKanjiRange2 = ret
End Function

Private Function isKanji(ByVal strIn As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "[一-龠〃々〆〇]"

    If re.test(strIn) Then
        'MsgBox "入力文字列には、漢字が含まれてます。"
        isKanji = True
    Else
        'MsgBox "入力文字列には、漢字が含まれていません。"
        isKanji = False
    End If
End Function

Private Function inKanjiArray(ByVal str As String) As Boolean
  Dim ret As Boolean
  Dim i As Long
  ret = False

  For i = 0 To KI + 1
    If StrComp(kanjiArray(i), str) = 0 Then
      ret = True
      Exit For
    End If
  Next
  inKanjiArray = ret
End Function

Private Function addKanjiArray(ByVal str As String) As Boolean
  kanjiArray(KI) = str
  KI = KI + 1
End Function

'==================================================================
'  ↓ここから下は「リボンのボタン」から呼び出すための受け口です。
'   表示→マクロから実行する場合は使いません（あっても害はありません）。
'   ribbon/customUI14.xml のボタンが、この4つを呼び出します。
'==================================================================
Public Sub Ribbon_MakeRubiAll(ByVal control As IRibbonControl)
  MakeRubiAll
End Sub

Public Sub Ribbon_MakeRubiPartial(ByVal control As IRibbonControl)
  MakeRubiPartial
End Sub

Public Sub Ribbon_MakeFirstRubiAll(ByVal control As IRibbonControl)
  MakeFirstRubiAll
End Sub

Public Sub Ribbon_MakeFirstRubiPartial(ByVal control As IRibbonControl)
  MakeFirstRubiPartial
End Sub
