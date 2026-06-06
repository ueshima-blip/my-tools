Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    If Left(Sh.Name, 1) <> "第" Then Exit Sub
    If Target.Value = "テスト" Then HandleTestEntry Sh, Target: Exit Sub
    Dim cC As Variant: cC = Array(4, 7, 10, 13, 16)
    Dim sC As Variant: sC = Array(2, 5, 8, 11, 14)
    Dim mC As Variant: mC = Array(2, 3, 4, 5, 6)
    Dim i As Integer
    For i = 0 To 4
        If Target.Column = cC(i) Then
            Dim pr As Integer
            For pr = 1 To 7
                If Target.Row = 5 + (pr - 1) * 2 Then
                    Dim ms As Worksheet
                    Set ms = ThisWorkbook.Sheets("時間割マスター")
                    Dim cv As String: cv = Trim(CStr(Target.Value))
                    Dim cl As String
                    If cv = "" Or cv = ChrW(215) Or UCase(cv) = "X" Then
                        cl = CStr(ms.Cells(pr + 2, mC(i)).Text)
                    Else
                        Dim dc As String: dc = Left(cv, 1)
                        Dim pn As Integer: pn = 0
                        On Error Resume Next: pn = CInt(Mid(cv, 2)): On Error GoTo 0
                        Dim mc2 As Integer: mc2 = 0
                        Select Case dc
                            Case "月": mc2 = 2: Case "火": mc2 = 3: Case "水": mc2 = 4
                            Case "木": mc2 = 5: Case "金": mc2 = 6
                        End Select
                        If mc2 > 0 And pn >= 1 And pn <= 7 Then
                            cl = CStr(ms.Cells(pn + 2, mc2).Text)
                        Else: cl = cv
                        End If
                    End If
                    Sh.Unprotect ""
                    Application.EnableEvents = False
                    Sh.Cells(Target.Row, sC(i)).NumberFormat = "@"
                    Sh.Cells(Target.Row, sC(i)).Value = cl
                    Application.EnableEvents = True
                    Sh.Protect ""
                    Exit Sub
                End If
            Next pr
            Exit Sub
        End If
    Next i
End Sub
