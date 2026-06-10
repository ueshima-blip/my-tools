Attribute VB_Name = "Module2"
Dim permPos As Integer
Dim permOrd() As Integer
Public pOrd() As Integer

Sub Permutation()

'コマ順の順列を求める
    ReDim pOrd(1 To WorksheetFunction.Permut(dd, dd), 1 To dd)
    Dim i As Long
    ReDim permOrd(1 To dd)
    permPos = 0
    For i = 1 To dd
        permOrd(i) = i
    Next i
    Perm 1

End Sub
Private Sub Perm(i As Long)
  
    Dim j As Long, t As Long
    If i < dd Then
        For j = i To dd
            t = permOrd(i): permOrd(i) = permOrd(j): permOrd(j) = t
            Perm i + 1
            t = permOrd(i): permOrd(i) = permOrd(j): permOrd(j) = t
        Next j
    Else
        For j = 1 To dd
            pOrd(permPos + 1, j) = permOrd(j)
        Next j
        permPos = permPos + 1
    End If

End Sub
Function CellColor(ByVal rng As Range)
 
'確定コマの判定用
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
Function MyPrev(c As Variant)

'休憩時間帯の判定用
    If ActiveSheet.Cells(Range("FIRST").Row, c).Interior.ColorIndex = kColor Then
        MyPrev = kColor
    Else
        MyPrev = xlNone
    End If
    
End Function
