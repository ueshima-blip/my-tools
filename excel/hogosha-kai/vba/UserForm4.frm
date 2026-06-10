Attribute VB_Name = "UserForm4"
Attribute VB_Base = "0{557BAD17-9B0D-4D80-9177-0160AED537DB}{BA4A2786-E619-42BA-A47A-9682990EBD1C}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Dim dp As Integer

Private Sub CommandButton1_Click()
    
    If OptionButton6 Then
        Dim t As String
        Dim tt(5) As Integer
        t = TextBox1
        If Len(t) = 0 Then Exit Sub
        For i = 1 To dd
            tt(Val(Mid(t, i, 1))) = tt(Val(Mid(t, i, 1))) + 1
        Next
        For i = 1 To dd
            If tt(i) <> 1 Then Exit Sub
        Next
        m1 = "Specd" & t
    End If
    If OptionButton1 Then m1 = "Least"
    If OptionButton2 Then m1 = "Permu"
    If OptionButton3 Then m2 = "Least"
    If OptionButton4 Then m2 = "Randm"
    If OptionButton5 Then m1 = "First"
    ChoseiSagyo = m1 & m2
    
    TryLimit = ComboBox1
    If CheckBox1 = True Then scrUpdate = True Else scrUpdate = False

    Me.Hide
    

End Sub

Private Sub CommandButton2_Click()

    Unload Me

End Sub

Private Sub OptionButton1_Click()
    ComboItemChange
    TextBox1.Visible = False
    Label3.Visible = False
End Sub

Private Sub OptionButton2_Click()
    ComboItemChange
    TextBox1.Visible = False
    Label3.Visible = False
End Sub

Private Sub OptionButton3_Click()
    ComboItemChange
End Sub

Private Sub OptionButton4_Click()
    ComboItemChange
End Sub

Private Sub OptionButton5_Click()
    ComboItemChange
    
    TextBox1.Visible = False
    Label3.Visible = False

End Sub

Private Sub OptionButton6_Click()
    ComboItemChange
    
    TextBox1.Visible = True
    Label3.Visible = True

End Sub

Private Sub UserForm_Initialize()

    ComboBox1.AddItem 1
    ComboBox1 = 1
    dp = WorksheetFunction.Permut(dd, dd)
    ChoseiSagyo = ""
    TextBox1.Visible = False
    Label3.Visible = False

End Sub
Private Sub ComboItemChange()

    With ComboBox1
        .Clear
        Select Case OptionButton2
        Case False
            If OptionButton3 Then
                .AddItem 1
                .Value = 1
            Else
                .AddItem 10
                .AddItem 20
                .AddItem 50
                .Value = 10
            End If
        Case True
            If OptionButton3 Then
                .AddItem dp
                .Value = dp
            Else
                .AddItem dp
                .AddItem dp * 2
                .AddItem dp * 5
                .Value = dp
            End If
        End Select
        
    End With


End Sub
