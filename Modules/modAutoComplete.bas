Attribute VB_Name = "modAutoComplete"
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''' ћќƒ”Ћ№ ј¬“ќ«јѕќЋЌ≈Ќ»я '''''''''''''''''''''''
        ''''''''''''''''''''''''' ƒЋя ЁЋ≈ћ≈Ќ“ќ¬ COMBOBOX Ќј √Ћј¬Ќќ…'''''''''''''''
        ''''''''''''''''''''''''''''''''''''' ‘ќ–ћ≈ ''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit 'повышаем "придирчивость" компил€тора - увеличиваем надежность кода

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
                ByVal hwnd As Long, _
                ByVal wMsg As Long, _
                ByVal wParam As Long, _
                lParam As Any) As Long
Private Const CB_ERR = -1, CB_SELECTSTRING = &H14D, CB_SHOWDROPDOWN = &H14F, CBN_SELENDOK = 9

Public Function AutoMatchCBBox(ByRef cbBox As ComboBox, ByVal KeyAscii As Integer) As Integer
    
        
    Dim strFindThis As String, bContinueSearch As Boolean
    Dim lResult As Long, lStart As Long, lLength As Long
    AutoMatchCBBox = 0
    bContinueSearch = True
    lStart = cbBox.SelStart
    lLength = cbBox.SelLength

    On Error GoTo ErrHandle
        
    If KeyAscii < 32 Then
        bContinueSearch = False
        cbBox.SelLength = 0
        If KeyAscii = Asc(vbBack) Then
            If lLength = 0 Then
                If Len(cbBox) > 0 Then
                    cbBox.Text = Left(cbBox.Text, Len(cbBox) - 1)
                End If
            Else
                cbBox.Text = Left(cbBox.Text, lStart)
            End If
            cbBox.SelStart = Len(cbBox)
        ElseIf KeyAscii = vbKeyReturn Then
            cbBox.SelStart = Len(cbBox)
            lResult = SendMessage(cbBox.hwnd, CBN_SELENDOK, 0, 0)
            AutoMatchCBBox = KeyAscii
        End If
    Else
        If lLength = 0 Then
            strFindThis = cbBox.Text & Chr(KeyAscii)
        Else
            strFindThis = Left(cbBox.Text, lStart) & Chr(KeyAscii)
        End If
    End If
    
    If bContinueSearch Then
        Call VBComBoBoxDroppedDown(cbBox)
        lResult = SendMessage(cbBox.hwnd, CB_SELECTSTRING, -1, ByVal strFindThis)
        If lResult = CB_ERR Then
            cbBox.Text = strFindThis
            cbBox.SelLength = 0
            cbBox.SelStart = Len(cbBox)
        Else
            cbBox.SelStart = Len(strFindThis)
            cbBox.SelLength = Len(cbBox) - cbBox.SelStart
        End If
    End If
    On Error GoTo 0
    Exit Function
    
ErrHandle:
    WriteToLog "Ётого метода здесь быть не должно, но € его все же исполнил с ошибкой " & Err.description
    Debug.Assert False
    AutoMatchCBBox = KeyAscii
    On Error GoTo 0
End Function

Private Sub VBComBoBoxDroppedDown(ByRef cbBox As ComboBox)
    Call SendMessage(cbBox.hwnd, CB_SHOWDROPDOWN, Abs(True), 0)
End Sub
