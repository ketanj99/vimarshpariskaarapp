Option Explicit

Private Sub cmdDelete_Click()
    Tatvagnan_DeleteSelected Me.ListView1
End Sub

Private Sub UserForm_Initialize()
    ' Default values
    Me.cboAction.Clear
    Me.cboAction.AddItem "ADD"
    Me.cboAction.AddItem "REMOVE"
    Me.cboAction.ListIndex = 0
    Tatvagnan_LoadHistory Me.ListView1
End Sub

Private Sub cmdImport_Click()
    Dim act As String
    Dim pno As String

    act = Me.cboAction.text
    pno = Me.txtPushpNo.text

    If Trim$(act) = "" Or Trim$(pno) = "" Then
        MsgBox "Action (ADD/REMOVE) aur Pushp No dono zaroori hain.", vbExclamation
        Exit Sub
    End If

    Tatvagnan_ImportFromUI act, pno   ' files dialog yahin se khulega
    Tatvagnan_LoadHistory Me.ListView1
End Sub

Private Sub cmdRefresh_Click()
    Tatvagnan_LoadHistory Me.ListView1
End Sub

Private Sub cmdSummaryPdf_Click()
    Dim pno As String
    pno = Trim$(Me.txtPushpNo.text)   ' yahi textbox me Pushp No. likhte ho

    If pno = "" Then
        MsgBox "Pushp No. required for PDF.", vbExclamation
        Exit Sub
    End If

    Tatvagnan_GenerateSummaryPdf pno   ' TatvagnanReportModule ka function
End Sub

Private Sub cmdGroupWisePdf_Click()
    Dim pno As String
    pno = Trim$(Me.txtPushpNo.text)
    
    If pno = "" Then
        MsgBox "Pushp No. required for Group-Wise PDF.", vbExclamation
        Exit Sub
    End If
    
    ' Group-wise PDF generate (4 groups per page)
    Tatvagnan_GenerateGroupWisePdf pno
End Sub
