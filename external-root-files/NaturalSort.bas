Option Explicit

' Natural sort for sector/group names: E1, E2, ... E9, E10 (not E1, E10, E2)
' Use in all 4 report modules for sector/group ordering

Public Function NaturalCompare(ByVal a As String, ByVal b As String) As Integer
    Dim ai As Long, bi As Long, aLen As Long, bLen As Long
    Dim aNum As String, bNum As String
    ai = 1: bi = 1
    aLen = Len(a): bLen = Len(b)
    Do While ai <= aLen Or bi <= bLen
        aNum = ""
        bNum = ""
        Do While ai <= aLen
            If Mid(a, ai, 1) >= "0" And Mid(a, ai, 1) <= "9" Then
                aNum = aNum & Mid(a, ai, 1)
                ai = ai + 1
            Else
                Exit Do
            End If
        Loop
        Do While bi <= bLen
            If Mid(b, bi, 1) >= "0" And Mid(b, bi, 1) <= "9" Then
                bNum = bNum & Mid(b, bi, 1)
                bi = bi + 1
            Else
                Exit Do
            End If
        Loop
        If aNum <> "" And bNum <> "" Then
            Dim d As Long
            d = CLng(aNum) - CLng(bNum)
            If d < 0 Then NaturalCompare = -1: Exit Function
            If d > 0 Then NaturalCompare = 1: Exit Function
        ElseIf aNum <> "" Then
            NaturalCompare = 1: Exit Function
        ElseIf bNum <> "" Then
            NaturalCompare = -1: Exit Function
        Else
            If ai <= aLen And bi <= bLen Then
                Dim ac As String, bc As String
                ac = Mid(a, ai, 1): bc = Mid(b, bi, 1)
                If ac < bc Then NaturalCompare = -1: Exit Function
                If ac > bc Then NaturalCompare = 1: Exit Function
                ai = ai + 1: bi = bi + 1
            ElseIf ai <= aLen Then
                NaturalCompare = 1: Exit Function
            ElseIf bi <= bLen Then
                NaturalCompare = -1: Exit Function
            Else
                NaturalCompare = 0: Exit Function
            End If
        End If
    Loop
    NaturalCompare = 0
End Function

' Sort array of keys (Variant) in place - natural order
Public Sub SortKeysNatural(keysArr As Variant)
    Dim i As Long, j As Long, lb As Long, ub As Long
    Dim tmp As Variant
    lb = LBound(keysArr): ub = UBound(keysArr)
    For i = lb To ub - 1
        For j = i + 1 To ub
            If NaturalCompare(CStr(keysArr(i)), CStr(keysArr(j))) > 0 Then
                tmp = keysArr(i)
                keysArr(i) = keysArr(j)
                keysArr(j) = tmp
            End If
        Next j
    Next i
End Sub
