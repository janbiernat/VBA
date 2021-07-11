Function Inicjaly(Str As String) 
'Inicjały (c)by Jan T. Biernat 
' 
    Dim I As Integer 
    Dim Wynik As String 
    I = 0 
    Wynik = "" 
    Str = LCase(Trim(Str)) 
    If (Str <> "") Then 
        For I = 2 To Len(Str) 
            If (Mid(Str, I - 1, 1) = " ") And (Mid(Str, I, 1) <> " ") Then 
                If (Mid(Str, I + 1, 1) = " ") Then 
                    Wynik = Wynik + Mid(Str, I, 1) 
                Else 
                    Wynik = Wynik + UCase(Mid(Str, I, 1)) 
                End If 
            End If 
        Next I 
        Inicjaly = UCase(Mid(Str, 1, 1)) + Wynik 
    Else 
        Inicjaly = "" 
    End If 
End Function 
 
Sub pobierz_inicjaly() 
    Dim WierszNr As Integer 
    WierszNr = 0 
    Do 
        WierszNr = WierszNr + 1 
        Cells(WierszNr, 2).Value = Inicjaly(Cells(WierszNr, 1).Value) 
    Loop Until (Cells(WierszNr, 1).Value = "") 
End Sub 
