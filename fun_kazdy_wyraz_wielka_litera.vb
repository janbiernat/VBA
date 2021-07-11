Function WyrazWielkaLitera(Str As String) 
'WyrazWielkaLitera (c)by Jan T. Biernat 
' 
    Dim I As Integer 
    Dim Wynik As String 
    I = 0 
    Wynik = "" 
    Str = LCase(Trim(Str)) 
    If (Str <> "") Then 
        For I = 2 To Len(Str) 
            If (((Mid(Str, I - 1, 1) = " ") And (Mid(Str, I, 1) <> " ")) Or ((Mid(Str, I - 1, 1) = "-") And (Mid(Str, I, 1) <> " "))) Then 
                Wynik = Wynik + UCase(Mid(Str, I, 1)) 
            Else 
                Wynik = Wynik + Mid(Str, I, 1) 
            End If 
        Next I 
        WyrazWielkaLitera = UCase(Mid(Str, 1, 1)) + Wynik 
    Else 
        WyrazWielkaLitera = "BŁĄD -?" 
    End If 
End Function 
 
Sub kazdy_wyraz_wielka_litera() 
'Każdy wyraz wielką literą. 
'Copyright (c)by Jan T. Biernat 
' 
    Dim WierszNr As Integer 
    WierszNr = 0 
    Do 
        WierszNr = WierszNr + 1 
        Cells(WierszNr, 2).Value = WyrazWielkaLitera(Cells(WierszNr, 1).Value) 
    Loop Until (Cells(WierszNr, 1).Value = "") 
End Sub 
