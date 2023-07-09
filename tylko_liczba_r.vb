Function TylkoLiczbaR(Str As String, Dl As Integer, IlePoKropce As Integer) As String
'TylkoLiczbaR - Funkcja wyciąga z podanego ciągu znaków tylko cyfry, znak minus i kropkę.
'Copyright (c)by Jan T. Biernat
'
    Dim T As String
    Dim Tylko As String
    Dim Start As Integer
    Dim Kropka As Boolean
    Dim LiczPoKropce As Integer
    Dim I As Integer
    'Wartości startowe
        T = ""
        Tylko = ""
        Start = 0
        Kropka = False
        LiczPoKropce = 0
    'Zabezpieczenie parametrów funkcji.
        If (Dl < 3) Then
            Dl = 3
        End If
        If (IlePoKropce < 2) Then
            IlePoKropce = 2
        End If
        T = Trim(Mid(Str, 1, Dl))
    'Wyodrębnij cyfry z podanego ciągu znaków.
        If (T <> "") Then
            If (Mid(T, 1, 1) = "-") Then
                Tylko = Tylko + Mid(T, 1, 1)
                Start = 2
            Else
                Start = 1
            End If
            I = 0
            For I = Start To Len(T)
                If ((Mid(T, I, 1) >= "0") And (Mid(T, I, 1) <= "9") And (LiczPoKropce < IlePoKropce)) Then
                    Tylko = Tylko + Mid(T, I, 1)
                    If (Kropka = True) Then
                        LiczPoKropce = LiczPoKropce + 1
                    End If
                Else
                    If ((Mid(T, I, 1) = ".") And (Kropka = False)) Then
                        Tylko = Tylko + Mid(T, I, 1)
                        Kropka = True
                    End If
                End If
            Next I
            TylkoLiczbaR = Tylko
        Else
            TylkoLiczbaR = ""
        End If
End Function

Sub WyodrebnijTylkoCyfry()
'Wyodrębnij tylko cyfry (Wywołanie funkcji "TylkoLiczbaR").
'
    Dim I As Integer
    I = 0
    Do
        I = I + 1
        Cells(I, 2).Value = TylkoLiczbaR(Cells(I, 1).Value, 8, 1)
    Loop Until (Cells(I, 1).Value = "")
End Sub