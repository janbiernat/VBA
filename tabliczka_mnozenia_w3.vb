'ThisWorkbook
Private Sub Workbook_Open()
'Zainicjowanie programu.
'
   ZG_iDobra = 0
   ZG_iZla = 0
   Cells(1, 2).Value = "Dobrych odpowiedzi: " + Str(ZG_iDobra)  'CStr(P1) - Zamienia cyfrę na liczbę traktowaną jako tekst.
                                                                '           W miejsce parametru (tj. P1) należy umieścić zmienną liczbową.
   Cells(1, 5).Value = "Złych odpowiedzi: " + Str(ZG_iZla)
   Cells(2, 2).Value = "" 'Wyczyść komórkę (tj. B2) w której wyświetlana jest ocena.
   Cells(6, 5).Value = "" 'Wyczyść komórkę (tj. E6) w której wyświetlana jest podpowiedź.
   Arkusz1.LosujPytanie
   Arkusz1.Range("A1").Select
End Sub
'==
'Arkusz1
'
Dim ZG_iDobra, ZG_iZla As Integer 'Zadeklarowanie zmiennych globalnych.
Dim ZG_iLiczbaA, ZG_iLiczbaB As Integer

Sub LosujPytanie()
'Losowanie pytania.
'
    Cells(6, 5).Value = "" 'Wyczyść komórkę (tj. E6) w której wyświetlana jest podpowiedź.
    Randomize 'Zainicjowanie generatora liczb losowych.
    'Losuje liczbę z przedziału od 1 do 10 i przypisuje ją do odpowiedniej komórki (tj. B5).
    ZG_iLiczbaA = 0
    ZG_iLiczbaA = Int((10 * Rnd) + 0) 'Generuje losową liczbę z zakresu od 0 do 10 i przypisuje zmiennej "ZG_iLiczbaA".
    ZG_iLiczbaB = 0
    ZG_iLiczbaB = Int((10 * Rnd) + 0)
    Cells(3, 1).Value = CStr(ZG_iLiczbaA) + " x " + CStr(ZG_iLiczbaB)   'CStr(P1) - Zamienia cyfrę na liczbę traktowaną jako tekst.
                                                                        '           W miejsce parametru (tj. P1) należy umieścić zmienną liczbową.
End Sub

Sub Komentarz(logDobry As Boolean)
'Wyświetlenie komentarza do podanej odpowiedzi.
'
    If (logDobry = True) Then
       Cells(5, 5).Value = "Bardzo dobrze!" 'Wpisanie komentarza do odpowiedniej komórki (tj. F12).
    Else
       Cells(5, 5).Value = "Bardzo źle!"
    End If
    Cells(4, 4).Value = "" 'Wyczyszczenie komórki (tj. D4) do której wpisuje się odpowiedź.
    Application.Wait (Now + TimeValue("0:00:01")) 'Wstrzymuje działanie makra na 2 sekundy.
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'Sprawdzenie odpowiedzi.
'
    Dim tOcena As String 'Zadeklarowanie zmiennej tekstowej.
    Dim iProcent As Integer 'Zadeklarowanie zmiennej liczbowej całkowitej.
    If (Cells(4, 4).Value <> "") Then
        'Sprawdzenie odpowiedzi.
        If (ZG_iLiczbaA * ZG_iLiczbaB = Cells(4, 4).Value) Then
            'Wyświetlenie informacji o poprawnej odpowiedzi.
            Komentarz (True)
            LosujPytanie
            Cells(1, 2).Value = ""
            ZG_iDobra = ZG_iDobra + 1 'Zwiększ zawartość zmiennej liczbowej "ZG_iDobra" o wartość 1.
        Else
            'Wyświetlenie informacji o złej odpowiedzi.
            ZG_iZla = ZG_iZla + 1
            Cells(6, 5).Value = CStr(ZG_iLiczbaA) + " x " + CStr(ZG_iLiczbaB) + " = " + CStr(ZG_iLiczbaA * ZG_iLiczbaB) 'Wyświetl podpowiedź.
            Komentarz (False)
        End If
        iProcent = 0
        iProcent = Round((ZG_iDobra * 100) / (ZG_iDobra + ZG_iZla))
        Cells(1, 2).Value = "Dobrych odpowiedzi: " + Str(ZG_iDobra) + " (" + Str(iProcent) + "%)"
        Cells(1, 5).Value = "Złych odpowiedzi: " + Str(ZG_iZla) + " (" + Str(100 - iProcent) + "%)"
        'Wystaw ocenę.
        tOcena = ""
        If ((ZG_iDobra > 0) Or (ZG_iZla > 0)) Then
            If (iProcent > 95) Then
                tOcena = "Bardzo dobry (5.0)"
            Else
                If (iProcent > 80) Then
                    tOcena = "Dobry (4.0)"
                Else
                    If (iProcent > 70) Then
                        tOcena = "Dostateczny (3.0)"
                    Else
                        If (iProcent > 50) Then
                            tOcena = "Mierny (2.0)"
                        Else
                            tOcena = "Jedynka (1.0)"
                        End If
                    End If
                End If
            End If
            Cells(2, 2).Value = "Ocena: " + tOcena
        Else
            Cells(2, 2).Value = "" 'Wyczyść komórkę (tj. B2) w której jest wyświetlana ocena.
        End If
    Else
        Cells(5, 5).Value = "" 'Wyczyszczenie komentarza.
    End If
    Cells(5, 5).Value = "" 'Wyczyszczenie komentarza.
    Arkusz1.Range("D4").Select 'Uaktywnienie komórki E10.
End Sub
