Sub ComboBoxPobierzDane(NrKolumny As Integer, NazwaKomponentu As ComboBox, NazwaArkusza As String) 
'ComboBoxPobierzDane - Pobiera dane z arkusza z wybranej kolumny bez powtórzeń. 
'Copyright (c)by Jan T. Biernat 
' 
'Wywołanie procedury: 
'Call ComboBoxPobierzDane(2, ComboBox1, "Arkusz1") 
' 
    Dim Licznik As Integer 
    Dim Spr As Integer 
    Dim CzyIstnieje As Boolean 
    Dim Komorka As String 
    With NazwaKomponentu 
        .Clear 'Wyczyść zawartość komponentu. 
        Licznik = 0 'Przypisanie wartości 0 do zmiennej liczbowej "Licznik". 
        Do 
            Komorka = "" 'Wyczyszczenie zmiennej tekstowej "Komorka". 
            Komorka = Trim(Sheets(NazwaArkusza).Cells(2 + Licznik, NrKolumny).Value) 'Pobranie danych z wybranego arkusza, kolumny i komórki. 
            If (Komorka <> "") Then 
                'Jeżeli komórka zawiera jakąś wartość, to wykonaj poniższe instrukcje. 
                'Sprawdzenie, czy na liście znajduje się pobrana z arkusza informacja. 
                CzyIstnieje = False 
                For Spr = 0 To .ListCount - 1 
                    If (LCase(Komorka) = LCase(.List(Spr))) Then 
                        CzyIstnieje = True 
                        Exit For 
                    End If 
                Next Spr 
                'Dodaj dane do listy. 
                If (CzyIstnieje = False) Then 
                    .AddItem (Komorka) 
                End If 
            End If 
            Licznik = Licznik + 1 'Zwiększenie zawartości zmiennej "Licznik" o wartość 1. 
        Loop Until Komorka = "" 'Jeżeli natrafi na pustą komórkę, to pętla zakończy swoje działanie. 
    End With 
End Sub 