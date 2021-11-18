Private Sub UserForm_Initialize() 
'UserForm_Initialize - Ustawienia startowe. 
' 
    'ListBox - Pobranie danych z arkusza. 
    Dim Licznik As Integer 
    Dim Spr As Integer 
    Dim KolNr As Integer 
    Dim CzyIstnieje As Boolean 
    Dim ArkuszNazwa As String 
    Dim Komorka As String 
    KolNr = 0 
    KolNr = 4 
    ArkuszNazwa = "" 
    ArkuszNazwa = "Opcje" 
    With ListBox1 
        .Clear 'Wyczyść zawartość komponentu. 
        Licznik = 0 'Przypisanie wartości 0 do zmiennej liczbowej "Licznik". 
        Do 
            Komorka = "" 'Wyczyszczenie zmiennej tekstowej "Komorka". 
            Komorka = Trim(Sheets(ArkuszNazwa).Cells(2 + Licznik, KolNr).Value) 'Pobranie danych z wybranego arkusza, kolumny i komórki. 
            If (Komorka <> "") Then 
                'Jeżeli komórka zawiera jakąś wartość, to wykonaj poniższe instrukcje. 
                'Sprawdzenie, czy na liście znajduje się pobrana z arkusza informacja. 
                CzyIstnieje = False 
                For Spr = 0 To .ListCount - 1 
                    If (LCase(Komorka) = Trim(LCase(.List(Spr, 0)))) Then 
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