'Form1
'=
Dim PytaniaIlosc, PytanieNr As Integer 'Zadeklarowanie zmiennych globalnych.

Private Sub CommandButton1_Click()
'Wyjdź z programu.
'
    Form1.Hide 'Ukrycie formatki.
End Sub

Private Sub CommandButton2_Click()
'Wykonaj poniższe instrukcje w momencie, gdy użytkownik kliknie lewym klawiszem myszy na ten komponent.
'
    'Sprawdź poprawność odpowiedzi.
    If (TextBox1.Text = Sheets("Daty historyczne").Cells(PytanieNr, 2)) Then
        'Wylosowanie 1 pytania z listy pytań.
        PytanieNr = 0
        PytanieNr = Int(((PytaniaIlosc - 2) * Rnd) + 2)
        Label5.Caption = Sheets("Daty historyczne").Cells(PytanieNr, 3)
        Label7.Caption = "Podpowiedź: " + CStr(Sheets("Daty historyczne").Cells(PytanieNr, 2)) + "." 'Wyświetla podpowiedź (ang. hint).
        TextBox1.Text = "" 'Wyczyszczenie komponentu "TextBox1".
    Else
        MsgBox ("BŁĄD -?Odpowiedź jest zła!" & Chr(13) & Chr(13) & "Prawidłowa odpowiedź to " & Sheets("Daty historyczne").Cells(PytanieNr, 2) & ".")
    End If
    TextBox1.SetFocus
End Sub

Private Sub Label4_Click()
'Wykonaj poniższe instrukcje, gdy użytkownik kliknie lewym klawiszem myszy na tym komponencie.
'
    If (Label7.Visible = False) Then
        Label7.Visible = True 'Włącz widoczność komponentu na formatce.
    Else
        Label7.Visible = False 'Wyłącz widoczność komponentu na formatce.
    End If
End Sub

Private Sub TextBox1_Change()
'Wykonaj poniższe instrukcje, gdy użytkownik zacznie pisać w polu edycyjnym "TextBox1".
'
    'Włącz lub wyłącz aktywność przycisku sprawdź.
    If (Trim(TextBox1.Text) <> "") Then
        CommandButton2.Enabled = True
    Else
        CommandButton2.Enabled = False
    End If
End Sub

Private Sub UserForm_Initialize()
'Poniższe instrukcje są wykonywane w momencie, gdy formatka jest wyświetlana na ekranie.
'
    Dim Licznik As Integer 'Zadeklarowanie zmiennych
    Dim Komorka As String
    Caption = "Nauka z dat historycznych"
    Label2.Caption = "Nauka ze znajomości dat."
    Label3.Caption = "Należy jako odpowiedź na zadane pytanie wpisać prawidłową datę."
    Label1.Width = Form1.Width 'Szerokość komponentu "Label1" jest taka sama jak formatki.
    Label7.Visible = False 'Wyłączenie podpowiedzi.
    PytaniaIlosc = 0 'Zadeklarowanie zmiennej liczbowej całkowitej "PytaniaIlosc" i przypisanie jej wartości zerowej.
    Licznik = 0
    Do
        Licznik = Licznik + 1
        Komorka = ""
        Komorka = Trim(Sheets("Daty historyczne").Cells(Licznik + 1, 2).Value)
        If (Komorka <> "") Then
            PytaniaIlosc = PytaniaIlosc + 1 'Zwiększ zawartość zmiennej liczbowej "PytaniaIlosc" o wartość 1.
        End If
    Loop Until Komorka = "" 'Blok instrukcji wewnątrz pętli DO ... LOOP UNTIL wykonuje się tak długo, aż natrafi na pustą komórkę.
    'Wylosowanie 1 pytania z listy pytań.
    PytanieNr = 0
    PytanieNr = Int(((PytaniaIlosc - 2) * Rnd) + 2)
    Label5.Caption = Sheets("Daty historyczne").Cells(PytanieNr, 3)
    Label7.Caption = "Podpowiedź: " + CStr(Sheets("Daty historyczne").Cells(PytanieNr, 2)) + "." 'Wyświetla podpowiedź (ang. hint).
    TextBox1.Text = "" 'Wyczyszczenie komponentu "TextBox1".
End Sub
