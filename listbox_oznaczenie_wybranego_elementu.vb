Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean) 
'Oznaczenie wybranego elementu na liście. 
'Copyright (c)by Jan T. Biernat 
' 
    Dim T As String                                        '1. 
    With ListBox1                                          '2. 
        T = ""                                             '3. 
        T = .List(.ListIndex, 0)                           '4. 
        If (Left(T, 1) = ">") Then                         '5. 
            .List(.ListIndex, 0) = " " + Mid(T, 2, Len(T)) '6. 
        Else                                               '5. 
            .List(.ListIndex, 0) = ">" + Mid(T, 2, Len(T)) '7. 
        End If                                             '5. 
    End With                                               '2. 
    ' 
    'Legenda: 
    '1. Zadeklarowanie zmiennej tekstowej o nazwie "T". 
    '2. Konstrukcja wiążąca WITH ... END WITH. 
    '   Po słowie WITH umieszcza się np. nazwę komponentu (np. ListBox1) i od tego momentu 
    '   nie trzeba poprzedzać żadnej właściwości danego komponentu jego nazwą. 
    '   Gdyby w tym przykładzie nie było instrukcji wiążącej WITH ... END WITH, to 
    '   przed każdą właściwością ListIndex trzeba byłoby umieszczać konstrukcję 
    '   ListBox1.ListIndex. 
    '3. Wyczyszczenie zmiennej tekstowej "T". 
    '4. Pobranie z listy ListBox1 zaznaczonego elementu (tj. tekstu) i przypisanie go do 
    '   zmiennej tekstowej "T". Nr pobranego tekstu z listy określa właściwość ListIndex. 
    '5. Sprawdzenie, czy 1 znakiem podanego ciągu znaków jest znak ">". 
    '   Jeżeli 1 znakiem jest znak ">", to wykonaj instrukcje po słowie THEN. 
    '   W innym przypadku wykonaj instrukcje po słowie ELSE. 
    '   Instrukcja LEFT(P1, P2) służy do pobrania fragmentu tekstu z podanego ciągu znaków. 
    '                           W parametrze P1 umieszcza się zmienną tekstową. 
    '                           Natomiast w parametrze P2 umieszcza się liczbę całkowitą, 
    '                           określającą ile znaków licząc od lewej strony ma być wyciągnięta 
    '                           z podanego ciągu znaków. 
    '6. Przypisanie do wybranego elementu (tj. tekstu) z listy znaku spacji wraz z wcześniej 
    '   umieszczonym tekstem. 
    '   Instrukcja MID(P1, P2, P3) służy do wyciągnięcia dowolnego fragmentu tekstu 
    '                              z podanego ciągu znaków. 
    '                              P1 - w tym parametrze umieszcza się zmienną tekstową. 
    '                              P2 - w tym parametrze umieszcza się liczbę całkowitą, 
    '                                   określającą początek pobierania fragmentu tekstu 
    '                                   z podanego ciągu znaków. 
    '                              P3 - w tym parametrze umieszcza się liczbę całkowitą, 
    '                                   określającą ile znaków ma być pobranych. 
    '7. Przypisanie do wybranego elementu (tj. tekstu) z listy znaku ">" wraz z wcześniej 
    '   umieszczonym tekstem. 
End Sub 