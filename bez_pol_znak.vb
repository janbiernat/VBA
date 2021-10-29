Function BezPolZnak(Str As String) 
'BezPolZnak - Funkcja usuwa polskie znaki z podanego ciągu znaków. 
'Copyright (c)by Jan T. Biernat 
' 
    Dim I As Integer 
    Dim S As String 
    I = 0 
    S = "" 
    If (Trim(Str) <> "") Then 
        For I = 1 To Len(Str) 
            Select Case Mid(Str, I, 1) 
                Case "Ą"        'Ą, ą. 
                    S = S + "A" 
                Case "ą" 
                    S = S + "a" 
                Case "Ć"        'Ć, ć. 
                    S = S + "C" 
                Case "ć" 
                    S = S + "c" 
                Case "Ę"        'Ę, ę. 
                    S = S + "E" 
                Case "ę" 
                    S = S + "e" 
                Case "Ł"        'Ł, ł. 
                    S = S + "L" 
                Case "ł" 
                    S = S + "l" 
                Case "Ń"        'Ń, ń. 
                    S = S + "N" 
                Case "ń" 
                    S = S + "n" 
                Case "Ó"        'Ó, ó. 
                    S = S + "O" 
                Case "ó" 
                    S = S + "o" 
                Case "Ś"        'Ś, ś. 
                    S = S + "S" 
                Case "ś" 
                    S = S + "s" 
                Case "Ź", "Ż"   'Ź, ź, Ż, ż. 
                    S = S + "Z" 
                Case "ź", "ż" 
                    S = S + "z" 
                Case Else 
                    S = S + Mid(Str, I, 1) 
            End Select 
        Next I 
    End If 
    BezPolZnak = S 
End Function 
 
Sub bez_polskich_liter() 
'Wywołanie funkcji "BezPolZnak". 
' 
    Cells(1, 1).Value = BezPolZnak("Test: Ą, ą, Ć, ć, Ę, ę, Ł, ł, Ń, ń, Ó, ó, Ś, ś, Ź, ź, Ż, ż. Koniec testu.") 
End Sub 