Function DataRMD2_TerazJest() As String
'DataRMD2_TerazJest - Funkcja zwraca datę systemową (bieżącą) w postaci DD nazwa_miesiąca RRRR (np. 22 grudnia 2022r.).
'Copyright (c)by Jan T. Biernat
'
    Dim R As String
    Dim M As String
    Dim D As String
    Dim NazMie As String
    R = ""
    R = CStr(Year(Date))
    M = ""
    M = CStr(Month(Date))
    D = ""
    D = CStr(Day(Date))
    NazMie = ""
    Select Case M
        Case 1
            NazMie = "stycznia"
        Case 2
            NazMie = "lutego"
        Case 3
            NazMie = "marca"
        Case 4
            NazMie = "kwietnia"
        Case 5
            NazMie = "maja"
        Case 6
            NazMie = "czerwca"
        Case 7
            NazMie = "lipca"
        Case 8
            NazMie = "sierpnia"
        Case 9
            NazMie = "września"
        Case 10
            NazMie = "października"
        Case 11
            NazMie = "listopada"
        Case 12
            NazMie = "grudnia"
        Case Else
            NazMie = "?"
    End Select
    If (Len(D) = 1) Then
        D = "0" + D
    End If
    DataRMD2_TerazJest = D + " " + NazMie + " " + R + "r."
End Function

Sub PobierzDate()
    Cells(1, 1).Value = "Data:"
    Cells(1, 2).Value = "'" + DataRMD2_TerazJest()
End Sub