Function DataRMD_TerazJest() As String
'DataRMD_TerazJest - Funkcja zwraca datę systemową (bieżącą) w postaci RRRR-MM-DD.
'Copyright (c)by Jan T. Biernat
'
    Dim R As String
    Dim M As String
    Dim D As String
    R = ""
    R = CStr(Year(Date))
    M = ""
    M = CStr(Month(Date))
    D = ""
    D = CStr(Day(Date))
    If (Len(M) = 1) Then
        M = "0" + M
    End If
    If (Len(D) = 1) Then
        D = "0" + D
    End If
    DataRMD_TerazJest = R + "-" + M + "-" + D
End Function
 
Sub PobierzDate()
    Cells(1, 1).Value = "Data:"
    Cells(1, 2).Value = "'" + DataRMD_TerazJest()
End Sub