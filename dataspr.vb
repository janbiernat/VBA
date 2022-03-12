Function DataSpr(Data_RMD As String) As String 
'DataSpr - Funkcja sprawdza poprawność wprowadzonej daty w formacie RRRR-MM-DD. 
'Copyright (c)by Jan T. Biernat 
' 
'Wywołanie funkcji: Cells(1, 2).Value = "'" + DataSpr(Cells(1, 1).Value) 
' 
    Dim R As Integer 
    Dim M As Integer 
    Dim D As Integer 
    Dim DniIlosc As Integer 
    Dim MM As String 
    Dim DD As String 
    Dim T As String 
    R = 0 
    M = 0 
    D = 0 
    DniIlosc = 0 
    MM = "" 
    DD = "" 
    T = "" 
    Data_RMD = Trim(Mid(Data_RMD, 1, 10)) 
    If (Data_RMD <> "") Then 
        If (IsNumeric(Mid(Data_RMD, 1, 4)) = True) Then 
            R = CInt(Mid(Data_RMD, 1, 4)) 
        End If 
        If (IsNumeric(Mid(Data_RMD, 6, 2)) = True) Then 
            M = CInt(Mid(Data_RMD, 6, 2)) 
        End If 
        If (IsNumeric(Mid(Data_RMD, 9, 2)) = True) Then 
            D = CInt(Mid(Data_RMD, 9, 2)) 
        End If 
        If (R > 1947) Then 
            If ((M > 0) And (M < 13)) Then 
                If ((M = 1) Or (M = 3) Or (M = 5) Or (M = 7) Or (M = 8) Or (M = 10) Or (M = 12)) Then '31 dni. 
                    DniIlosc = 31 
                ElseIf ((M = 4) Or (M = 6) Or (M = 9) Or (M = 11)) Then '30 dni. 
                        DniIlosc = 30 
                    ElseIf (M = 2) Then 
                            If (((R Mod 4 = 0) And (R Mod 100 <> 0)) Or (R Mod 400 = 0)) Then 
                                DniIlosc = 29 
                            Else 
                                DniIlosc = 28 
                            End If 
                        End If 
                If ((D > 0) And (D < DniIlosc + 1)) Then 
                    MM = CStr(M) 
                    DD = CStr(D) 
                    If (Len(MM) = 1) Then 
                        MM = "0" + MM 
                    End If 
                    If (Len(DD) = 1) Then 
                        DD = "0" + DD 
                    End If 
                    T = CStr(R) + "-" + MM + "-" + DD 
                End If 
            End If 
        End If 
    End If 
    DataSpr = T 
End Function 