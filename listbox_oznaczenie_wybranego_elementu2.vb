Function TekstOznacz(S As String) 
'TekstOznacz - Funkcja oznacza tekst, umieszczając znak ">" na początku podanego ciągu znaków. 
'Copyright (c)by Jan T. Biernat 
' 
    If (Left(S, 1) = ">") Then 
        TekstOznacz = " " + Mid(S, 2, Len(S)) 
    Else 
        TekstOznacz = ">" + Mid(S, 2, Len(S)) 
    End If 
End Function 
 
Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean) 
    With ListBox1 
        .List(.ListIndex, 0) = TekstOznacz(.List(.ListIndex, 0)) 
    End With 
End Sub 