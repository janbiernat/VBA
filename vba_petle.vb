Sub Petle()
'--== Pętla FOR, DO ... LOOP UNTIL() ==--
'Copyright (c)by Jan T. Biernat
'
    'Pętla FOR.
    Dim A As Integer                            '1.
    Dim B As Integer                            '2.
    A = 0                                       '3.
    B = 0                                       '4.
    For A = 1 To 10                             '5.
        Cells(1, 1 + A).Value = A               '6.
        Cells(1 + A, 1).Value = A               '7.
        For B = 1 To 10                         '8.
            Cells(1 + A, 1 + B).Value = (A * B) '9.
        Next B                                  '8.
    Next A                                      '5.
    '
    'Pętla DO ... LOOP UNTIL(warunek).
    A = 0                                           '10.
    Do                                              '11.
        A = A + 1                                   '12.
        Cells(13, 1 + A).Value = A                  '13.
        Cells(13 + A, 1).Value = A                  '14.
        B = 0                                       '15.
        Do                                          '16.
            B = B + 1                               '17.
            Cells(13 + A, 1 + B).Value = (A * B)    '18.
        Loop Until (B > 9)                          '16.
    Loop Until (A > 9)                              '11.
    Range("A1").Select                              '19.
'
'Legenda:
'1. ... .
End Sub