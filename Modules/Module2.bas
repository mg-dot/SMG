Attribute VB_Name = "Module2"
Option Explicit
Option Base 1

Private Function Modulus_Operator(Value1, Value2)

' Funkcja operatora modulo

    Modulus_Operator = Value1 - (Int(Value1 / Value2) * Value2)

End Function


Function Generator(tab01() As Variant, max As Long, number As Integer)

' Generator liczb pseudolosowych
' Funkcja wpisuje do tablicy tab01 max liczb ca³kowitych, pseudolosowych z przedzia³u od 1 do number

    Dim i As Long
    
    For i = 1 To max
        tab01(i) = Modulus_Operator(genrand_int32(), number) + 1
    Next i
    
End Function


Function CopyReels(tab02() As Variant, max_stop)

'Funkcja kopiuje Reele do tablicy tab02
'B2:F32 z Szitu Reels

    Dim FromWsh As Worksheet
    With ThisWorkbook
            Set FromWsh = .Worksheets("Reels")
    End With
    
    tab02 = FromWsh.Range("B" & 2 & ":F" & max_stop + 2)
    
End Function


Function CopyPayTable(tab03() As Variant)

'Funkcja kopiuje PayTable do tablicy tab03
'B2:D9 z Szitu PayTable
    Dim FromWsh As Worksheet
    With ThisWorkbook
            Set FromWsh = .Worksheets("PayTable")
    End With
    
    tab03 = FromWsh.Range("B" & 2 & ":D" & 9)
    
End Function

Function CopyStopsinReels(tab04())

'Funkcja kopiuje PayTable do tablicy tab03
'B2:D9 z Szitu PayTable
    Dim FromWsh As Worksheet
    With ThisWorkbook
            Set FromWsh = .Worksheets("Reels")
    End With
    
    tab04 = FromWsh.Range("N" & 2 & ":N" & 6)
    
End Function




