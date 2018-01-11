Attribute VB_Name = "Module4"
Option Explicit
Option Base 1

Sub simulation()

'Symulacja gry
    
'Stats
    Dim StartTime As Double
    Dim MinutesElapsed As String
    
    StartTime = Timer
    
'Vars const
    Const max_spin As Long = 100         'Liczba spinów w symulacji
    Const max_stop As Integer = 30
    Const symbols_number As Integer = 8 'Liczba symboli
    Const reels_number As Integer = 5 'Liczba Reeli
    
'Vars
    Dim reels() As Variant
    Dim paytable() As Variant
    Dim reel_stops() As Variant
    
    'Dim randGen_table(1 To reels_number, 1 To max_spin) As Variant
    
    
    Dim randGen1(), randGen2(), randGen3(), randGen4(), randGen5() As Variant
    
    Dim screen(1 To 3, 1 To reels_number) As Variant
    Dim symbols(1 To symbols_number) As String
    
    Dim hits(1 To symbols_number, 1 To 3) As Long    ' Output
    Dim balance(1 To max_spin) As Long               ' Output
    Dim wins(1 To max_spin) As Long                  ' Output
    
    balance(1) = 0
    wins(1) = 0
    
    Call CopyStopsinReels(reel_stops)
    Call CopyReels(reels, max_stop)     ' Kopiowanie Reelsów
    Call CopyPayTable(paytable)         ' Kopiowanie PayTable

   ' Call Generator(randGen1(), max_spin, reel_stops(1))  ' Losowanie max liczb i zapisanie do tablicy RandGen()
   ' Call Generator(randGen2(), max_spin, reel_stops(2))
   ' Call Generator(randGen3(), max_spin, reel_stops(3))
   ' Call Generator(randGen4(), max_spin, reel_stops(4))
   ' Call Generator(randGen5(), max_spin, reel_stops(5))

    
   
    Dim i As Long
    Dim k, l, j As Integer
    Dim k1(1 To 5), k2(1 To 5), k3(1 To 5) As Variant
    
    symbols(1) = "WILD"
    symbols(2) = "seven"
    symbols(3) = "apple"
    symbols(4) = "orange"
    symbols(5) = "watermelon"
    symbols(6) = "plum"
    symbols(7) = "lemon"
    symbols(8) = "cherry"
    
'Let's Game
    
    
    
    For i = 1 To max_spin
        
        balance(i) = balance(i - 1) - 100 ' Ka¿dy spin = -100 na balance
        
        ' Wybieramy i-ta wylosowana liczbe
        For j = 1 To 5
            k1(j) = randGen1(j)
            
            If k1(j) = max_stop - 1 Then
                k2(j) = max_stop
                k3(j) = 1
            ElseIf k1(j) = max_stop Then
                k2(j) = 1
                k3(j) = 2
            Else
                k2(j) = (k1(j) + 1)
                k3(j) = (k2(j) + 1)
            End If
        Next j
        ' Ustawiamy Screen 3x5
        
        For k = 1 To 5
                screen(1, k) = reels(k1(j), k)
                screen(2, k) = reels(k2(j), k)
                screen(3, k) = reels(k3(j), k)
        Next k
                
        For l = 1 To 8
        
            Select Case CheckLine01(screen(), symbols(l))
                Case 5
                    hits(l, 3) = hits(l, 3) + 1
                    balance(i) = balance(i) + paytable(l, 1)
                Case 4
                    hits(l, 2) = hits(l, 2) + 1
                    balance(i) = balance(i) + paytable(l, 2)
                Case 3
                    hits(l, 1) = hits(l, 1) + 1
                    balance(i) = balance(i) + paytable(l, 3)
                End Select
                
            Select Case CheckLine02(screen(), symbols(l))
                Case 5
                    hits(l, 3) = hits(l, 3) + 1
                    balance(i) = balance(i) + paytable(l, 1)
                Case 4
                    hits(l, 2) = hits(l, 2) + 1
                    balance(i) = balance(i) + paytable(l, 2)
                Case 3
                    hits(l, 1) = hits(l, 1) + 1
                    balance(i) = balance(i) + paytable(l, 3)
                End Select
            Select Case CheckLine03(screen(), symbols(l))
                Case 5
                    hits(l, 3) = hits(l, 3) + 1
                    balance(i) = balance(i) + paytable(l, 1)
                Case 4
                    hits(l, 2) = hits(l, 2) + 1
                    balance(i) = balance(i) + paytable(l, 2)
                Case 3
                    hits(l, 1) = hits(l, 1) + 1
                    balance(i) = balance(i) + paytable(l, 3)
                End Select
            Select Case CheckLine04(screen(), symbols(l))
                Case 5
                    hits(l, 3) = hits(l, 3) + 1
                    balance(i) = balance(i) + paytable(l, 1)
                Case 4
                    hits(l, 2) = hits(l, 2) + 1
                    balance(i) = balance(i) + paytable(l, 2)
                Case 3
                    hits(l, 1) = hits(l, 1) + 1
                    balance(i) = balance(i) + paytable(l, 3)
                End Select
                
            Select Case CheckLine05(screen(), symbols(l))
                Case 5
                    hits(l, 3) = hits(l, 3) + 1
                    balance(i) = balance(i) + paytable(l, 1)
                Case 4
                    hits(l, 2) = hits(l, 2) + 1
                    balance(i) = balance(i) + paytable(l, 2)
                Case 3
                    hits(l, 1) = hits(l, 1) + 1
                    balance(i) = balance(i) + paytable(l, 3)
                End Select
            Select Case CheckLine06(screen(), symbols(l))
                Case 5
                    hits(l, 3) = hits(l, 3) + 1
                    balance(i) = balance(i) + paytable(l, 1)
                Case 4
                    hits(l, 2) = hits(l, 2) + 1
                    balance(i) = balance(i) + paytable(l, 2)
                Case 3
                    hits(l, 1) = hits(l, 1) + 1
                    balance(i) = balance(i) + paytable(l, 3)
                End Select
            Select Case CheckLine07(screen(), symbols(l))
                Case 5
                    hits(l, 3) = hits(l, 3) + 1
                    balance(i) = balance(i) + paytable(l, 1)
                Case 4
                    hits(l, 2) = hits(l, 2) + 1
                    balance(i) = balance(i) + paytable(l, 2)
                Case 3
                    hits(l, 1) = hits(l, 1) + 1
                    balance(i) = balance(i) + paytable(l, 3)
                End Select
                
            Select Case CheckLine08(screen(), symbols(l))
                Case 5
                    hits(l, 3) = hits(l, 3) + 1
                    balance(i) = balance(i) + paytable(l, 1)
                Case 4
                    hits(l, 2) = hits(l, 2) + 1
                    balance(i) = balance(i) + paytable(l, 2)
                Case 3
                    hits(l, 1) = hits(l, 1) + 1
                    balance(i) = balance(i) + paytable(l, 3)
                End Select
            Select Case CheckLine09(screen(), symbols(l))
                Case 5
                    hits(l, 3) = hits(l, 3) + 1
                    balance(i) = balance(i) + paytable(l, 1)
                Case 4
                    hits(l, 2) = hits(l, 2) + 1
                    balance(i) = balance(i) + paytable(l, 2)
                Case 3
                    hits(l, 1) = hits(l, 1) + 1
                    balance(i) = balance(i) + paytable(l, 3)
                End Select
            Select Case CheckLine10(screen(), symbols(l))
                Case 5
                    hits(l, 3) = hits(l, 3) + 1
                    balance(i) = balance(i) + paytable(l, 1)
                Case 4
                    hits(l, 2) = hits(l, 2) + 1
                    balance(i) = balance(i) + paytable(l, 2)
                Case 3
                    hits(l, 1) = hits(l, 1) + 1
                    balance(i) = balance(i) + paytable(l, 3)
                End Select
                
            Select Case CheckLine11(screen(), symbols(l))
                Case 5
                    hits(l, 3) = hits(l, 3) + 1
                    balance(i) = balance(i) + paytable(l, 1)
                Case 4
                    hits(l, 2) = hits(l, 2) + 1
                    balance(i) = balance(i) + paytable(l, 2)
                Case 3
                    hits(l, 1) = hits(l, 1) + 1
                    balance(i) = balance(i) + paytable(l, 3)
                End Select
            Select Case CheckLine12(screen(), symbols(l))
                Case 5
                    hits(l, 3) = hits(l, 3) + 1
                    balance(i) = balance(i) + paytable(l, 1)
                Case 4
                    hits(l, 2) = hits(l, 2) + 1
                    balance(i) = balance(i) + paytable(l, 2)
                Case 3
                    hits(l, 1) = hits(l, 1) + 1
                    balance(i) = balance(i) + paytable(l, 3)
                End Select
            Select Case CheckLine13(screen(), symbols(l))
                Case 5
                    hits(l, 3) = hits(l, 3) + 1
                    balance(i) = balance(i) + paytable(l, 1)
                Case 4
                    hits(l, 2) = hits(l, 2) + 1
                    balance(i) = balance(i) + paytable(l, 2)
                Case 3
                    hits(l, 1) = hits(l, 1) + 1
                    balance(i) = balance(i) + paytable(l, 3)
                End Select
            Select Case CheckLine14(screen(), symbols(l))
                Case 5
                    hits(l, 3) = hits(l, 3) + 1
                    balance(i) = balance(i) + paytable(l, 1)
                Case 4
                    hits(l, 2) = hits(l, 2) + 1
                    balance(i) = balance(i) + paytable(l, 2)
                Case 3
                    hits(l, 1) = hits(l, 1) + 1
                    balance(i) = balance(i) + paytable(l, 3)
                End Select
            Select Case CheckLine15(screen(), symbols(l))
                Case 5
                    hits(l, 3) = hits(l, 3) + 1
                    balance(i) = balance(i) + paytable(l, 1)
                Case 4
                    hits(l, 2) = hits(l, 2) + 1
                    balance(i) = balance(i) + paytable(l, 2)
                Case 3
                    hits(l, 1) = hits(l, 1) + 1
                    balance(i) = balance(i) + paytable(l, 3)
                End Select
                
            Next l
        
        If balance(i) = balance(i - 1) - 100 Then
            wins(i) = wins(i - 1)
        Else: wins(i) = wins(i - 1) + 1
        End If
        
        
        Next i
        
        
        
'Output

    Dim ToWsh As Worksheet
    With ThisWorkbook
            Set ToWsh = .Worksheets("SimOut")
    End With
    
    ToWsh.Range("B5:D12") = hits()
    ToWsh.Range("G5:G" & max_spin + 3) = Application.Transpose(balance())
    ToWsh.Range("H5:H" & max_spin + 3) = Application.Transpose(wins())
    ToWsh.Range("V37") = max_spin
    ToWsh.Range("V38") = balance(max_spin)
    ToWsh.Range("V39") = wins(max_spin)
'Finish
    Dim rtp As Double
    rtp = ToWsh.Range("AA29")
    Dim hitsOfall
    hitsOfall = wins(max_spin) / max_spin
    
    
    
    MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")
    MsgBox "This code ran successfully in " & MinutesElapsed & " minutes", vbInformation
    MsgBox max_spin & " spins" & vbNewLine & "RTP: " & rtp * 100 & "%" & vbNewLine & "Wins: " & hitsOfall * 100 & "%"

End Sub
