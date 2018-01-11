Attribute VB_Name = "Module3"
Option Explicit
Option Base 1

' Funckje CheckLineXX - dla ka¿dej z XX/15 linii, spr,  czy wyst¹pi³ ci¹g tych samych symbolii,
' input - Tablica Screen, Symbol
' jeœli x5 to return 5,
' jeœli x4 to return 4,
' jeœli x3 to return 3,
' else 0

Function CheckLine01(screen() As Variant, symbol As Variant) As Integer
    CheckLine01 = 0
    
    'Line 1
        If (screen(1, 1) = symbol Or screen(1, 1) = "WILD") And (screen(1, 2) = symbol Or screen(1, 2) = "WILD") And (screen(1, 3) = symbol Or screen(1, 3) = "WILD") And (screen(1, 4) = symbol Or screen(1, 4) = "WILD") And (screen(1, 5) = symbol Or screen(1, 5) = "WILD") Then
        CheckLine01 = 5
    ElseIf (screen(1, 1) = symbol Or screen(1, 1) = "WILD") And (screen(1, 2) = symbol Or screen(1, 2) = "WILD") And (screen(1, 3) = symbol Or screen(1, 3) = "WILD") And (screen(1, 4) = symbol Or screen(1, 4) = "WILD") Then
        CheckLine01 = 4
    ElseIf (screen(1, 1) = symbol Or screen(1, 1) = "WILD") And (screen(1, 2) = symbol Or screen(1, 2) = "WILD") And (screen(1, 3) = symbol Or screen(1, 3) = "WILD") Then
        CheckLine01 = 3
    End If
End Function


Function CheckLine02(screen() As Variant, symbol As Variant) As Integer
     CheckLine02 = 0
    
    'Line 2
        If (screen(2, 1) = symbol Or screen(2, 1) = "WILD") And (screen(2, 2) = symbol Or screen(2, 2) = "WILD") And (screen(2, 3) = symbol Or screen(2, 3) = "WILD") And (screen(2, 4) = symbol Or screen(2, 4) = "WILD") And (screen(2, 5) = symbol Or screen(2, 5) = "WILD") Then
        CheckLine02 = 5
    ElseIf (screen(2, 1) = symbol Or screen(2, 1) = "WILD") And (screen(2, 2) = symbol Or screen(2, 2) = "WILD") And (screen(2, 3) = symbol Or screen(2, 3) = "WILD") And (screen(2, 4) = symbol Or screen(2, 4) = "WILD") Then
        CheckLine02 = 4
    ElseIf (screen(1, 1) = symbol Or screen(2, 1) = "WILD") And (screen(2, 2) = symbol Or screen(2, 2) = "WILD") And (screen(2, 3) = symbol Or screen(2, 3) = "WILD") Then
        CheckLine02 = 3
    End If
End Function


Function CheckLine03(screen() As Variant, symbol As Variant) As Integer
     CheckLine03 = 0

    'Line 3
        If (screen(3, 1) = symbol Or screen(3, 1) = "WILD") And (screen(3, 2) = symbol Or screen(3, 2) = "WILD") And (screen(3, 3) = symbol Or screen(3, 3) = "WILD") And (screen(3, 4) = symbol Or screen(3, 4) = "WILD") And (screen(3, 5) = symbol Or screen(3, 5) = "WILD") Then
        CheckLine03 = 5
    ElseIf (screen(3, 1) = symbol Or screen(3, 1) = "WILD") And (screen(3, 2) = symbol Or screen(3, 2) = "WILD") And (screen(3, 3) = symbol Or screen(3, 3) = "WILD") And (screen(3, 4) = symbol Or screen(3, 4) = "WILD") Then
        CheckLine03 = 4
    ElseIf (screen(2, 1) = symbol Or screen(2, 1) = "WILD") And (screen(2, 2) = symbol Or screen(2, 2) = "WILD") And (screen(2, 3) = symbol Or screen(2, 3) = "WILD") Then
        CheckLine03 = 3
    End If
End Function
    
    
Function CheckLine04(screen() As Variant, symbol As Variant) As Integer
     CheckLine04 = 0
    'Line 4

        If (screen(1, 1) = symbol Or screen(1, 1) = "WILD") And (screen(2, 2) = symbol Or screen(2, 2) = "WILD") And (screen(3, 3) = symbol Or screen(3, 3) = "WILD") And (screen(2, 4) = symbol Or screen(2, 4) = "WILD") And (screen(1, 5) = symbol Or screen(1, 5) = "WILD") Then
         CheckLine04 = 5
    ElseIf (screen(1, 1) = symbol Or screen(1, 1) = "WILD") And (screen(2, 2) = symbol Or screen(2, 2) = "WILD") And (screen(3, 3) = symbol Or screen(3, 3) = "WILD") And (screen(2, 4) = symbol Or screen(2, 4) = "WILD") Then
         CheckLine04 = 4
    ElseIf (screen(1, 1) = symbol Or screen(1, 1) = "WILD") And (screen(2, 2) = symbol Or screen(2, 2) = "WILD") And (screen(3, 3) = symbol Or screen(3, 3) = "WILD") Then
         CheckLine04 = 3
    End If
End Function


Function CheckLine05(screen() As Variant, symbol As Variant) As Integer
     CheckLine05 = 0
    'Line 5
    
        If (screen(3, 1) = symbol Or screen(3, 1) = "WILD") And (screen(2, 2) = symbol Or screen(2, 2) = "WILD") And (screen(1, 3) = symbol Or screen(1, 3) = "WILD") And (screen(2, 4) = symbol Or screen(2, 4) = "WILD") And (screen(3, 5) = symbol Or screen(3, 5) = "WILD") Then
         CheckLine05 = 5
    ElseIf (screen(3, 1) = symbol Or screen(3, 1) = "WILD") And (screen(2, 2) = symbol Or screen(2, 2) = "WILD") And (screen(1, 3) = symbol Or screen(1, 3) = "WILD") And (screen(2, 4) = symbol Or screen(2, 4) = "WILD") Then
         CheckLine05 = 4
    ElseIf (screen(3, 1) = symbol Or screen(3, 1) = "WILD") And (screen(2, 2) = symbol Or screen(2, 2) = "WILD") And (screen(1, 3) = symbol Or screen(1, 3) = "WILD") Then
         CheckLine05 = 3
    End If
End Function


Function CheckLine06(screen() As Variant, symbol As Variant) As Integer
     CheckLine06 = 0

    'Line 6
    
        If (screen(2, 1) = symbol Or screen(2, 1) = "WILD") And (screen(1, 2) = symbol Or screen(1, 2) = "WILD") And (screen(1, 3) = symbol Or screen(1, 3) = "WILD") And (screen(1, 4) = symbol Or screen(1, 4) = "WILD") And (screen(2, 5) = symbol Or screen(2, 5) = "WILD") Then
        CheckLine06 = 5
    ElseIf (screen(2, 1) = symbol Or screen(2, 1) = "WILD") And (screen(1, 2) = symbol Or screen(1, 2) = "WILD") And (screen(1, 3) = symbol Or screen(1, 3) = "WILD") And (screen(1, 4) = symbol Or screen(1, 4) = "WILD") Then
        CheckLine06 = 4
    ElseIf (screen(2, 1) = symbol Or screen(2, 1) = "WILD") And (screen(1, 2) = symbol Or screen(1, 2) = "WILD") And (screen(1, 3) = symbol Or screen(1, 3) = "WILD") Then
        CheckLine06 = 3
    End If
End Function
        
Function CheckLine07(screen() As Variant, symbol As Variant) As Integer
     CheckLine07 = 0
        
      'Line 7
      
        If (screen(2, 1) = symbol Or screen(2, 1) = "WILD") And (screen(3, 2) = symbol Or screen(3, 2) = "WILD") And (screen(3, 3) = symbol Or screen(3, 3) = "WILD") And (screen(3, 4) = symbol Or screen(3, 4) = "WILD") And (screen(2, 5) = symbol Or screen(2, 5) = "WILD") Then
        CheckLine07 = 5
    ElseIf (screen(2, 1) = symbol Or screen(2, 1) = "WILD") And (screen(3, 2) = symbol Or screen(3, 2) = "WILD") And (screen(3, 3) = symbol Or screen(3, 3) = "WILD") And (screen(3, 4) = symbol Or screen(3, 4) = "WILD") Then
        CheckLine07 = 4
    ElseIf (screen(2, 1) = symbol Or screen(2, 1) = "WILD") And (screen(3, 2) = symbol Or screen(3, 2) = "WILD") And (screen(3, 3) = symbol Or screen(3, 3) = "WILD") Then
        CheckLine07 = 3
    End If
End Function


Function CheckLine08(screen() As Variant, symbol As Variant) As Integer
     CheckLine08 = 0
    'Line 8
    
        If (screen(1, 1) = symbol Or screen(1, 1) = "WILD") And (screen(1, 2) = symbol Or screen(1, 2) = "WILD") And (screen(2, 3) = symbol Or screen(2, 3) = "WILD") And (screen(3, 4) = symbol Or screen(3, 4) = "WILD") And (screen(3, 5) = symbol Or screen(3, 5) = "WILD") Then
        CheckLine08 = 5
    ElseIf (screen(1, 1) = symbol Or screen(1, 1) = "WILD") And (screen(1, 2) = symbol Or screen(1, 2) = "WILD") And (screen(2, 3) = symbol Or screen(2, 3) = "WILD") And (screen(3, 4) = symbol Or screen(3, 4) = "WILD") Then
        CheckLine08 = 4
    ElseIf (screen(1, 1) = symbol Or screen(1, 1) = "WILD") And (screen(1, 2) = symbol Or screen(1, 2) = "WILD") And (screen(2, 3) = symbol Or screen(2, 3) = "WILD") Then
        CheckLine08 = 3
    End If
End Function


Function CheckLine09(screen() As Variant, symbol As Variant) As Integer
     CheckLine09 = 0
    'Line 9
    
        If (screen(3, 1) = symbol Or screen(3, 1) = "WILD") And (screen(3, 2) = symbol Or screen(3, 2) = "WILD") And (screen(2, 3) = symbol Or screen(2, 3) = "WILD") And (screen(1, 4) = symbol Or screen(1, 4) = "WILD") And (screen(1, 5) = symbol Or screen(1, 5) = "WILD") Then
        CheckLine09 = 5
    ElseIf (screen(3, 1) = symbol Or screen(3, 1) = "WILD") And (screen(3, 2) = symbol Or screen(3, 2) = "WILD") And (screen(2, 3) = symbol Or screen(2, 3) = "WILD") And (screen(1, 4) = symbol Or screen(1, 4) = "WILD") Then
        CheckLine09 = 4
    ElseIf (screen(3, 1) = symbol Or screen(3, 1) = "WILD") And (screen(3, 2) = symbol Or screen(3, 2) = "WILD") And (screen(2, 3) = symbol Or screen(2, 3) = "WILD") Then
        CheckLine09 = 3
    End If
End Function



Function CheckLine10(screen() As Variant, symbol As Variant) As Integer
     CheckLine10 = 0
    'Line 10
        
        If (screen(2, 1) = symbol Or screen(2, 1) = "WILD") And (screen(3, 2) = symbol Or screen(3, 2) = "WILD") And (screen(3, 3) = symbol Or screen(3, 3) = "WILD") And (screen(1, 4) = symbol Or screen(1, 4) = "WILD") And (screen(2, 5) = symbol Or screen(2, 5) = "WILD") Then
        CheckLine10 = 5
    ElseIf (screen(2, 1) = symbol Or screen(2, 1) = "WILD") And (screen(3, 2) = symbol Or screen(3, 2) = "WILD") And (screen(3, 3) = symbol Or screen(3, 3) = "WILD") And (screen(1, 4) = symbol Or screen(1, 4) = "WILD") Then
        CheckLine10 = 4
    ElseIf (screen(2, 1) = symbol Or screen(2, 1) = "WILD") And (screen(3, 2) = symbol Or screen(3, 2) = "WILD") And (screen(3, 3) = symbol Or screen(3, 3) = "WILD") Then
        CheckLine10 = 3
    End If
End Function


Function CheckLine11(screen() As Variant, symbol As Variant) As Integer
     CheckLine11 = 0
    'Line 11
    
        If (screen(2, 1) = symbol Or screen(2, 1) = "WILD") And (screen(1, 2) = symbol Or screen(1, 2) = "WILD") And (screen(2, 3) = symbol Or screen(2, 3) = "WILD") And (screen(3, 4) = symbol Or screen(3, 4) = "WILD") And (screen(2, 5) = symbol Or screen(2, 5) = "WILD") Then
        CheckLine11 = 5
    ElseIf (screen(2, 1) = symbol Or screen(2, 1) = "WILD") And (screen(1, 2) = symbol Or screen(1, 2) = "WILD") And (screen(2, 3) = symbol Or screen(2, 3) = "WILD") And (screen(3, 4) = symbol Or screen(3, 4) = "WILD") Then
        CheckLine11 = 4
    ElseIf (screen(2, 1) = symbol Or screen(2, 1) = "WILD") And (screen(1, 2) = symbol Or screen(1, 2) = "WILD") And (screen(2, 3) = symbol Or screen(2, 3) = "WILD") Then
        CheckLine11 = 3
    End If
End Function


Function CheckLine12(screen() As Variant, symbol As Variant) As Integer
     CheckLine12 = 0
    'Line 12

        If (screen(1, 1) = symbol Or screen(1, 1) = "WILD") And (screen(2, 2) = symbol Or screen(2, 2) = "WILD") And (screen(2, 3) = symbol Or screen(2, 3) = "WILD") And (screen(2, 4) = symbol Or screen(2, 4) = "WILD") And (screen(1, 5) = symbol Or screen(1, 5) = "WILD") Then
        CheckLine12 = 5
    ElseIf (screen(1, 1) = symbol Or screen(1, 1) = "WILD") And (screen(2, 2) = symbol Or screen(2, 2) = "WILD") And (screen(2, 3) = symbol Or screen(2, 3) = "WILD") And (screen(2, 4) = symbol Or screen(2, 4) = "WILD") Then
        CheckLine12 = 4
    ElseIf (screen(1, 1) = symbol Or screen(1, 1) = "WILD") And (screen(2, 2) = symbol Or screen(2, 2) = "WILD") And (screen(2, 3) = symbol Or screen(2, 3) = "WILD") Then
        CheckLine12 = 3
    End If
End Function


Function CheckLine13(screen() As Variant, symbol As Variant) As Integer
     CheckLine13 = 0
    'Line 13
    
        If (screen(3, 1) = symbol Or screen(3, 1) = "WILD") And (screen(2, 2) = symbol Or screen(2, 2) = "WILD") And (screen(2, 3) = symbol Or screen(2, 3) = "WILD") And (screen(2, 4) = symbol Or screen(2, 4) = "WILD") And (screen(3, 5) = symbol Or screen(3, 5) = "WILD") Then
        CheckLine13 = 5
    ElseIf (screen(3, 1) = symbol Or screen(3, 1) = "WILD") And (screen(2, 2) = symbol Or screen(2, 2) = "WILD") And (screen(2, 3) = symbol Or screen(2, 3) = "WILD") And (screen(2, 4) = symbol Or screen(2, 4) = "WILD") Then
        CheckLine13 = 4
    ElseIf (screen(3, 1) = symbol Or screen(3, 1) = "WILD") And (screen(2, 2) = symbol Or screen(2, 2) = "WILD") And (screen(2, 3) = symbol Or screen(2, 3) = "WILD") Then
        CheckLine13 = 3
    End If
End Function


Function CheckLine14(screen() As Variant, symbol As Variant) As Integer
     CheckLine14 = 0
    'Line 14
    
        If (screen(1, 1) = symbol Or screen(1, 1) = "WILD") And (screen(2, 2) = symbol Or screen(2, 2) = "WILD") And (screen(1, 3) = symbol Or screen(1, 3) = "WILD") And (screen(2, 4) = symbol Or screen(2, 4) = "WILD") And (screen(1, 5) = symbol Or screen(1, 5) = "WILD") Then
        CheckLine14 = 5
    ElseIf (screen(1, 1) = symbol Or screen(1, 1) = "WILD") And (screen(2, 2) = symbol Or screen(2, 2) = "WILD") And (screen(1, 3) = symbol Or screen(1, 3) = "WILD") And (screen(2, 4) = symbol Or screen(2, 4) = "WILD") Then
        CheckLine14 = 4
    ElseIf (screen(1, 1) = symbol Or screen(1, 1) = "WILD") And (screen(2, 2) = symbol Or screen(2, 2) = "WILD") And (screen(1, 3) = symbol Or screen(1, 3) = "WILD") Then
        CheckLine14 = 3
    End If
End Function


Function CheckLine15(screen() As Variant, symbol As Variant) As Integer
     CheckLine15 = 0
    'Line 15
    
        If (screen(3, 1) = symbol Or screen(3, 1) = "WILD") And (screen(2, 2) = symbol Or screen(2, 2) = "WILD") And (screen(3, 3) = symbol Or screen(3, 3) = "WILD") And (screen(2, 4) = symbol Or screen(2, 4) = "WILD") And (screen(3, 5) = symbol Or screen(3, 5) = "WILD") Then
        CheckLine15 = 5
    ElseIf (screen(3, 1) = symbol Or screen(3, 1) = "WILD") And (screen(2, 2) = symbol Or screen(2, 2) = "WILD") And (screen(3, 3) = symbol Or screen(3, 3) = "WILD") And (screen(2, 4) = symbol Or screen(2, 4) = "WILD") Then
        CheckLine15 = 4
    ElseIf (screen(3, 1) = symbol Or screen(3, 1) = "WILD") And (screen(2, 2) = symbol Or screen(2, 2) = "WILD") And (screen(3, 3) = symbol Or screen(3, 3) = "WILD") Then
        CheckLine15 = 3
    End If
End Function


