Function cumplimeintoejecutados(celdaI As Range) As Variant
    Dim valorI As Double
    valorI = celdaI.Value
    
    If valorI = 1 Then
        cumplimeintoejecutados = 1
    Else
        cumplimeintoejecutados = 0
    End If
End Function

Function cumplimeintoPQRS(celdaI As Range) As Variant
    Dim valorI As Double
    valorI = celdaI.Value
    
    If valorI < 10 Then
        cumplimeintoPQRS = 1
    ElseIf valorI >= 10 And valorI <= 15 Then
        cumplimeintoPQRS = 0.9
    Else
        cumplimeintoPQRS = 0
    End If
End Function
