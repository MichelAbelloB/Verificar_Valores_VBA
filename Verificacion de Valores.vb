'/////////////////////////////////////// Verificacion de "A" /////////////////////////////////////// 
Function CalcularValor(celdaI As Range, celdaE As Range) As Variant
    Dim valorE As Double
    valorE = celdaE.Value
    
    If celdaI.Value = "Grupo 3" Then
        If valorE < 0.240277777777778 Then
            CalcularValor = 1
        ElseIf valorE >= 0.240277777777778 And valorE <= 0.246527778 Then
            CalcularValor = 0.9
        Else
            CalcularValor = 0
        End If
    ElseIf celdaI.Value = "Grupo 2" Then
        If valorE < 0.906944444444444 Then
            CalcularValor = 0
        ElseIf valorE >= 0.906944444444444 And valorE <= 0.913194444 Then
            CalcularValor = 0.9
        Else
            CalcularValor = 1
        End If
    ElseIf celdaI.Value = "Grupo 1" Then
        If valorE < 0.197916667 Then
            CalcularValor = 0
        ElseIf valorE >= 0.197916667 And valorE <= 0.204861111 Then
            CalcularValor = 0.9
        Else
            CalcularValor = 1
        End If
    Else
        CalcularValor = 0
    End If
End Function
'/////////////////////////////////////// Verificacion de "B" /////////////////////////////////////// 
Function VerificarValorB(celdaI As Range) As Variant
    Dim valorI As Double
    valorI = celdaI.Value
    
    If valorI < 0.573611111111111 Then
        VerificarValorB = 1
    ElseIf valorI >= 0.573611111111111 And valorI <= 0.579861111 Then
        VerificarValorB = 0.9
    Else
        VerificarValorB = 0
    End If
End Function
'/////////////////////////////////////// Verificacion de "C" /////////////////////////////////////// 
Function VerificarValorC(celdaI As Range) As Variant
    Dim valorI As Double
    valorI = celdaI.Value
    
    If valorI < 0.906944444444444 Then
        VerificarValorC = 1
    ElseIf valorI >= 0.906944444444444 And valorI <= 0.913194444 Then
        VerificarValorC = 0.9
    Else
        VerificarValorC = 0
    End If
End Function