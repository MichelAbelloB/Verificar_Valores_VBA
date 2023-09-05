Sub LLamarfunciones()

    Call Cumplimientoenero
    Call Cumplimiento1enero
    Call Cumplimiento2enero
    Call Cumplimiento3enero
    
End Sub


Sub Cumplimiento1enero()
    Dim wsEnero As Worksheet
    Dim wsAns As Worksheet
    Dim ultimaFilaEnero As Long
    Dim valorUltimaFilaEnero As Variant
    
    Set wsEnero = ThisWorkbook.Sheets("ENERO")
    Set wsAns = ThisWorkbook.Sheets("TABLA ANS")
    
    ultimaFilaEnero = wsEnero.Cells(wsEnero.Rows.Count, "X").End(xlUp).Row
    
    valorUltimaFilaEnero = wsEnero.Cells(ultimaFilaEnero, "X").Value
    
    wsAns.Range("G4").Value = valorUltimaFilaEnero
End Sub

Sub Cumplimiento2enero()
    Dim wsEnero As Worksheet
    Dim wsAns As Worksheet
    Dim ultimaFilaEnero As Long
    Dim valorUltimaFilaEnero As Variant
    
    Set wsEnero = ThisWorkbook.Sheets("ENERO")
    Set wsAns = ThisWorkbook.Sheets("TABLA ANS")
    
    ultimaFilaEnero = wsEnero.Cells(wsEnero.Rows.Count, "P").End(xlUp).Row
    
    valorUltimaFilaEnero = wsEnero.Cells(ultimaFilaEnero, "P").Value
    
    wsAns.Range("G5").Value = valorUltimaFilaEnero
End Sub

Sub Cumplimiento3enero()
    Dim wsEnero As Worksheet
    Dim wsAns As Worksheet
    Dim ultimaFilaEnero As Long
    Dim valorUltimaFilaEnero As Variant
    
    Set wsEnero = ThisWorkbook.Sheets("ENERO")
    Set wsAns = ThisWorkbook.Sheets("TABLA ANS")
    
    ultimaFilaEnero = wsEnero.Cells(wsEnero.Rows.Count, "T").End(xlUp).Row
    
    valorUltimaFilaEnero = wsEnero.Cells(ultimaFilaEnero, "T").Value
    
    wsAns.Range("G6").Value = valorUltimaFilaEnero
End Sub
Sub Cumplimientoenero()
    Dim wsEnero As Worksheet
    Dim wsAns As Worksheet
    Dim ultimaFilaEnero As Long
    Dim valorUltimaFilaEnero As Variant
    
    Set wsEnero = ThisWorkbook.Sheets("ENERO")
    Set wsAns = ThisWorkbook.Sheets("TABLA ANS")
    
    ultimaFilaEnero = wsEnero.Cells(wsEnero.Rows.Count, "J").End(xlUp).Row
    
    valorUltimaFilaEnero = wsEnero.Cells(ultimaFilaEnero, "P").Value
    valorUltimaFilaEnero2 = wsEnero.Cells(ultimaFilaEnero, "J").Value
    
    promedio = (valorUltimaFilaEnero + valorUltimaFilaEnero2) / 2
    
    wsAns.Range("G3").Value = promedio
    
End Sub