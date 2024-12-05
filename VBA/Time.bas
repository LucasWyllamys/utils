Attribute VB_Name = "Time"
Option Explicit

Public Function Aguardar(tempoEspera As String)   'Formato tempoespera: 00:00:00
    Dim tempo As Double
    tempo = Now + TimeValue(tempoEspera)
    Application.Wait tempo
End Function

Public Function Contador(indiceAtual As Long, indiceFim As Long, Optional etapaAtual As Integer, Optional etapaFim As Integer) As String
    If etapaAtual = 0 And etapaFim = 0 Then
        Contador = indiceAtual & " de " & indiceFim
    Else
        Contador = indiceAtual & " de " & indiceFim & " - " & etapaAtual & "/" & etapaFim
    End If
End Function

Public Function Percentual(indiceAtual As Long, indiceFim As Long) As Double
    Percentual = indiceAtual / indiceFim
End Function

Public Function Cronometro(tempoInicial As Date) As Date
    Cronometro = Now - tempoInicial
End Function


