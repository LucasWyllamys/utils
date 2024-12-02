Attribute VB_Name = "Time"
Option Explicit

Public Function Aguardar(tempoEspera As String)   'Formato tempoespera: 00:00:00
    Dim tempo As Double
    tempo = Now + TimeValue(tempoEspera)
    Application.Wait tempo
End Function

Public Function Contador(indiceAtual As Long, indiceFim As Long) As String
    Contador = indiceAtual & " de " & indiceFim
End Function

Public Function Percentual(indiceAtual As Long, indiceFim As Long) As Double
    Percentual = indiceAtual / indiceFim
End Function

Public Function Cronometro(tempoInicial As Date) As Date
    Cronometro = Now - tempoInicial
End Function


