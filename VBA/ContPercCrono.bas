Attribute VB_Name = "ContPercCrono"
Option Explicit

Function Contador(Indice As Long, TotalRegistros As Long) As String
    Contador = Indice & " de " & TotalRegistros
End Function

Function Percentual(Indice As Long, TotalRegistros As Long) As Double
    Percentual = Indice / TotalRegistros
End Function

Function Cronometro(TempoInicial As Date) As Date
    Cronometro = Now - TempoInicial
End Function

Function LimparContador()
    Planilha3.Range("B1").Value = Empty
End Function

Function LimparPercentual()
    Planilha3.Range("B2").Value = Empty
End Function

Function LimparCronometro()
    Planilha3.Range("B3").Value = Empty
End Function
