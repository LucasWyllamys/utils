Attribute VB_Name = "OrdenarTabela"
Option Explicit

Public Function msgConfirmacao() As Integer
    msgConfirmacao = MsgBox("Deseja prosseguir com a opera��o?", vbQuestion + vbYesNo)
End Function

Public Function msgOperacaoConcluida() As Integer
    msgOperacaoConcluida = MsgBox("Opera��o conclu�da!", vbInformation)
End Function
