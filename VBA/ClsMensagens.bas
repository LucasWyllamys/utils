Attribute VB_Name = "OrdenarTabela"
Option Explicit

Public Function msgConfirmacao() As Integer
    msgConfirmacao = MsgBox("Deseja prosseguir com a operação?", vbQuestion + vbYesNo)
End Function

Public Function msgOperacaoConcluida() As Integer
    msgOperacaoConcluida = MsgBox("Operação concluída!", vbInformation)
End Function
