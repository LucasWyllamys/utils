Attribute VB_Name = "OrdenarTabela"
Option Explicit

Public Sub Ordenar()
    Dim msg As New ClsMensagens
    If msg.msgConfirmacao = vbNo Then Exit Sub
    ThisWorkbook.Activate
    With Planilha3
        .Range("E1").Activate                      ' Obrigatório selecionar uma célula da tabela/intervalo!
        If .FilterMode = True Then .ShowAllData    ' Limpar todos os filtros caso haja algum.
        .ListObjects("BaseClientesCGR").Sort.SortFields.Clear
        .ListObjects("BaseClientesCGR").Sort.SortFields.Add2 Key:=Range("BaseClientesCGR[[#All],[Máxima Demanda / Montante]]"), _
            SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        With .ListObjects("BaseClientesCGR").Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End With
    Call msg.msgOperacaoConcluida
End Sub

