Attribute VB_Name = "FormattedLog1"
Option Explicit

'================================================================================
' Módulo VBA: FormattedLog1
' Versão: 1.0.0
' Autor: Lucas Wyllamys Carmo da Silva
' Criado em: 23/02/2026
' Atualizado em: 23/02/2026
' Importar classe: ClsLogger: https://github.com/LucasWyllamys/utils/blob/main/VBA/clsLogger.cls
'================================================================================

Public Sub inicioLog(caminhoLog As String, log As Logger, qtdItens As Long, ti As Date, sistema As String)
    Dim psPID As Long
    
    FolderAndFileManipulator.CreateTextFile caminhoLog
    Set log = New Logger
    log.SetLogFilePath caminhoLog
    
    psPID = Planilha4.Range("A2").value ' Captura o PID da janela do PowerSheel aberta anteriormente
    Planilha4.Range("A2").value = log.CloseLogInRealTimeWMI(psPID) ' Fecha a janela do PowerSheel aberta anteriormente
    Planilha4.Range("A2").value = log.OpenLogInRealTime(10000) ' Abre uma janela do PowerSheel com as última 1000 linhas e atualiza o PID na célula do Excel
    
    Aguardar "00:00:05"
    
    With log
        .UnformattedMessage ""
        .UnformattedMessage "========================================================================"
        .UnformattedMessage "Sistema: " & sistema
        .UnformattedMessage "Lote: LOT" & Format$(ti, "yyyymmddhhnnss")
        .UnformattedMessage "Início: " & Format(ti, "yyyy-mm-dd hh:nn:ss")
        .UnformattedMessage "Qtd. de Itens: " & qtdItens
        .UnformattedMessage "------------------------------------------------------------------------"
    End With
End Sub

Public Sub linhaLog(log As Logger, linha As Long, qtdItens As Long, status As String, template As String)
    With log
        .UnformattedMessage ""
        .InfoLog "", linha - 4, qtdItens
        .UnformattedMessage "Status: " & status
        .UnformattedMessage "Linha: " & linha
        .UnformattedMessage "Template: " & template
    End With
End Sub

Public Sub fimLog(log As Logger, qtdItens As Long, ti As Date, num_executados As Long, num_naoExecutados As Long, num_falhas As Long)
    With log
        .UnformattedMessage ""
        .UnformattedMessage "------------------------------------------------------------------------"
        .UnformattedMessage "RESUMO"
        .UnformattedMessage "------------------------------------------------------------------------"
        .UnformattedMessage "Total Processados: " & qtdItens
        .UnformattedMessage "Executados: " & num_executados
        .UnformattedMessage "Não Executados: " & num_naoExecutados
        .UnformattedMessage "Falhas: " & num_falhas
        .UnformattedMessage "Fim: " & Format(ti, "yyyy-mm-dd hh:nn:ss")
        .UnformattedMessage "Tempo Total: " & Format(ti - Now(), "hh:nn:ss")
        .UnformattedMessage "========================================================================"
    End With
End Sub
