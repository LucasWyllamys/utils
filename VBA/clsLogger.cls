VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==========================================================================================
' Classe VBA: clsLogger - Gravação de logs
' Versão: 1.0.0
' Autor: Lucas Wyllamys Carmo da Silva
' Criado em: 17/10/2025
'==========================================================================================
'
' 1. DESCRIÇÃO
'   - Esta classe permite a gravação de registros em um arquivo de log (.log), com diferentes níveis de severidade (DEBUG, INFO, WARN, ERROR, CRITICAL), facilitando o diagnóstico e rastreamento de eventos e erros em aplicações VBA.
'
' 2. PRÉ-REQUISITOS
'   - O arquivo do projeto VBA onde esta classe for importada deve estar salvo em um diretório.
'   - O caminho padrão do log será definido com base no diretório do arquivo Excel (ThisWorkbook.Path).
'
' 3. ESTRUTURA INTERNA
'   3.1. Propriedades encapsuladas:
'       - pCaminhoArquivoLog: Armazena o caminho completo do arquivo de log
'
'   3.2. Constantes de erro:
'       - ERR_LOG_INIT_FAILED: Código de erro para falha na inicialização do caminho do log
'       - ERR_LOG_WRITE_FAILED: Código de erro para falha na escrita do log
'
' 4. MÉTODOS PÚBLICOS
'   4.1. DefinirCaminhoArquivoLog(Optional caminhoArquivo As String = "")
'       - Descrição: Permite definir um caminho personalizado para o arquivo de log.
'       - Parâmetros:
'           - caminhoArquivo: caminho completo do arquivo de log (opcional).
'       - Comportamento:
'           - Se informado, valida se o diretório existe e atualiza o caminho do log.
'           - Se não informado, mantém o caminho padrão definido na inicialização.
'
'   4.2. LogDebug(mensagem As String, Optional etapa As Integer, Optional total As Integer)
'        - Descrição: Registra uma mensagem de depuração (diagnóstico técnico).
'        - Parâmetros:
'            - mensagem: texto da mensagem a ser registrada.
'            - etapa: número da etapa atual (opcional).
'            - total: número total de etapas (opcional).
'        - Comportamento:
'            - Gera uma linha de log com o tipo "DEBUG" e registra no arquivo.
'
'   4.3. LogInfo(mensagem As String, Optional etapa As Integer, Optional total As Integer)
'        - Descrição: Registra uma mensagem informativa (eventos importantes).
'        - Parâmetros: iguais ao método LogDebug.
'        - Comportamento: tipo "INFO".
'
'   4.4. LogWarn(mensagem As String, Optional etapa As Integer, Optional total As Integer)
'        - Descrição: Registra um alerta (possíveis problemas).
'        - Parâmetros: iguais ao método LogDebug.
'        - Comportamento: tipo "WARN".
'
'   4.5. LogError(mensagem As String, Optional etapa As Integer, Optional total As Integer)
'        - Descrição: Registra um erro (falhas que afetam funcionalidades).
'        - Parâmetros: iguais ao método LogDebug.
'        - Comportamento: tipo "ERROR".
'
'   4.6. LogCritical(mensagem As String, Optional etapa As Integer, Optional total As Integer)
'        - Descrição: Registra um erro crítico (falhas graves que comprometem o sistema).
'        - Parâmetros: iguais ao método LogDebug.
'        - Comportamento: tipo "CRITICAL".
'
' 5. MÉTODOS PRIVADOS
'   5.1. MontarLinhaLog(tipo As String, mensagem As String, Optional etapa As Integer = 0, Optional total As Integer = 0) As String
'       - Descrição: Monta uma linha de log formatada com data, tipo, etapa e mensagem.
'       - Comportamento:
'           - Formata a data/hora atual.
'           - Inclui tipo de log e, se aplicável, progresso (etapa/total).
'           - Retorna a linha pronta para registro.
'
'   5.2. Registrar(mensagem As String)
'       - Descrição: Escreve a linha de log no arquivo especificado.
'       - Comportamento:
'           - Abre o arquivo no modo Append.
'           - Escreve a mensagem.
'           - Fecha o arquivo.
'           - Em caso de erro, lança exceção com código ERR_LOG_WRITE_FAILED.
'
' 6. EXEMPLO DE USO
'    Dim logger As clsLogger ' Declaração do objeto
'    Set logger = New clsLogger ' Instanciamento do objeto
'
'    logger.DefinirCaminhoArquivoLog "C:\Logs\meu_log.log" ' Define caminho personalizado
'
'    logger.LogDebug "Registra uma mensagem de depuração (diagnóstico técnico)"
'    logger.LogInfo " Registra uma mensagem informativa (eventos importantes)"
'    logger.LogWarn "Registra um alerta (possíveis problemas)"
'    logger.LogError "Registra um erro (falhas que afetam funcionalidades)"
'    logger.LogCritical "Registra um erro crítico (falhas graves que comprometem o sistema)"
'
' 7. OBSERVAÇÕES
'   - O arquivo de log é criado automaticamente se não existir.
'   - O caminho do log deve apontar para um diretório válido.
'   - A classe pode ser utilizada em qualquer projeto VBA que precise de rastreamento de eventos.
'   - Ideal para depuração, auditoria e monitoramento de processos automatizados.

'============================================ Constantes de Erro ============================================

Private Const ERR_LOG_INIT_FAILED As Long = 2000 ' Erro ao inicializar o caminho do arquivo de log
Private Const ERR_LOG_WRITE_FAILED As Long = 2001 ' Erro ao tentar escrever no arquivo de log

'=========================================== Propriedades Privadas ===========================================

Private pCaminhoArquivoLog As String ' Armazena o caminho completo do arquivo de log

'=========================================== Métodos Assessores ===========================================

Private Property Get CaminhoArquivoLog() As String: CaminhoArquivoLog = pCaminhoArquivoLog: End Property ' Retorna o caminho atual do arquivo de log

Private Property Let CaminhoArquivoLog(value As String)
    ' Define o caminho do arquivo de log, validando se o diretório existe
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(fso.GetParentFolderName(value)) Then
        pCaminhoArquivoLog = value
    Else
        Err.Raise ERR_LOG_INIT_FAILED, "clsLogger", "Diretório do caminho de log não existe." ' Lançamento de erro
    End If
End Property

'=========================================== Inicialização ===========================================

Private Sub Class_Initialize()
    ' Define o caminho padrão do arquivo de log ao instanciar a classe
    On Error GoTo TratarErro
    CaminhoArquivoLog = ThisWorkbook.Path & "\log.log" ' Caminho de arquivo log padrão
    Exit Sub
TratarErro:
    Err.Raise ERR_LOG_INIT_FAILED, "clsLogger", "Erro ao definir caminho do arquivo de log: " & Err.Description
End Sub

'=========================================== Métodos Públicos ===========================================

Public Sub DefinirCaminhoArquivoLog(Optional caminhoArquivo As String = "")
    ' Permite definir um caminho personalizado para o arquivo de log
    If caminhoArquivo <> "" Then CaminhoArquivoLog = caminhoArquivo
End Sub

Public Sub LogDebug(mensagem As String, Optional etapa As Integer, Optional total As Integer)
    ' Registra uma mensagem de depuração (diagnóstico técnico)
    Registrar MontarLinhaLog("DEBUG", mensagem, etapa, total)
End Sub

Public Sub LogInfo(mensagem As String, Optional etapa As Integer, Optional total As Integer)
    ' Registra uma mensagem informativa (eventos importantes)
    Registrar MontarLinhaLog("INFO", mensagem, etapa, total)
End Sub

Public Sub LogWarn(mensagem As String, Optional etapa As Integer, Optional total As Integer)
    ' Registra um alerta (possíveis problemas)
    Registrar MontarLinhaLog("WARN", mensagem, etapa, total)
End Sub

Public Sub LogError(mensagem As String, Optional etapa As Integer, Optional total As Integer)
    ' Registra um erro (falhas que afetam funcionalidades)
    Registrar MontarLinhaLog("ERROR", mensagem, etapa, total)
End Sub

Public Sub LogCritical(mensagem As String, Optional etapa As Integer, Optional total As Integer)
    ' Registra um erro crítico (falhas graves que comprometem o sistema)
    Registrar MontarLinhaLog("CRITICAL", mensagem, etapa, total)
End Sub

'=========================================== Métodos Privados ===========================================

Private Function MontarLinhaLog(tipo As String, mensagem As String, Optional etapa As Integer = 0, Optional total As Integer = 0) As String
    ' Monta a linha de log formatada com data, tipo, etapa e mensagem
    Dim linhaLog As String
    
    linhaLog = "[" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "][" & tipo & "]"
    If etapa <> 0 And total <> 0 Then
        linhaLog = linhaLog & "[" & Format(etapa, "0000") & "/" & Format(total, "0000") & "]"
    End If
    linhaLog = linhaLog & " " & mensagem
    
    MontarLinhaLog = linhaLog
End Function

Private Sub Registrar(mensagem As String)
    ' Escreve a mensagem no arquivo de log
    Dim arquivo As Integer

    On Error GoTo TrataErro
        arquivo = FreeFile
        Open CaminhoArquivoLog For Append As #arquivo
        Print #arquivo, mensagem
        Close #arquivo
    Exit Sub
TrataErro:
    Err.Raise ERR_LOG_WRITE_FAILED, "clsLogger", "Erro ao registrar log: " & Err.Description
End Sub
