VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsEmailOutlook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==========================================================================================
' Classe VBA: clsEmailOutlook - Envio de E-mails com Microsoft Outlook
' Versão: 2.0.0
' Autor: Lucas Wyllamys Carmo da Silva
' Criado em: 06/10/2025
'==========================================================================================
'
' 1. DESCRIÇÃO
'   Esta classe permite o envio de e-mails utilizando o Microsoft Outlook, com suporte a:
'       - Corpo do e-mail em HTML (com ou sem template).
'       - Inclusão de anexos.
'       - Cópia (CC) e cópia oculta (BCC).
'       - Substituição de chaves em templates HTML.
'
' 2. PRÉ-REQUISITOS
'   - Ativar a biblioteca "Microsoft Outlook 16.0 Object Library" no Editor VBA: Ferramentas > Referências > Microsoft Outlook 16.0 Object Library
'   - Ativar a biblioteca "Microsoft Scripting Runtime" no Editor VBA: Ferramentas > Referências > Microsoft Scripting Runtime
'   - Baixar e instalar a classe ClsLogger:
'       - Baixe no link: https://github.com/LucasWyllamys/utils/blob/main/VBA/clsLogger.cls
'       - Importe no projeto.
'
' 3. ESTRUTURA INTERNA
'   3.1. Propriedades encapsuladas:
'       - pHTMLTemplate: armazena o conteúdo HTML do template.
'       - pLogger: instância da classe clsLogger.
'       - pOutlookApp, pEmailItem: objetos do Outlook para envio de e-mail.
'
'   3.2. Constantes de erro:
'       - ERR_OUTLOOK_INIT_FAILED: Código de erro para falha na inicialização do outlook.
'       - ERR_TEMPLATE_NOT_FOUND: Código de erro para falha na localização do arquivo template.
'       - ERR_ATTACHMENT_NOT_FOUND: Código de erro para falha na localização do arquivo de anexo.
'       - ERR_TEMPLATE_SUBSTITUTION_FAILED: Código de erro para falha na substituição de chaves.
'       - ERR_EMAIL_SEND_FAILED: Código de erro para falha ao enviar e-mail.
'
'
' 4. MÉTODOS PÚBLICOS
'   4.1. EnviarEmail(...)
'       - Descrição: Envia um e-mail com os parâmetros fornecidos.
'       - Parâmetros:
'           EnviarEmail(remetente As String, _
'                       destinatario As String, _
'                       assunto As String, _
'                       Optional corpoEmailHTML As String = "", _
'                       Optional copia As String = "", _
'                       Optional copiaOculta As String = "", _
'                       Optional caminhoAnexo As String = "")
'       - Comportamento:
'           - Se pHTMLTemplate estiver definido, ele será usado como corpo do e-mail.
'           - Caso contrário, será usado corpoEmailHTML.
'           - A assinatura padrão do Outlook será mantida se corpoEmailHTML for usado.
'           - Anexos são adicionados se os caminhos forem informados.
'           - O e-mail é enviado automaticamente (.Send).
'
'   4.2. ImportarTemplateHTML(caminhoTemplateHTML As String)
'        - Descrição: Importa um arquivo HTML como template.
'        - Parâmetros:
'            - caminhoTemplateHTML: caminho do arquivo .html.
'        - Comportamento:
'            - Lê o conteúdo do arquivo HTML.
'            - Ajusta caminhos relativos para absolutos.
'            - Armazena o conteúdo em pHTMLTemplate.
'
'   4.3. SubstituirChavesTemplateHTML(Dicionario As Scripting.Dictionary)
'        - Descrição: Substitui variáveis no template HTML com base em um dicionário.
'        - Parâmetros:
'            - Dicionario: objeto Dictionary com pares chave-valor.
'        - Comportamento:
'            - Substitui todas as ocorrências das chaves pelos respectivos valores.
'
'   4.4. Resetar()
'        - Descrição: Limpa o estado da classe.
'        - Comportamento:
'            - Remove o item de e-mail.
'            - Limpa o conteúdo do template HTML.
'
' 5. MÉTODOS PRIVADOS
'   5.1. AjustarCaminhosTemplate(caminhoTemplateHTML As String, conteudoHTML As String) As String
'       - Descrição: Corrige caminhos relativos no HTML para caminhos absolutos.
'
'   5.2. LogParametrosEmail(...)
'       - Descrição: Registra os parâmetros utilizados no envio do e-mail.
'
' 6. COMO CRIAR UM TEMPLATE HTML DO E-MAIL DO OUTLOOK
'    6.1. Abra o e-mail no Outlook
'    6.2. Escreva e formate o corpo do e-mail, incluindo variáveis como [NOME], [DATA], etc.
'    6.3. Clique na guia "Arquivo"
'    6.4. Clique em "Salvar como"
'    6.5. Escolha o diretório onde o template será salvo
'    6.6. Modifique o tipo para: HTML (*.htm, *.html)
'    6.7. Clique em "Salvar"
'
' 7. EXEMPLO DE USO
'    Dim email As ClsEmailOutlook ' Declara como objeto do outlook
'    Dim chave_valor As Scripting.Dictionary ' Declara como dicionário
'
'    Set email = New ClsEmailOutlook ' Instancia o objeto
'    Set chave_valor = New Scripting.Dictionary ' Instancia o dicionário
'
'    chave_valor.Add "[CHAVE1]", "Olá," ' Adiciona chave-valor
'    chave_valor.Add "[CHAVE2]", "Mundo!" ' Adiciona chave-valor
'
'    email.ImportarTemplateHTML "C:\Users\U350504\Desktop\Classe de interação com arquivo JSON\Template.htm"
'    email.SubstituirChavesTemplateHTML chave_valor
'    email.EnviarEmail "remetente@exemplo.com", "destinatario@exemplo.com", "Assunto do E-mail", , , , "C:\Users\U350504\Downloads\Documentacao_Classe_Email_Outlook.pdf"
'
' 8. OBSERVAÇÕES
'   - A assinatura padrão do Outlook é preservada se não for usado um template.
'   - O método ImportarTemplateHTML deve ser chamado antes do envio para preparar o corpo do e-mail.
'   - Para anexar múltiplos arquivos, recomenda-se compactá-los em um único .zip.
'   - A classe pode ser facilmente estendida para suportar múltiplos anexos ou outros formatos.

'======================================= Constantes de Erro =========================================

Private Const ERR_OUTLOOK_INIT_FAILED As Long = 1000 ' Erro ao inicializar o Outlook
Private Const ERR_TEMPLATE_NOT_FOUND As Long = 1001 ' Erro de template não encontrado
Private Const ERR_ATTACHMENT_NOT_FOUND As Long = 1002 ' Erro de anexo não encontrado
Private Const ERR_TEMPLATE_SUBSTITUTION_FAILED As Long = 1003 ' Erro de falha de substituição de chaves
Private Const ERR_EMAIL_SEND_FAILED As Long = 1004 ' Erro de envio de e-mail

'====================================== Propriedades Privadas ========================================
 
Private pOutlookApp                 As Outlook.Application          ' Referência do Outlook
Private pEmailItem                  As Outlook.MailItem             ' Referência do E-mail
Private pLogger                     As clsLogger                    ' Para registro em arquivo de log
Private pHTMLTemplate               As String                       ' Armazena o HTML do template de e-mail

'=========================================== Inicialização ===========================================

Private Sub Class_Initialize()
    ' Método chamado automaticamente ao instanciar a classe
    ' Inicializa o Outlook e o logger
    On Error GoTo TratarErro
    
    Set OutlookApp = New Outlook.Application 'Instanciamento do Outlook
    Set Logger = New clsLogger
    
    Logger.LogDebug "========== Início ClsEmailOutlook =========="
    
    Exit Sub
TratarErro:
    If Not Logger Is Nothing Then Logger.LogError "Erro ao inicializar: " & Err.Description
    Err.Raise ERR_OUTLOOK_INIT_FAILED, "ClsEmailOutlook", "Erro ao inicializar: " & Err.Description ' Lançamento de erro
End Sub

Private Sub Class_Terminate()
    ' Método chamado automaticamente ao destruir a instância da classe
    ' Libera os objetos utilizados
    Set OutlookApp = Nothing
    Set EmailItem = Nothing
    
    If Not Logger Is Nothing Then Logger.LogDebug "========== Fim ClsEmailOutlook =========="
End Sub

'==================================== Métodos Assessores Privados ====================================

' Propriedades encapsuladas para acesso aos objetos privados
Private Property Get OutlookApp() As Outlook.Application: Set OutlookApp = pOutlookApp: End Property

Private Property Set OutlookApp(objeto As Outlook.Application): Set pOutlookApp = objeto: End Property

Private Property Get EmailItem() As Outlook.MailItem: Set EmailItem = pEmailItem: End Property

Private Property Set EmailItem(objeto As Outlook.MailItem): Set pEmailItem = objeto: End Property

Private Property Get Logger() As clsLogger: Set Logger = pLogger: End Property

Private Property Set Logger(objeto As clsLogger): Set pLogger = objeto: End Property

Private Property Get HTMLTemplate() As String: HTMLTemplate = pHTMLTemplate: End Property

Private Property Let HTMLTemplate(value As String): pHTMLTemplate = value: End Property

'=========================================== Métodos Públicos ===========================================

Public Sub ImportarTemplateHTML(caminhoTemplateHTML As String)
    ' Importa um template HTML de e-mail
    ' caminhoTemplateHTML: caminho completo do arquivo HTML
    Dim arquivo As Integer
    Dim linha As String
    Dim conteudoHTML As String
    
    On Error GoTo TratarErro

    Logger.LogInfo "Importando template HTML", 1, 3
    
    ' Abre o arquivo para leitura
    arquivo = FreeFile
    Open caminhoTemplateHTML For Input As #arquivo
    ' Lê o conteúdo do arquivo linha por linha
    Do While Not EOF(arquivo)
        Line Input #arquivo, linha
        conteudoHTML = conteudoHTML & linha & vbCrLf
    Loop
    ' Fecha o arquivo
    Close #arquivo

    HTMLTemplate = AjustarCaminhosTemplate(caminhoTemplateHTML, conteudoHTML)
    Logger.LogInfo "Template HTML importado com sucesso", 1, 3

    Exit Sub
TratarErro:
    Logger.LogError "Erro ao importar template HTML: " & Err.Description
    Err.Raise ERR_TEMPLATE_NOT_FOUND, "ClsEmailOutlook", "Erro ao importar template HTML: " & Err.Description
End Sub

Public Sub SubstituirChavesTemplateHTML(dicionario As Scripting.Dictionary)
    ' Substitui chaves no template HTML por valores do dicionário
    ' Dicionario: objeto Scripting.Dictionary com pares chave-valor
    Dim key As Variant
    
    On Error GoTo TratarErro
    
    Logger.LogInfo "Substituindo chaves no template HTML", 2, 3
    
    If Not dicionario Is Nothing Then ' Verifica se o dicionário não está vazio
        For Each key In dicionario.Keys   ' Itera sobre todos as chaves do dicionário
            HTMLTemplate = Replace(HTMLTemplate, key, dicionario(key)) ' Substitui os valores das respectivas chaves
        Next key
        
        Logger.LogInfo "Substituições concluídas com sucesso", 2, 3
    Else
        Logger.LogWarn "Dicionário não definido"
    End If
    
    Exit Sub
TratarErro:
    Logger.LogError "Erro ao substituir chaves no template HTML: " & Err.Description
    Err.Raise ERR_TEMPLATE_SUBSTITUTION_FAILED, "ClsEmailOutlook", "Erro ao substituir chaves no template HTML: " & Err.Description
End Sub

Public Sub EnviarEmail(remetente As String, destinatario As String, assunto As String, _
                        Optional corpoEmailHTML As String = "", _
                        Optional copia As String = "", _
                        Optional copiaOculta As String = "", _
                        Optional caminhoAnexo As String = "")
    ' Envia e-mail com ou sem template HTML e com ou sem anexo
    ' remetente: endereço do remetente
    ' destinatario: endereço do destinatário
    ' assunto: assunto do e-mail
    ' corpoEmailHTML: corpo do e-mail (se não usar template)
    ' copia: destinatários em cópia
    ' copiaOculta: destinatários em cópia oculta
    ' caminhoAnexo: caminho do arquivo de anexo (para anexar múltiplos arquivos, compactá-los em um único .zip)
    Dim corpoEmail As String
    
    On Error GoTo TratarErro
    
    Logger.LogInfo "Preparando envio de e-mail", 3, 3
    LogParametrosEmail remetente, destinatario, copia, copiaOculta, caminhoAnexo

    Set EmailItem = OutlookApp.CreateItem(olMailItem) 'Cria um item de e-mail
    
    With EmailItem
        .Display 'Exibir e-mail
        .SentOnBehalfOfName = remetente
        .To = destinatario
        If copia <> "" Then .CC = copia
        If copiaOculta <> "" Then .BCC = copiaOculta
        .Subject = assunto
        
        If HTMLTemplate = "" Then
            .HTMLBody = corpoEmailHTML & .HTMLBody ' Corpo + Assinatura padrão
        Else
            .HTMLBody = HTMLTemplate
        End If
        
        If caminhoAnexo <> "" Then
            If Dir(caminhoAnexo) <> "" Then
                .Attachments.Add caminhoAnexo
            Else
                Logger.LogError "Arquivo de anexo não encontrado: " & caminhoAnexo
            End If
        Else
            Logger.LogWarn "Arquivo de anexo não definido"
        End If
        
        .Send 'Enviar e-mail
        Logger.LogInfo "E-mail enviado com sucesso", 3, 3
    End With
    
    Exit Sub
TratarErro:
    Logger.LogError "Erro ao enviar e-mail: " & Err.Description
    Err.Raise ERR_EMAIL_SEND_FAILED, "ClsEmailOutlook", "Erro ao enviar e-mail: " & Err.Description
End Sub

Public Sub Resetar()
    ' Reseta o estado da classe
    ' Remove o item de e-mail e limpa o template HTML
    Set EmailItem = Nothing
    HTMLTemplate = ""
    Logger.LogDebug "Estado da classe resetado"
End Sub

'=========================================== Métodos Privados ===========================================

Private Function AjustarCaminhosTemplate(caminhoTemplateHTML As String, conteudoHTML As String) As String
    ' Corrige caminhos relativos no HTML para caminhos absolutos
    ' caminhoTemplateHTML: caminho do arquivo HTML
    ' conteudoHTML: conteúdo lido do arquivo
    Dim nomeTemplate As String
    Dim caminhoPastaTemplate As String
    
    On Error GoTo TratarErro
    nomeTemplate = Replace(Replace(Dir(caminhoTemplateHTML), ".html", "_arquivos"), ".htm", "_arquivos")
    caminhoPastaTemplate = Left(caminhoTemplateHTML, InStrRev(caminhoTemplateHTML, "\") - 1)
    
    AjustarCaminhosTemplate = Replace(conteudoHTML, nomeTemplate & "/", caminhoPastaTemplate & "\" & nomeTemplate & "\")
    Exit Function
TratarErro:
    Logger.LogError "Erro ao corrigir caminhos HTML: " & Err.Description
    AjustarCaminhosTemplate = conteudoHTML ' Retorna o conteúdo original em caso de erro
End Function

Private Sub LogParametrosEmail(remetente As String, destinatario As String, copia As String, copiaOculta As String, caminhoAnexo As String)
    ' Registra os parâmetros utilizados no envio do e-mail
    Logger.LogInfo "Remetente: " & remetente
    Logger.LogInfo "Destinatário: " & destinatario
    Logger.LogInfo "Cópia: " & IIf(copia <> "", copia, "Não definida")
    Logger.LogInfo "Cópia oculta: " & IIf(copiaOculta <> "", copiaOculta, "Não definida")
    Logger.LogInfo "Template HTML: " & IIf(HTMLTemplate <> "", "Sim", "Não")
    Logger.LogInfo "Anexo: " & IIf(caminhoAnexo <> "", "Sim", "Não")
End Sub
