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
'    Esta classe permite o envio de e-mails utilizando o Microsoft Outlook, com suporte a:
'      - Corpo do e-mail em HTML (com ou sem template)
'      - Inclusão de anexos
'      - Cópia (CC) e cópia oculta (BCC)
'      - Substituição de chaves em templates HTML
'
' 2. PRÉ-REQUISITOS
'    - Ativar a biblioteca "Microsoft Outlook 16.0 Object Library" no Editor VBA: Ferramentas > Referências > Microsoft Outlook 16.0 Object Library
'    - Ativar a biblioteca "Microsoft Scripting Runtime" no Editor VBA: Ferramentas > Referências > Microsoft Scripting Runtime
'    - Ter a classe ClsLogger instalada:
'
' 3. ESTRUTURA INTERNA
'    3.1. A classe utiliza um Type privado chamado ClassType para armazenar os dados do e-mail e objetos do Outlook.
'    3.2. Propriedades encapsuladas:
'         - pHTMLTemplate
'         - pOutlookApp, pEmailItem (objetos do Outlook)
'
' 4. MÉTODOS PÚBLICOS
'    4.1. EnviarEmail(...)
'         - Descrição: Envia um e-mail com os parâmetros fornecidos.
'         - Parâmetros:
'             EnviarEmail(remetente As String, _
'                         destinatario As String, _
'                         assunto As String, _
'                         Optional corpoEmailHTML As String, _
'                         Optional copia As String, _
'                         Optional copiaOculta As String, _
'                         Optional caminhoAnexo As String)
'         - Comportamento:
'             - Se pHTMLTemplate estiver definido, ele será usado como corpo do e-mail.
'             - Caso contrário, será usado corpoEmailHTML.
'             - Se houver assinatura padrão do Outlook, ela será mantida ao final do corpo do e-mail.
'             - Anexos são adicionados se os caminhos forem informados.
'             - O e-mail é exibido (.Display) e enviado automaticamente (.Send).
'
'    4.2. ImportarTemplateHTML(caminhoArquivoTemplateHTML As String, Optional Dicionario As Object) As String
'         - Descrição: Importa um arquivo HTML como template e realiza substituições de texto.
'         - Parâmetros:
'             - caminhoArquivoTemplateHTML: caminho do arquivo .html
'             - Dicionario: Dictionary com pares chave -> valor para substituição no conteúdo
'         - Retorno:
'             - Conteúdo HTML final com substituições aplicadas, armazenado em pHTMLTemplate.
'
' 5. MÉTODO PRIVADO
'    5.1. SubstituirConteudoArquivoHTML(conteudoHTML As String, Optional Dicionario As Object) As String
'         - Descrição: Realiza substituições no conteúdo HTML.
'         - Comportamento:
'             - Ajusta caminhos de imagens para a pasta templateEmail_arquivos
'             - Substitui variáveis com base no dicionário fornecido
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
'    Dim email As clsEmailOutlook ' Declara como objeto do outlook
'    Dim variaveis As Scripting.Dictionary ' Declara como dicionário
'
'    Set email = New clsEmailOutlook ' Instancia o objeto
'    Set variaveis = New Scripting.Dictionary ' Instancia o dicionário
'
'    variaveis.Add "[NOME]", "Lucas" ' Adiciona chave-valor
'    variaveis.Add "[IDADE]", "29 anos"
'
'    email.ImportarTemplateHTML "C:\Templates\templateEmail.html", variaveis
'    email.EnviarEmail "lucas@empresa.com", "destinatario@exemplo.com", "Assunto do E-mail", , , , "C:\Anexos\arquivo1.pdf"
'
' 8. OBSERVAÇÕES
'    - A assinatura padrão do Outlook é preservada se não for usado um template.
'    - O método ImportarTemplateHTML é usado separadamente para preparar o corpo do e-mail antes do envio.
'    - Para anexar múltiplos arquivos, recomenda-se compactá-los em um único .zip.
'    - A classe pode ser facilmente estendida para suportar mais anexos ou formatos.

'======================================= Constantes de Erro =========================================

Private Const ERR_OUTLOOK_INIT_FAILED As Long = 1000
Private Const ERR_TEMPLATE_NOT_FOUND As Long = 1001
Private Const ERR_ATTACHMENT_NOT_FOUND As Long = 1002
Private Const ERR_TEMPLATE_SUBSTITUTION_FAILED As Long = 1003
Private Const ERR_EMAIL_SEND_FAILED As Long = 1004

'====================================== Propriedades Privadas ========================================
 
Private pOutlookApp                 As Outlook.Application          'Referência do Outlook
Private pEmailItem                  As Outlook.MailItem             'Referência do E-mail
Private pLogger                     As clsLogger
Private pHTMLTemplate               As String

'=========================================== Inicialização ===========================================

Private Sub Class_Initialize()
    On Error GoTo TratarErro
    
    Set OutlookApp = New Outlook.Application 'Instanciamento do Outlook
    Set Log = New clsLogger
    
    Log.Debug_ "========== Início ClsEmailOutlook =========="
    
    Exit Sub
TratarErro:
    If Not Log Is Nothing Then Log.Error_ "Erro ao inicializar: " & Err.Description
    Err.Raise ERR_OUTLOOK_INIT_FAILED, "ClsEmailOutlook", "Erro ao inicializar: " & Err.Description ' Lançamento de erro
End Sub

Private Sub Class_Terminate()
    Set OutlookApp = Nothing
    Set EmailItem = Nothing
    
    If Not Log Is Nothing Then Log.Debug_ "========== Fim ClsEmailOutlook =========="
End Sub

'==================================== Métodos Assessores Privados ====================================

Private Property Get OutlookApp() As Outlook.Application: Set OutlookApp = pOutlookApp: End Property

Private Property Set OutlookApp(objeto As Outlook.Application): Set pOutlookApp = objeto: End Property

Private Property Get EmailItem() As Outlook.MailItem: Set EmailItem = pEmailItem: End Property

Private Property Set EmailItem(objeto As Outlook.MailItem): Set pEmailItem = objeto: End Property

Private Property Get Log() As clsLogger: Set Log = pLogger: End Property

Private Property Set Log(objeto As clsLogger): Set pLogger = objeto: End Property

Private Property Get HTMLTemplate() As String: HTMLTemplate = pHTMLTemplate: End Property

Private Property Let HTMLTemplate(value As String): pHTMLTemplate = value: End Property

'=========================================== Métodos Públicos ===========================================

Public Sub ImportarTemplateHTML(caminhoTemplateHTML As String)
    ' Importa o template HTML realizando tratativas necessárias
    Dim arquivo As Integer
    Dim linha As String
    Dim conteudoHTML As String
    
    On Error GoTo TratarErro
    
    Log.Info_ "Importando template HTML", 1, 3
    
    If Dir(caminhoTemplateHTML) = "" Then
        Log.Error_ "Arquivo de template não encontrado: " & caminhoTemplateHTML
        Err.Raise Number:=1001, Description:="Arquivo de template não encontrado: " & caminhoTemplateHTML ' Força exceção
    End If
    
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
    Log.Info_ "Template HTML importado com sucesso", 1, 3

    Exit Sub
TratarErro:
    Log.Warn_ "Erro ao importar template HTML: " & Err.Description
    Err.Raise ERR_TEMPLATE_NOT_FOUND, "ClsEmailOutlook", "Erro ao importar template HTML: " & Err.Description
End Sub

Public Sub SubstituirChavesTemplateHTML(Dicionario As Scripting.Dictionary)
    ' Substitui as chaves no corpo do template pelos seus respectivos valores definidos no dicionário
    Dim key As Variant
    
    On Error GoTo TratarErro
    
    Log.Info_ "Substituindo chaves no template HTML", 2, 3
    
    If Not Dicionario Is Nothing Then ' Verifica se o dicionário não está vazio
        For Each key In Dicionario.Keys   ' Itera sobre todos as chaves do dicionário
            HTMLTemplate = Replace(HTMLTemplate, key, Dicionario(key)) ' Substitui os valores das respectivas chaves
        Next key
        
        Log.Info_ "Substituições concluídas com sucesso", 2, 3
    Else
        Log.Warn_ "Dicionário não definido"
    End If
    
    Exit Sub
TratarErro:
    Log.Warn_ "Erro ao substituir chaves no template HTML: " & Err.Description
    Err.Raise ERR_TEMPLATE_SUBSTITUTION_FAILED, "ClsEmailOutlook", "Erro ao substituir chaves no template HTML: " & Err.Description
End Sub

Public Sub EnviarEmail(remetente As String, destinatario As String, assunto As String, _
                        Optional corpoEmailHTML As String = "", _
                        Optional copia As String = "", _
                        Optional copiaOculta As String = "", _
                        Optional caminhoAnexo As String = "")
    ' Envia e-mail com ou sem: anexo, template
    Dim corpoEmail As String
    
    On Error GoTo TratarErro
    
    Log.Info_ "Preparando envio de e-mail", 3, 3
    LogParametrosEmail remetente, destinatario, copia, copiaOculta, caminhoAnexo

    Set EmailItem = OutlookApp.CreateItem(olMailItem) 'Cria um item de e-mail
    
    With EmailItem
        .SentOnBehalfOfName = remetente
        .To = destinatario
        If copia <> "" Then .CC = copia
        If copiaOculta <> "" Then .BCC = copiaOculta
        .Subject = assunto
        
        If HTMLTemplate = "" Then
            .Display 'Exibir e-mail
            .HTMLBody = corpoEmailHTML & .HTMLBody ' Corpo + Assinatura padrão
        Else
            .HTMLBody = HTMLTemplate
        End If
        
        If caminhoAnexo <> "" Then
            If Dir(caminhoAnexo) <> "" Then
                .Attachments.Add caminhoAnexo
            Else
                Log.Error_ "Arquivo de anexo não encontrado: " & caminhoAnexo
                Err.Raise Number:=1001, Description:="Arquivo de anexo não encontrado: " & caminhoAnexo ' Força exceção
            End If
        Else
            Log.Warn_ "Arquivo de anexo não definido"
        End If
        
        .Send 'Enviar e-mail
        Log.Info_ "E-mail enviado com sucesso", 3, 3
    End With
    
    Exit Sub
TratarErro:
    Log.Error_ "Erro ao enviar e-mail: " & Err.Description
    Err.Raise ERR_EMAIL_SEND_FAILED, "ClsEmailOutlook", "Erro ao enviar e-mail: " & Err.Description
End Sub

Public Sub Resetar()
    ' Limpa o objeto de e-mail e o template HTML
    Set EmailItem = Nothing
    HTMLTemplate = ""
    Log.Debug_ "Estado da classe resetado"
End Sub

'=========================================== Métodos Privados ===========================================

Private Function AjustarCaminhosTemplate(caminhoTemplateHTML As String, conteudoHTML As String) As String
    ' Substitui as referências de caminho relativo pelo caminho absoluto
    Dim nomeTemplate As String
    Dim caminhoPastaTemplate As String
    
    On Error GoTo TratarErro
    nomeTemplate = Replace(Replace(Dir(caminhoTemplateHTML), ".html", "_arquivos"), ".htm", "_arquivos")
    caminhoPastaTemplate = Left(caminhoTemplateHTML, InStrRev(caminhoTemplateHTML, "\") - 1)
    
    AjustarCaminhosTemplate = Replace(conteudoHTML, nomeTemplate & "/", caminhoPastaTemplate & "\" & nomeTemplate & "\")
    Exit Function
TratarErro:
    Log.Warn_ "Erro ao corrigir caminhos HTML: " & Err.Description
    AjustarCaminhosTemplate = conteudoHTML ' Retorna o conteúdo original em caso de erro
End Function

Private Sub LogParametrosEmail(remetente As String, destinatario As String, copia As String, copiaOculta As String, caminhoAnexo As String)
    Log.Info_ "Remetente: " & remetente
    Log.Info_ "Destinatário: " & destinatario
    Log.Info_ "Cópia: " & IIf(copia <> "", copia, "Não definida")
    Log.Info_ "Cópia oculta: " & IIf(copiaOculta <> "", copiaOculta, "Não definida")
    Log.Info_ "Template HTML: " & IIf(HTMLTemplate <> "", "Sim", "Não")
    Log.Info_ "Anexo: " & IIf(caminhoAnexo <> "", "Sim", "Não")
End Sub
