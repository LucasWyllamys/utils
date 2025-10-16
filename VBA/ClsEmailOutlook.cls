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
' Versão: 1.2.0
' Autor: Lucas Wyllamys Carmo da Silva
' Criado em: 06/10/2025
'==========================================================================================
'
' 1. DESCRIÇÃO
'    Esta classe permite o envio de e-mails utilizando o Microsoft Outlook, com suporte a:
'      - Corpo do e-mail em HTML (com ou sem template)
'      - Inclusão de anexos
'      - Cópia (CC) e cópia oculta (BCC)
'      - Substituição de variáveis em templates HTML
'
' 2. PRÉ-REQUISITOS
'    - Ativar a biblioteca "Microsoft Outlook 16.0 Object Library" no Editor VBA:
'        Ferramentas > Referências > Microsoft Outlook 16.0 Object Library
'    - Ativar a biblioteca "Microsoft Scripting Runtime" no Editor VBA:
'        Ferramentas > Referências > Microsoft Scripting Runtime
'
' 3. ESTRUTURA INTERNA
'    3.1. A classe utiliza um Type privado chamado ClassType para armazenar os dados do e-mail e objetos do Outlook.
'    3.2. Propriedades encapsuladas:
'         - pTemplateHTML
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
'             - Se pTemplateHTML estiver definido, ele será usado como corpo do e-mail.
'             - Caso contrário, será usado corpoEmailHTML.
'             - Se houver assinatura padrão do Outlook, ela será mantida ao final do corpo do e-mail.
'             - Anexos são adicionados se os caminhos forem informados.
'             - O e-mail é exibido (.Display) e enviado automaticamente (.Send).
'
'    4.2. ImportarTemplateHTML(caminhoArquivoTemplateHTML As String, Optional listaFindReplace As Object) As String
'         - Descrição: Importa um arquivo HTML como template e realiza substituições de texto.
'         - Parâmetros:
'             - caminhoArquivoTemplateHTML: caminho do arquivo .html
'             - listaFindReplace: Dictionary com pares chave -> valor para substituição no conteúdo
'         - Retorno:
'             - Conteúdo HTML final com substituições aplicadas, armazenado em pTemplateHTML.
'
' 5. MÉTODO PRIVADO
'    5.1. SubstituirConteudoArquivoHTML(conteudoHTML As String, Optional listaFindReplace As Object) As String
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
'    Call email.ImportarTemplateHTML("C:\Templates\templateEmail.html", variaveis)
'    Call email.EnviarEmail("lucas@empresa.com", "destinatario@exemplo.com", "Assunto do E-mail", , , , "C:\Anexos\arquivo1.pdf")
'
' 8. OBSERVAÇÕES
'    - A assinatura padrão do Outlook é preservada se não for usado um template.
'    - O método ImportarTemplateHTML é usado separadamente para preparar o corpo do e-mail antes do envio.
'    - Para anexar múltiplos arquivos, recomenda-se compactá-los em um único .zip.
'    - A classe pode ser facilmente estendida para suportar mais anexos ou formatos.

'=========================================== Propriedades Privadas ===========================================

Private Type ClassType
    pOutlookApp                 As Outlook.Application          'Referência do Outlook
    pEmailItem                  As Outlook.MailItem             'Referência do E-mail
    pLogger                     As clsLogger
    
    pTemplateHTML               As String
End Type

Private self As ClassType

'=========================================== Métodos Assessores ===========================================

Private Property Get OutlookApp() As Outlook.Application
    Set OutlookApp = self.pOutlookApp
End Property

Private Property Set OutlookApp(objeto As Outlook.Application)
    Set self.pOutlookApp = objeto
End Property

Private Property Get EmailItem() As Outlook.MailItem
    Set EmailItem = self.pEmailItem
End Property

Private Property Set EmailItem(objeto As Outlook.MailItem)
    Set self.pEmailItem = objeto
End Property

Private Property Get Logger() As clsLogger
    Set Logger = self.pLogger
End Property

Private Property Set Logger(objeto As clsLogger)
    Set self.pLogger = objeto
End Property

Private Property Get TemplateHTML() As String
    TemplateHTML = self.pTemplateHTML
End Property

Private Property Let TemplateHTML(value As String)
    self.pTemplateHTML = value
End Property

'=========================================== Inicialização ===========================================

Private Sub Class_Initialize()
    On Error GoTo TratarErro
    
    Set OutlookApp = New Outlook.Application 'Instanciamento do Outlook
    Set Logger = New clsLogger
    
    Logger.Log "============================ Início ClsEmailOutlook ============================ ", "DEBUG"
    
    Exit Sub
TratarErro:
    Logger.Log "Erro ao iniciar Outlook: " & Err.Description, "ERROR"
    MsgBox "Erro ao iniciar Outlook: " & Err.Description, vbCritical
End Sub

Private Sub Class_Terminate()
    Set OutlookApp = Nothing
    Set EmailItem = Nothing
    
    Logger.Log "============================== Fim ClsEmailOutlook =============================", "DEBUG"
End Sub

'=========================================== Métodos Públicos ===========================================

Public Sub EnviarEmail(remetente As String, destinatario As String, assunto As String, _
                        Optional corpoEmailHTML As String = "", _
                        Optional copia As String = "", _
                        Optional copiaOculta As String = "", _
                        Optional caminhoAnexo As String = "")
    On Error GoTo TratarErro
    
    Dim corpoEmail As String
    
    Logger.Log "Enviando e-mail", "INFO", 3, 3
    Logger.Log "Remetente: " & remetente, "INFO", 3, 3
    Logger.Log "Destinatário: " & destinatario, "INFO", 3, 3
    If copia <> "" Then
        Logger.Log "Cópia: " & copia, "INFO", 3, 3
    Else
        Logger.Log "Cópia: " & False, "INFO", 3, 3
    End If
    If copiaOculta <> "" Then
        Logger.Log "Cópia oculta: " & copiaOculta, "INFO", 3, 3
    Else
        Logger.Log "Cópia oculta: " & False, "INFO", 3, 3
    End If
    If TemplateHTML <> "" Then
        Logger.Log "Template HTML: " & True, "INFO", 3, 3
    Else
        Logger.Log "Template HTML: " & False, "WARN", 3, 3
    End If
    If caminhoAnexo <> "" Then
        Logger.Log "Anexo: " & True, "INFO", 3, 3
    Else
        Logger.Log "Anexo: " & False, "WARN", 3, 3
    End If
    
    Set EmailItem = OutlookApp.CreateItem(olMailItem) 'Cria um item de e-mail
    
    With EmailItem
        .Display 'Exibir e-mail
        .SentOnBehalfOfName = remetente
        .To = destinatario
        If copia <> "" Then .CC = copia
        If copiaOculta <> "" Then .BCC = copiaOculta
        .Subject = assunto
        
        If TemplateHTML = "" Then
            If .HTMLBody <> "" Then
                .HTMLBody = corpoEmailHTML & .HTMLBody ' Corpo + Assinatura padrão
            Else
                .HTMLBody = corpoEmailHTML
            End If
        Else
            .HTMLBody = TemplateHTML
        End If
        
        If caminhoAnexo <> "" Then .Attachments.Add caminhoAnexo
        .Send 'Enviar e-mail
    End With
    
    Logger.Log "E-mail enviado com sucesso", "INFO", 3, 3
    
    Exit Sub
TratarErro:
    Logger.Log "Erro ao enviar e-mail: " & Err.Description, "ERROR"
    MsgBox "Erro ao enviar e-mail: " & Err.Description, vbCritical
End Sub

Public Sub ImportarTemplateHTML(caminhoTemplateHTML As String, Optional ListaFindReplace As Scripting.Dictionary)
    On Error GoTo TratarErro
    
    Dim arquivo As Integer
    Dim linha As String
    Dim conteudoHTML As String
    
    Logger.Log "Importando template HTML", "INFO", 1, 3
    
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
    
    Logger.Log "Importação concluída com sucesso", "INFO", 1, 3
    
    If Not ListaFindReplace Is Nothing And TypeName(ListaFindReplace) = "Dictionary" Then ' Verifica se o objeto não está vazio e é um dicionário
        TemplateHTML = SubstituirConteudoArquivoHTML(caminhoTemplateHTML, conteudoHTML, ListaFindReplace)
    Else
        TemplateHTML = conteudoHTML
    End If
    
    Exit Sub
TratarErro:
    Logger.Log "Erro ao importar template HTML: " & Err.Description, "ERROR"
    MsgBox "Erro ao importar template HTML: " & Err.Description, vbCritical
End Sub

'=========================================== Métodos Privados ===========================================

Private Function SubstituirConteudoArquivoHTML(caminhoTemplateHTML As String, conteudoHTML As String, ListaFindReplace As Scripting.Dictionary) As String
    Dim key As Variant
    Dim nomeTemplate As String
    
    Logger.Log "Substituindo chaves no conteúdo do template HTML", "INFO", 2, 3
    
    nomeTemplate = Dir(caminhoTemplateHTML)
    nomeTemplate = Replace(Replace(nomeTemplate, ".html", "_arquivos"), ".htm", "_arquivos")
    ' default: Altera o caminho dos elementos para a pasta templateEmail_arquivos
    conteudoHTML = Replace(conteudoHTML, nomeTemplate & "/", ThisWorkbook.Path & "\" & nomeTemplate & "\")
    
    For Each key In ListaFindReplace.Keys   ' Itera sobre todos as chaves do dicionário
        conteudoHTML = Replace(conteudoHTML, key, ListaFindReplace(key)) ' Substitui os valores das respectivas chaves
    Next key
    
    Logger.Log "Substituição concluída com sucesso", "INFO", 2, 3

    SubstituirConteudoArquivoHTML = conteudoHTML
End Function
