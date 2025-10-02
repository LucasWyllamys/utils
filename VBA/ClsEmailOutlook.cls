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
'         - remetente, destinatario, copia, copiaOculta
'         - assunto, corpoEmail, templateHTML
'         - anexo1, anexo2
'         - outApp, outEmail (objetos do Outlook)
'
' 4. MÉTODOS PÚBLICOS
'
'    4.1. EnviarEmail(...)
'         - Descrição: Envia um e-mail com os parâmetros fornecidos.
'         - Parâmetros:
'             EnviarEmail(paramDe As String, _
'                         paramPara As String, _
'                         paramAssunto As String, _
'                         Optional paramCorpoEmailHTML As String, _
'                         Optional paramCopia As String, _
'                         Optional paramCopiaOculta As String, _
'                         Optional paramCaminhoAnexo1 As String, _
'                         Optional paramCaminhoAnexo2 As String)
'         - Comportamento:
'             - Se templateHTML estiver definido, ele será usado como corpo do e-mail.
'             - Caso contrário, será usado paramCorpoEmailHTML.
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
'             - Conteúdo HTML final com substituições aplicadas, armazenado em templateHTML.
'
' 5. MÉTODO PRIVADO
'
'    5.1. SubstituirConteudoArquivoHTML(conteudoHTML As String, Optional listaFindReplace As Object) As String
'         - Descrição: Realiza substituições no conteúdo HTML.
'         - Comportamento:
'             - Ajusta caminhos de imagens para a pasta templateEmail_arquivos
'             - Substitui variáveis com base no dicionário fornecido
'
' 6. COMO CRIAR UM TEMPLATE HTML DO E-MAIL DO OUTLOOK
'
'    6.1. Abra o e-mail no Outlook
'    6.2. Escreva e formate o corpo do e-mail, incluindo variáveis como [NOME], [DATA], etc.
'    6.3. Clique na guia "Arquivo"
'    6.4. Clique em "Salvar como"
'    6.5. Escolha o diretório onde o template será salvo
'    6.6. Modifique o tipo para: HTML (*.htm, *.html)
'    6.7. Clique em "Salvar"
'
' 7. EXEMPLO DE USO
'
'    Dim email As New clsEmailOutlook           ' Instancia o objeto
'    Dim variaveis As New Scripting.Dictionary  ' Instancia o dicionário
'
'    variaveis.Add "[NOME]", "Lucas"            ' Adiciona chave-valor
'    variaveis.Add "[IDADE]", "29 anos"         ' Adiciona chave-valor
'
'    Call email.ImportarTemplateHTML("C:\Templates\templateEmail.html", variaveis)
'    Call email.EnviarEmail("lucas@empresa.com", "destinatario@exemplo.com", "Assunto do E-mail", , , , "C:\Anexos\arquivo1.pdf")
'
' 8. OBSERVAÇÕES
'
'    - A assinatura padrão do Outlook é preservada se não for usado um template.
'    - O método ImportarTemplateHTML é usado separadamente para preparar o corpo do e-mail antes do envio.
'    - Para anexar múltiplos arquivos, recomenda-se compactá-los em um único .zip.
'    - A classe pode ser facilmente estendida para suportar mais anexos ou formatos.

Private Type ClassType
    outApp          As New Outlook.Application      'Referência do Outlook
    outEmail        As Outlook.MailItem             'Referência do E-mail
        
    remetente       As String
    destinatario    As String
    copia           As String
    copiaOculta     As String
    assunto         As String
    corpoEmail      As String                       'Deve ser inserido em HTML
    anexo1          As String
    anexo2          As String
    
    templateHTML    As String
End Type

Private self As ClassType

'Métodos acessores (propriedades):--------------------------------------------------------------------------------------------------------------------

Private Property Get outApp() As Outlook.Application
    Set outApp = self.outApp
End Property

Private Property Set outApp(Objeto As Outlook.Application)
    Set self.outApp = Objeto
End Property

Private Property Get outEmail() As Outlook.MailItem
    Set outEmail = self.outEmail
End Property

Private Property Set outEmail(Objeto As Outlook.MailItem)
    Set self.outEmail = Objeto
End Property

Private Property Get remetente() As String
    remetente = self.remetente
End Property

Private Property Let remetente(Value As String)
    self.remetente = Value
End Property

Private Property Get destinatario() As String
    destinatario = self.destinatario
End Property

Private Property Let destinatario(Value As String)
    self.destinatario = Value
End Property

Private Property Get copia() As String
    copia = self.copia
End Property

Private Property Let copia(Value As String)
    self.copia = Value
End Property

Private Property Get copiaOculta() As String
    copiaOculta = self.copiaOculta
End Property

Private Property Let copiaOculta(Value As String)
    self.copiaOculta = Value
End Property

Private Property Get assunto() As String
    assunto = self.assunto
End Property

Private Property Let assunto(Value As String)
    self.assunto = Value
End Property

Private Property Get corpoEmail() As String
    corpoEmail = self.corpoEmail
End Property

Private Property Let corpoEmail(Value As String)
    self.corpoEmail = Value
End Property

Private Property Get anexo1() As String
    anexo1 = self.anexo1
End Property

Private Property Let anexo1(caminhoArquivo As String)
    self.anexo1 = caminhoArquivo
End Property

Private Property Get anexo2() As String
    anexo2 = self.anexo2
End Property

Private Property Let anexo2(caminhoArquivo As String)
    self.anexo2 = caminhoArquivo
End Property

Private Property Get templateHTML() As String
    templateHTML = self.templateHTML
End Property

Private Property Let templateHTML(conteudoHTML As String)
    self.templateHTML = conteudoHTML
End Property

'Métodos:--------------------------------------------------------------------------------------------------------------------------------------

Public Sub EnviarEmail(paramDe As String, paramPara As String, paramAssunto As String, Optional paramCorpoEmailHTML As String, _
Optional paramCopia As String, Optional paramCopiaOculta As String, Optional paramCaminhoAnexo1 As String, Optional paramCaminhoAnexo2 As String)
   
    Dim Assinatura As String
    
    remetente = paramDe
    destinatario = paramPara
    assunto = paramAssunto
    If templateHTML <> "" Then
        corpoEmail = templateHTML
    Else
        corpoEmail = paramCorpoEmailHTML
    End If
    copia = paramCopia
    copiaOculta = paramCopiaOculta
    anexo1 = paramCaminhoAnexo1
    anexo2 = paramCaminhoAnexo2
    
    Set outEmail = outApp.CreateItem(0)                         'Cria um item de e-mail (objeto de e-mail)
    
    With outEmail
        .display                                                'Exibir e-mail
        .SentOnBehalfOfName = remetente                         'Remetente
        .To = destinatario                                      'Destinatário
        If copia <> "" Then .CC = copia                         'Cópia
        If copiaOculta <> "" Then .bcc = copiaOculta            'Cópia oculta
        .Subject = assunto                                      'Assunto
        If templateHTML = "" Then
            If .HTMLBody <> "" Then
                Assinatura = .HTMLBody                          'Assinatura Padrão
                .HTMLBody = corpoEmail & Assinatura             'Corpo do e-mail com assinatura
            Else
                .HTMLBody = corpoEmail                          'Corpo do e-mail
            End If
        Else
            .HTMLBody = corpoEmail                              'Corpo do e-mail
        End If
        If anexo1 <> "" Then .Attachments.Add anexo1            'Anexar arquivo
        If anexo2 <> "" Then .Attachments.Add anexo2            'Anexar arquivo
        .Send                                                   'Enviar e-mail
    End With
End Sub

Public Function ImportarTemplateHTML(caminhoArquivoTemplateHTML As String, Optional listaFindReplace As Object) As String
    Dim arquivo As Integer, conteudoHTML As String, linha As String
    
    ' Abre o arquivo para leitura
    arquivo = FreeFile
    Open caminhoArquivoTemplateHTML For Input As #arquivo
    ' Lê o conteúdo do arquivo linha por linha
    Do While Not EOF(arquivo)
        Line Input #arquivo, linha
        conteudoHTML = conteudoHTML & linha & vbCrLf
    Loop
    ' Fecha o arquivo
    Close #arquivo
    
    'Armazena conteúdo HTML na propriedade TemplateHTML substituindo os termos informados por meio da função SubstituirConteudoArquivoHTML
    templateHTML = SubstituirConteudoArquivoHTML(caminhoArquivoTemplateHTML, conteudoHTML, listaFindReplace)
End Function

Private Function SubstituirConteudoArquivoHTML(caminhoArquivoTemplateHTML As String, conteudoHTML As String, Optional listaFindReplace As Object) As String
    Dim key As Variant, nomeTemplate As String
    
    nomeTemplate = Dir(caminhoArquivoTemplateHTML)
    nomeTemplate = Replace(Replace(nomeTemplate, ".html", "_arquivos"), ".htm", "_arquivos")
    'default: Altera o caminho dos elementos para a pasta templateEmail_arquivos
    conteudoHTML = Replace(conteudoHTML, "templateEmail_arquivos/", ThisWorkbook.Path & "\templateEmail_arquivos\")
    
    If Not listaFindReplace Is Nothing Then
        For Each key In listaFindReplace.Keys   'Acessa todos os itens do dicionário
            conteudoHTML = Replace(conteudoHTML, key, listaFindReplace(key))
        Next key
    End If

    SubstituirConteudoArquivoHTML = conteudoHTML
End Function
