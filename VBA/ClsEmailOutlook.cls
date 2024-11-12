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

' Esta classe tem por objetivo enviar e-mails pelo servidor do Outlook utilizando ou não um template HTML.
' ATENÇÃO: Necessário habilitar a biblioteca "Microsoft Outlook 16.0 Object Library"

Private Type ClassType
    OutApp          As New Outlook.Application      'Referência do Outlook
    OutEmail        As Outlook.MailItem             'Referência do E-mail
        
    remetente       As String
    destinatario    As String
    copia           As String
    CopiaOculta     As String
    assunto         As String
    CorpoEmail      As String                       'Deve ser inserido em HTML
    anexo1          As String
    anexo2          As String
    
    TemplateHTML    As String
End Type

Private Self As ClassType

'Métodos acessores (propriedades):--------------------------------------------------------------------------------------------------------------------

Private Property Get OutApp() As Outlook.Application
    Set OutApp = Self.OutApp
End Property

Private Property Set OutApp(Objeto As Outlook.Application)
    Set Self.OutApp = Objeto
End Property

Private Property Get OutEmail() As Outlook.MailItem
    Set OutEmail = Self.OutEmail
End Property

Private Property Set OutEmail(Objeto As Outlook.MailItem)
    Set Self.OutEmail = Objeto
End Property

Private Property Get remetente() As String
    remetente = Self.remetente
End Property

Private Property Let remetente(Value As String)
    Self.remetente = Value
End Property

Private Property Get destinatario() As String
    destinatario = Self.destinatario
End Property

Private Property Let destinatario(Value As String)
    Self.destinatario = Value
End Property

Private Property Get copia() As String
    copia = Self.copia
End Property

Private Property Let copia(Value As String)
    Self.copia = Value
End Property

Private Property Get CopiaOculta() As String
    CopiaOculta = Self.CopiaOculta
End Property

Private Property Let CopiaOculta(Value As String)
    Self.CopiaOculta = Value
End Property

Private Property Get assunto() As String
    assunto = Self.assunto
End Property

Private Property Let assunto(Value As String)
    Self.assunto = Value
End Property

Private Property Get CorpoEmail() As String
    CorpoEmail = Self.CorpoEmail
End Property

Private Property Let CorpoEmail(Value As String)
    Self.CorpoEmail = Value
End Property

Private Property Get anexo1() As String
    anexo1 = Self.anexo1
End Property

Private Property Let anexo1(CaminhoArquivo As String)
    Self.anexo1 = CaminhoArquivo
End Property

Private Property Get anexo2() As String
    anexo2 = Self.anexo2
End Property

Private Property Let anexo2(CaminhoArquivo As String)
    Self.anexo2 = CaminhoArquivo
End Property

Private Property Get TemplateHTML() As String
    TemplateHTML = Self.TemplateHTML
End Property

Private Property Let TemplateHTML(ConteudoHTML As String)
    Self.TemplateHTML = ConteudoHTML
End Property

'Métodos:--------------------------------------------------------------------------------------------------------------------------------------

Public Sub EnviarEmail(ParamDe As String, ParamPara As String, ParamAssunto As String, Optional ParamCorpoEmailHTML As String, _
Optional ParamCopia As String, Optional ParamCopiaOculta As String, Optional ParamCaminhoAnexo1 As String, Optional ParamCaminhoAnexo2 As String)
   
    Dim Assinatura As String
    
    remetente = ParamDe
    destinatario = ParamPara
    assunto = ParamAssunto
    If TemplateHTML <> "" Then
        CorpoEmail = TemplateHTML
    Else
        CorpoEmail = ParamCorpoEmailHTML
    End If
    copia = ParamCopia
    CopiaOculta = ParamCopiaOculta
    anexo1 = ParamCaminhoAnexo1
    anexo2 = ParamCaminhoAnexo2
    
    Set OutEmail = OutApp.CreateItem(0)                         'Cria um item de e-mail (objeto de e-mail)
    
    With OutEmail
        .display                                                'Exibir e-mail
        .SentOnBehalfOfName = remetente                         'Remetente
        .To = destinatario                                      'Destinatário
        If copia <> "" Then .CC = copia                         'Cópia
        If CopiaOculta <> "" Then .bcc = CopiaOculta            'Cópia oculta
        .Subject = assunto                                      'Assunto
        If TemplateHTML = "" Then
            If .HTMLBody <> "" Then
                Assinatura = .HTMLBody                          'Assinatura Padrão
                .HTMLBody = CorpoEmail & Assinatura             'Corpo do e-mail com assinatura
            Else
                .HTMLBody = CorpoEmail                          'Corpo do e-mail
            End If
        Else
            .HTMLBody = CorpoEmail                              'Corpo do e-mail
        End If
        If anexo1 <> "" Then .Attachments.Add anexo1            'Anexar arquivo
        If anexo2 <> "" Then .Attachments.Add anexo2            'Anexar arquivo
        .Send                                                   'Enviar e-mail
    End With
End Sub

Public Function ImportarTemplateHTML(CaminhoArquivoTemplateHTML As String, Optional ListaFindReplace As Object) As String
    Dim Arquivo As Integer, ConteudoHTML As String, Linha As String
    
    ' Abre o arquivo para leitura
    Arquivo = FreeFile
    Open CaminhoArquivoTemplateHTML For Input As #Arquivo
    ' Lê o conteúdo do arquivo linha por linha
    Do While Not EOF(Arquivo)
        Line Input #Arquivo, Linha
        ConteudoHTML = ConteudoHTML & Linha & vbCrLf
    Loop
    ' Fecha o arquivo
    Close #Arquivo
    
    'Armazena conteúdo HTML na propriedade TemplateHTML substituindo os termos informados por meio da função SubstituirConteudoArquivoHTML
    TemplateHTML = SubstituirConteudoArquivoHTML(ConteudoHTML, ListaFindReplace)
End Function

Private Function SubstituirConteudoArquivoHTML(ConteudoHTML As String, Optional ListaFindReplace As Object) As String
    Dim Key As Variant
    'default: Altera o caminho dos elementos para a pasta templateEmail_arquivos
    ConteudoHTML = Replace(ConteudoHTML, "templateEmail_arquivos/", ThisWorkbook.Path & "\templateEmail_arquivos\")
    
    If Not ListaFindReplace Is Nothing Then
        For Each Key In ListaFindReplace.Keys   'Acessa todos os itens do dicionário
            ConteudoHTML = Replace(ConteudoHTML, Key, ListaFindReplace(Key))
        Next Key
    End If

    SubstituirConteudoArquivoHTML = ConteudoHTML
End Function
