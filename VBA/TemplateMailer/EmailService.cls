VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EmailService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'   Essa é a classe mais importante.
'
'   Função: Unifica o processo completo sem expor complexidade ao programador.
'
'   Ela é responsável por:
'       Carregar o template via loader
'       Renderizar o template com variáveis
'       Criar o objeto Email
'       Enviar o e-mail via IEmailClient
'
'   Por que isso é bom? Se amanhã você quiser trocar:
'       SMTP por API Rest?
'       HTML por Markdown?
'       OFT por banco de dados?
'       Nada muda no código que usa o EmailService.

Option Explicit

Private client As IEmailClient
Private loader As IEmailTemplateLoader

Public Sub Init(cl As IEmailClient, ld As IEmailTemplateLoader)
    Set client = cl
    Set loader = ld
End Sub

Public Sub SendUsingTemplate( _
    ByVal toAddress As String, _
    ByVal subject As String, _
    ByVal templatePath As String, _
    keysValues As Scripting.Dictionary)
        
    Dim template As IEmailTemplate
    Set template = loader.LoadTemplate(templatePath)
    
    Dim renderedBody As String
    renderedBody = template.Render(keysValues)
    
    Dim mail As New Email
    With mail
        .toAddress = toAddress
        .subject = subject
        .body = renderedBody
    End With
    
    client.SendEmail mail
End Sub

