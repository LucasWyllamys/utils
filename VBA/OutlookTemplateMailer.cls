VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OutlookTemplateMailer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================================================================================
' Classe VBA: OutlookTemplateMailer – Envio de E-mails via Microsoft Outlook
'
' Versão: 6.0.0
' Autor: Lucas Wyllamys Carmo da Silva
' Criado em: 22/10/2025
' Atualizado em: 26/02/2026
'
' Descrição Geral:
'   Esta classe encapsula toda a lógica necessária para criação, formatação e envio de e-mails
'   utilizando o Microsoft Outlook, permitindo automação avançada com suporte a:
'       • Corpo de e-mail em HTML (definido manualmente ou a partir de um template)
'       • Templates HTML locais (*.htm / *.html)
'       • Templates Outlook (*.oft)
'       • Substituição dinâmica de chaves no corpo do e-mail (ex.: [NOME], [DATA], etc.)
'       • Substituição dinâmica de chaves no assunto
'       • Inclusão de anexos (um ou vários)
'       • Envio usando remetente delegado (propriedade SentOnBehalfOfName)
'       • Destinatários em cópia (CC) e cópia oculta (BCC)
'       • Exibição do e-mail antes do envio (modo de revisão/edição pelo usuário)
'
'
' Pré-requisitos:
'   Referências obrigatórias no VBA:
'       • Microsoft Outlook 16.0 Object Library
'           VBA > Ferramentas > Referências > Microsoft Outlook 16.0 Object Library
'       • Microsoft Scripting Runtime
'           VBA > Ferramentas > Referências > Microsoft Scripting Runtime
'
'   Módulos auxiliares necessários:
'       • FileManipulation.bas
'           https://github.com/LucasWyllamys/utils/blob/main/VBA/FileManipulation.bas
'
'
' Como criar um template HTML:
'    1. Abra um novo e-mail no Outlook.
'    2. Escreva e formate o conteúdo normalmente, inserindo chaves como [NOME], [DATA], etc.
'    3. Acesse Arquivo > Salvar como.
'    4. Escolha o diretório desejado.
'    5. Defina o tipo: HTML (*.htm, *.html)
'    6. Clique em Salvar.
'
' Como criar um template OFT:
'    1. Abra um novo e-mail no Outlook.
'    2. Digite o conteúdo desejado, incluindo as chaves.
'    3. Acesse Arquivo > Salvar como.
'    4. Escolha o tipo: Modelo do Outlook (*.oft)
'
'
' Exemplo de uso:
'    Dim email As OutlookTemplateMailer
'    Dim chaveValor As Scripting.Dictionary
'    Dim caminhosAnexos As Collection
'
'    Set email = New OutlookTemplateMailer
'    Set chaveValor = New Scripting.Dictionary
'    Set caminhosAnexos = New Collection
'
'    email.ImportTemplate "C:\templates\Aviso.oft"
'
'    chaveValor.Add "[CHAVE1]", "Olá,"
'    chaveValor.Add "[CHAVE2]", "Mundo!"
'
'    caminhosAnexos.Add "C:\anexo1.pdf"
'    caminhosAnexos.Add "C:\anexo2.pdf"
'
'    email.SendEmail _
'        sender:="remetente@exemplo.com", _
'        recipient:="destinatario@exemplo.com", _
'        subject:="Assunto do E-mail", _
'        keysValues:=chaveValor, _
'        filePaths:=caminhosAnexos, _
'        showBeforeSend:=False
'
'
' Observações importantes:
'   • Quando nenhum template é utilizado, o Outlook preserva automaticamente a assinatura padrão.
'   • O método ImportTemplate deve ser chamado antes do envio para carregar e preparar o corpo.
'   • A substituição de chaves funciona para:
'       - Templates HTML
'       - Templates OFT
'       - Corpo HTML fornecido manualmente
'       - Assunto do e-mail
'   • Para anexar múltiplos arquivos, utilize a Collection ou compacte em .zip.
'   • Templates HTML exportados pelo Outlook criam uma pasta “_arquivos”; a classe converte
'     automaticamente caminhos relativos para absolutos (imagens, CSS, etc.).
'
'===========================================================================================================

Option Explicit

'==================== Constantes ==========================

Private Enum EmailErrorCodes
    ERR_OUTLOOK_INIT_FAILED = 1050 ' Erro ao inicializar o Outlook
    ERR_EMAIL_SEND_FAILED = 1051 ' Erro de envio de e-mail
End Enum

'==================== Propriedades =========================

Private outlookApp As Outlook.Application
Private emailItem As Outlook.MailItem

Private templatePath As String
Private templateType As String '"html", "oft" ou ""
Private htmlContent As String

'==================== Inicialização =========================

Private Sub Class_Initialize()
    On Error Resume Next
        Set outlookApp = New Outlook.Application
        If outlookApp Is Nothing Then
            Err.Raise ERR_OUTLOOK_INIT_FAILED, "ClsEmailOutlook", "Erro ao inicializar Outlook."
        End If
    On Error GoTo 0
End Sub

Private Sub Class_Terminate()
    Set emailItem = Nothing
    Set outlookApp = Nothing
End Sub

'==========================================================
'                   BLOCO: MÉTODOS PÚBLICOS
'==========================================================

Public Sub ImportTemplate(ByVal path As String)
    Dim ext As String
    ext = LCase(FileManipulation.GetFileExtension(path))

    Select Case ext
        Case "html", "htm"
            LoadHtmlTemplate path
        Case "oft"
            LoadOftTemplate path
        Case Else
            Err.Raise 1001, "ClsEmailOutlook", _
                "Formato inválido: apenas .oft e .html são suportados."
    End Select
End Sub

Public Sub SendEmail( _
        ByVal sender As String, _
        ByVal recipient As String, _
        ByVal subject As String, _
        Optional ByVal copy As String = "", _
        Optional ByVal hiddenCopy As String = "", _
        Optional ByVal htmlEmailBody As String = "", _
        Optional ByVal keysValues As Scripting.Dictionary = Nothing, _
        Optional ByVal filePaths As Collection = Nothing, _
        Optional ByVal showBeforeSend As Boolean = False)

    On Error GoTo EH

    '--- Criação do item de email
    Set emailItem = CreateEmailItem()

    '--- Cabeçalho
    ApplyRecipients sender, recipient, copy, hiddenCopy
    emailItem.subject = ReplaceKeys(subject, keysValues)

    '--- Corpo HTML conforme o tipo de template
    BuildEmailBody htmlEmailBody, keysValues, showBeforeSend

    '--- Anexos
    If Not filePaths Is Nothing Then AddAttachments filePaths

    '--- Envio
    If showBeforeSend Then
        emailItem.Display
    Else
        emailItem.Send
    End If

    Exit Sub

EH:
    Err.Raise ERR_EMAIL_SEND_FAILED, "ClsEmailOutlook.SendEmail", Err.Description
End Sub

Public Sub CleanTemplate()
    templatePath = ""
    templateType = ""
    htmlContent = ""
End Sub

'==========================================================
'               BLOCO: MÉTODOS PRIVADOS
'==========================================================

'-------------- Template OFT -------------------
Private Sub LoadOftTemplate(ByVal path As String)
    templatePath = path
    templateType = "oft"
End Sub

'-------------- Template HTML -------------------
Private Sub LoadHtmlTemplate(ByVal path As String)
    templatePath = path
    templateType = "html"

    htmlContent = FileManipulation.ReadFile(path)
    htmlContent = FixRelativePaths(path, htmlContent)
End Sub

'-------------- Criar email item -------------------
Private Function CreateEmailItem() As Outlook.MailItem
    If templateType = "oft" Then
        Set CreateEmailItem = outlookApp.CreateItemFromTemplate(templatePath)
    Else
        Set CreateEmailItem = outlookApp.CreateItem(olMailItem)
    End If
End Function

'-------------- Aplicar destinatários -------------------
Private Sub ApplyRecipients( _
        ByVal sender As String, _
        ByVal recipient As String, _
        ByVal copy As String, _
        ByVal hiddenCopy As String)

    With emailItem
        .SentOnBehalfOfName = sender
        .To = recipient
        If copy <> "" Then .CC = copy
        If hiddenCopy <> "" Then .BCC = hiddenCopy
    End With
End Sub

'-------------- Construção do corpo -------------------
Private Sub BuildEmailBody( _
        ByVal htmlEmailBody As String, _
        ByVal keys As Scripting.Dictionary, _
        ByVal showBefore As Boolean)

    Select Case templateType

        Case "html"
            Dim htmlFinal As String
            htmlFinal = ReplaceKeys(htmlContent, keys)
            emailItem.HTMLBody = htmlFinal

        Case "oft"
            ReplaceKeysInMailBody keys

        Case Else 'Sem template
            If showBefore Then
                'Outlook gerencia assinatura
                emailItem.HTMLBody = htmlEmailBody
            Else
                'Preserva assinatura padrão
                emailItem.HTMLBody = htmlEmailBody & emailItem.HTMLBody
            End If

            emailItem.HTMLBody = ReplaceKeys(emailItem.HTMLBody, keys)
    End Select
End Sub

'-------------- Substituir chaves no texto -------------------
Private Function ReplaceKeys(ByVal text As String, ByVal keys As Scripting.Dictionary) As String
    Dim k As Variant
    ReplaceKeys = text

    If Not keys Is Nothing Then
        For Each k In keys.keys
            ReplaceKeys = Replace(ReplaceKeys, k, keys(k))
        Next k
    End If
End Function

'-------------- Substituir chaves dentro do HTMLBody do OFT -------------------
Private Sub ReplaceKeysInMailBody(ByVal keys As Scripting.Dictionary)
    If keys Is Nothing Then Exit Sub

    Dim k As Variant
    For Each k In keys.keys
        emailItem.HTMLBody = Replace(emailItem.HTMLBody, k, keys(k))
    Next k
End Sub

'-------------- Converter caminhos relativos do template HTML -------------------
Private Function FixRelativePaths(ByVal htmlFile As String, ByVal html As String) As String
    Dim folder As String
    Dim folderName As String
    Dim folderEncoded As String

    folder = Left(htmlFile, InStrRev(htmlFile, "\") - 1)
    folderName = Replace(Replace(Dir(htmlFile), ".html", "_arquivos"), ".htm", "_arquivos")
    folderEncoded = Replace(folderName, " ", "%20")

    FixRelativePaths = Replace(html, folderEncoded & "/", folder & "\" & folderName & "\")
End Function

'-------------- Adicionar anexos -------------------
Private Sub AddAttachments(ByVal files As Collection)
    Dim i As Integer
    For i = 1 To files.Count
        If FileManipulation.FileExists(files(i)) Then
            emailItem.Attachments.Add files(i)
        End If
    Next i
End Sub
