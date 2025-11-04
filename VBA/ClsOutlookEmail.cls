VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsOutlookEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================================================================================
' Classe VBA: ClsOutlookEmail - Envio de E-mail
' Descrição: Esta classe permite o envio de e-mails utilizando o Microsoft Outlook, com suporte a:
'   - Corpo do e-mail em HTML (com ou sem template).
'   - Inclusão de anexos.
'   - Cópia (CC) e cópia oculta (BCC).
'   - Substituição de chaves em templates HTML.
' Versão: 3.0.0
' Autor: Lucas Wyllamys Carmo da Silva
' Criado em: 22/10/2025
' Atualizado em: 04/11/2025
' Pré-requisitos:
'   - Ativar bibliotecas
'       - Microsoft Outlook 16.0 Object Library: No Editor VBA: Ferramentas > Referências > Microsoft Outlook 16.0 Object Library
'       - Microsoft Scripting Runtime: No Editor VBA: Ferramentas > Referências > Microsoft Scripting Runtime
'   - Importar os módulos:
'       - FileManipulation: https://github.com/LucasWyllamys/utils/blob/main/VBA/FileManipulation.bas
'       - PublicConstants: https://github.com/LucasWyllamys/utils/blob/main/VBA/PublicConstants.bas
' Como criar um template HTML do e-mail do outlook:
'    1. Abra o e-mail no Outlook
'    2. Escreva e formate o corpo do e-mail, incluindo variáveis como [NOME], [DATA], etc.
'    3. Clique na guia "Arquivo"
'    4. Clique em "Salvar como"
'    5. Escolha o diretório onde o template será salvo
'    6. Modifique o tipo para: HTML (*.htm, *.html)
'    7. Clique em "Salvar"
' Exemplo de uso:
'    Dim email As ClsOutlookEmail ' Declara como objeto do outlook
'    Dim chave_valor As Scripting.Dictionary ' Declara como dicionário
'
'    Set email = New ClsOutlookEmail ' Instancia o objeto
'    Set chave_valor = New Scripting.Dictionary ' Instancia o dicionário
'
'    email.ImportHtmlTemplate "C:\Users\U350504\Desktop\Classe de interação com arquivo JSON\Template.htm"
'
'    chave_valor.Add "[CHAVE1]", "Olá," ' Adiciona chave-valor
'    chave_valor.Add "[CHAVE2]", "Mundo!" ' Adiciona chave-valor
'
'    email.ReplaceKeys chave_valor
'    email.SendEmail "remetente@exemplo.com", "destinatario@exemplo.com", "Assunto do E-mail", , , , "C:\Users\U350504\Downloads\Documentacao_Classe_Email_Outlook.pdf"
' Observações:
'   - A assinatura padrão do Outlook é preservada se não for usado um template.
'   - O método ImportHtmlTemplate deve ser chamado antes do envio para preparar o corpo do e-mail.
'   - Para anexar múltiplos arquivos, recomenda-se compactá-los em um único .zip.
'===========================================================================================================

Option Explicit

'====================================== Propriedades Privadas ========================================
 
Private pOutlookApp     As Outlook.Application ' Referência do Outlook
Private pEmailItem      As Outlook.MailItem ' Referência do E-mail
Private pHtmlTemplate   As String ' Armazena o HTML do template de e-mail

'==================================== Métodos Assessores Privados ====================================

Private Property Get outlookApp() As Outlook.Application: Set outlookApp = pOutlookApp: End Property

Private Property Set outlookApp(object As Outlook.Application): Set pOutlookApp = object: End Property

Private Property Get emailItem() As Outlook.MailItem: Set emailItem = pEmailItem: End Property

Private Property Set emailItem(object As Outlook.MailItem): Set pEmailItem = object: End Property

Private Property Get htmlTemplate() As String: htmlTemplate = pHtmlTemplate: End Property

Private Property Let htmlTemplate(value As String): pHtmlTemplate = value: End Property

'=========================================== Inicialização ===========================================

Private Sub Class_Initialize()
    On Error GoTo ErrorHandler
    
    Set outlookApp = New Outlook.Application 'Instanciamento do Outlook
    
    Exit Sub
ErrorHandler:
    Err.Raise ERR_OUTLOOK_INIT_FAILED, "ClsEmailOutlook", "Erro ao inicializar Outlook: " & Err.Description ' Lançamento de erro
End Sub

Private Sub Class_Terminate()
    Set outlookApp = Nothing
    Set emailItem = Nothing
    htmlTemplate = ""
End Sub

'=========================================== Métodos Públicos ===========================================

' Descrição: Importa um template HTML de e-mail
' Parâmetros:
'   - htmlTemplatePath: caminho completo do arquivo HTML
Public Sub ImportHtmlTemplate(htmlTemplatePath As String)
    Dim HtmlContent As String
    
    On Error GoTo ErrorHandler
    
    HtmlContent = FileManipulation.ReadFile(htmlTemplatePath)
    htmlTemplate = AdjustTemplatePath(htmlTemplatePath, HtmlContent)

    Exit Sub
ErrorHandler:
    Err.Raise ERR_FILE_NOT_FOUND, "ClsEmailOutlook", "Erro ao importar template HTML: " & Err.Description
End Sub

' Descrição: Substitui chaves no template HTML por valores do dicionário
' Parâmetros:
'   - dictKeysValues: objeto Scripting.Dictionary com pares chave-valor
Public Sub ReplaceKeys(dictKeysValues As Scripting.Dictionary)
    Dim key As Variant
    
    On Error GoTo ErrorHandler
    
    If Not Dictionary Is Nothing Then ' Verifica se o dicionário não está vazio
        For Each key In Dictionary.Keys   ' Itera sobre todos as chaves do dicionário
            htmlTemplate = Replace(htmlTemplate, key, Dictionary(key)) ' Substitui os valores das respectivas chaves
        Next key
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise ERR_WRITE_FILE_FAILED, "ClsEmailOutlook", "Erro ao substituir chaves no template HTML: " & Err.Description
End Sub

' Descrição: Envia um e-mail utilizando o Microsoft Outlook, com suporte a corpo HTML (com ou sem template), anexos, cópia (CC), cópia oculta (BCC) e controle da exibição antes do envio.
' Parâmetros:
'   - sender (String): Endereço de e-mail do remetente.
'   - recipient (String): Endereço de e-mail do destinatário.
'   - subject (String): Assunto do e-mail.
'   - htmlEmailBody (String, opcional): Corpo do e-mail em HTML, utilizado apenas se nenhum template for carregado.
'   - copy (String, opcional): Endereços de e-mail para cópia (CC), separados por ponto e vírgula.
'   - hiddenCopy (String, opcional): Endereços de e-mail para cópia oculta (BCC), separados por ponto e vírgula.
'   - attachmentPath (String, opcional): Caminho completo do arquivo a ser anexado. Para múltiplos arquivos, recomenda-se compactá-los em um único .zip.
'   - showBeforeSend (Boolean, opcional): Define se o e-mail será exibido antes do envio. Esse parâmetro afeta diretamente a aplicação da assinatura padrão do Outlook.
'       - True: O e-mail é exibido antes do envio. A assinatura padrão do Outlook (ou do template) é aplicada completa, incluindo imagens e formatação.
'       - False com template HTML: O e-mail é enviado automaticamente. A assinatura padrão pode ser aplicada parcialmente, geralmente sem imagens.
'       - False sem template HTML: O e-mail é enviado automaticamente sem assinatura, a menos que o corpo (htmlEmailBody) inclua manualmente uma.
Public Sub SendEmail(sender As String, recipient As String, subject As String, _
                        Optional htmlEmailBody As String = "", _
                        Optional copy As String = "", _
                        Optional hiddenCopy As String = "", _
                        Optional attachmentPath As String = "", _
                        Optional showBeforeSend As Boolean = True)
    Dim corpoEmail As String
    
    On Error GoTo ErrorHandler

    Set emailItem = outlookApp.CreateItem(olMailItem) 'Cria um item de e-mail
    
    With emailItem
        If showBeforeSend Then .Display 'Exibir e-mail
        
        .SentOnBehalfOfName = sender
        .To = recipient
        If copy <> "" Then .CC = copy
        If hiddenCopy <> "" Then .BCC = hiddenCopy
        .subject = subject
        
        If htmlTemplate = "" Then
            .HTMLBody = htmlEmailBody & .HTMLBody ' Corpo + Assinatura padrão
        Else
            .HTMLBody = htmlTemplate
        End If
        
        If attachmentPath <> "" Then
            If FileManipulation.FileExists(attachmentPath) Then
                .Attachments.Add attachmentPath
            End If
        End If
        
        .Send 'Enviar e-mail
    End With
    
    Exit Sub
ErrorHandler:
    Err.Raise ERR_EMAIL_SEND_FAILED, "ClsEmailOutlook", "Erro ao enviar e-mail: " & Err.Description
End Sub

' Descrição: Reseta o template HTML
Public Sub ResetTemplate()
    htmlTemplate = ""
End Sub

'=========================================== Métodos Privados ===========================================

' Descrição: Corrige caminhos relativos no HTML para caminhos absolutos
' Parâmetros:
'   - HtmlTemplatePath: caminho do arquivo HTML
'   - HtmlContent: conteúdo lido do arquivo
Private Function AdjustTemplatePath(htmlTemplatePath As String, HtmlContent As String) As String
    Dim templateName As String
    Dim templateFolderPath As String
    
    On Error GoTo ErrorHandler
    templateName = Replace(Replace(Dir(htmlTemplatePath), ".html", "_arquivos"), ".htm", "_arquivos")
    templateFolderPath = Left(htmlTemplatePath, InStrRev(htmlTemplatePath, "\") - 1)
    
    AdjustTemplatePath = Replace(HtmlContent, templateName & "/", templateFolderPath & "\" & templateName & "\")
    Exit Function
ErrorHandler:
    AdjustTemplatePath = HtmlContent ' Retorna o conteúdo original em caso de erro
End Function
