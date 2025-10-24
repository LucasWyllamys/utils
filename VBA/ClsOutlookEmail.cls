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
' Atualizado em: 24/10/2025
' Pré-requisitos:
'   - Ativar bibliotecas
'       - Microsoft Outlook 16.0 Object Library: No Editor VBA: Ferramentas > Referências > Microsoft Outlook 16.0 Object Library
'       - Microsoft Scripting Runtime: No Editor VBA: Ferramentas > Referências > Microsoft Scripting Runtime
'   - Importar os módulos:
'       - FileManipulation: https://github.com/LucasWyllamys/utils/blob/main/VBA/FileManipulation.bas
'       - PublicConstantsVariables: https://github.com/LucasWyllamys/utils/blob/main/VBA/PublicConstantsVariables.bas
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
'    email.ReplaceHtmlTemplateKey chave_valor
'    email.SendEmail "remetente@exemplo.com", "destinatario@exemplo.com", "Assunto do E-mail", , , , "C:\Users\U350504\Downloads\Documentacao_Classe_Email_Outlook.pdf"
' Observações:
'   - A assinatura padrão do Outlook é preservada se não for usado um template.
'   - O método ImportHtmlTemplate deve ser chamado antes do envio para preparar o corpo do e-mail.
'   - Para anexar múltiplos arquivos, recomenda-se compactá-los em um único .zip.
'===========================================================================================================

Option Explicit

'====================================== Propriedades Privadas ========================================
 
Private pOutlookApp                 As Outlook.Application          ' Referência do Outlook
Private pEmailItem                  As Outlook.MailItem             ' Referência do E-mail
Private pHtmlTemplate               As String                       ' Armazena o HTML do template de e-mail

'=========================================== Inicialização ===========================================

Private Sub Class_Initialize()
    On Error GoTo ErrorHandler
    
    Set OutlookApp = New Outlook.Application 'Instanciamento do Outlook
    
    Exit Sub
ErrorHandler:
    Err.Raise ERR_OUTLOOK_INIT_FAILED, "ClsEmailOutlook", "Erro ao inicializar: " & Err.Description ' Lançamento de erro
End Sub

Private Sub Class_Terminate()
    Set OutlookApp = Nothing
    Set EmailItem = Nothing
    HtmlTemplate = ""
End Sub

'==================================== Métodos Assessores Privados ====================================

' Propriedades encapsuladas para acesso aos objetos privados
Private Property Get OutlookApp() As Outlook.Application: Set OutlookApp = pOutlookApp: End Property

Private Property Set OutlookApp(objeto As Outlook.Application): Set pOutlookApp = objeto: End Property

Private Property Get EmailItem() As Outlook.MailItem: Set EmailItem = pEmailItem: End Property

Private Property Set EmailItem(objeto As Outlook.MailItem): Set pEmailItem = objeto: End Property

Private Property Get HtmlTemplate() As String: HtmlTemplate = pHtmlTemplate: End Property

Private Property Let HtmlTemplate(value As String): pHtmlTemplate = value: End Property

'=========================================== Métodos Públicos ===========================================

' Descrição: Importa um template HTML de e-mail
' Parâmetros:
'   - HtmlTemplatePath: caminho completo do arquivo HTML
Public Sub ImportHtmlTemplate(HtmlTemplatePath As String)
    Dim file As Integer
    Dim HtmlContent As String
    
    On Error GoTo ErrorHandler
    
    HtmlContent = FileManipulation.ReadFile(HtmlTemplatePath)
    HtmlTemplate = AdjustTemplatePath(HtmlTemplatePath, HtmlContent)

    Exit Sub
ErrorHandler:
    Err.Raise ERR_FILE_NOT_FOUND, "ClsEmailOutlook", "Erro ao importar template HTML: " & Err.Description
End Sub

' Descrição: Substitui chaves no template HTML por valores do dicionário
' Parâmetros:
'   - dictionary: objeto Scripting.Dictionary com pares chave-valor
Public Sub ReplaceHtmlTemplateKey(dictionary As Scripting.dictionary)
    Dim key As Variant
    
    On Error GoTo ErrorHandler
    
    If Not dictionary Is Nothing Then ' Verifica se o dicionário não está vazio
        For Each key In dictionary.Keys   ' Itera sobre todos as chaves do dicionário
            HtmlTemplate = Replace(HtmlTemplate, key, dictionary(key)) ' Substitui os valores das respectivas chaves
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

    Set EmailItem = OutlookApp.CreateItem(olMailItem) 'Cria um item de e-mail
    
    With EmailItem
        If showBeforeSend Then .Display 'Exibir e-mail
        
        .SentOnBehalfOfName = sender
        .To = recipient
        If copy <> "" Then .CC = copy
        If hiddenCopy <> "" Then .BCC = hiddenCopy
        .subject = subject
        
        If HtmlTemplate = "" Then
            .HTMLBody = htmlEmailBody & .HTMLBody ' Corpo + Assinatura padrão
        Else
            .HTMLBody = HtmlTemplate
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
    HtmlTemplate = ""
End Sub

'=========================================== Métodos Privados ===========================================

' Descrição: Corrige caminhos relativos no HTML para caminhos absolutos
' Parâmetros:
'   - HtmlTemplatePath: caminho do arquivo HTML
'   - HtmlContent: conteúdo lido do arquivo
Private Function AdjustTemplatePath(HtmlTemplatePath As String, HtmlContent As String) As String
    Dim templateName As String
    Dim templateFolderPath As String
    
    On Error GoTo ErrorHandler
    templateName = Replace(Replace(Dir(HtmlTemplatePath), ".html", "_arquivos"), ".htm", "_arquivos")
    templateFolderPath = Left(HtmlTemplatePath, InStrRev(HtmlTemplatePath, "\") - 1)
    
    AdjustTemplatePath = Replace(HtmlContent, templateName & "/", templateFolderPath & "\" & templateName & "\")
    Exit Function
ErrorHandler:
    AdjustTemplatePath = HtmlContent ' Retorna o conteúdo original em caso de erro
End Function
