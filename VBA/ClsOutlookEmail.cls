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
' Versão: 4.2.0
' Autor: Lucas Wyllamys Carmo da Silva
' Criado em: 22/10/2025
' Atualizado em: 21/01/2026
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
'    Dim chaveValor As Scripting.Dictionary ' Declara como dicionário
'
'    Set email = New ClsOutlookEmail ' Instancia o objeto
'    Set chaveValor = New Scripting.Dictionary ' Instancia o dicionário
'
'    email.ImportTemplate "C:\Users\U350504\Desktop\Classe de interação com arquivo JSON\Template.htm"
'
'    chaveValor.Add "[CHAVE1]", "Olá," ' Adiciona chave-valor
'    chaveValor.Add "[CHAVE2]", "Mundo!" ' Adiciona chave-valor
'
'    email.ReplaceKeys chaveValor
'    email.SendEmail "remetente@exemplo.com", "destinatario@exemplo.com", "Assunto do E-mail", , , , "C:\Users\U350504\Downloads\Documentacao_Classe_Email_Outlook.pdf"
' Observações:
'   - A assinatura padrão do Outlook é preservada se não for usado um template.
'   - O método ImportTemplate deve ser chamado antes do envio para preparar o corpo do e-mail.
'   - Para anexar múltiplos arquivos, recomenda-se compactá-los em um único .zip.
'===========================================================================================================

Option Explicit

'====================================== Propriedades Privadas ========================================
 
Private pOutlookApp         As Outlook.Application ' Referência do Outlook
Private pEmailItem          As Outlook.MailItem ' Referência do E-mail
Private pOriginalTemplate   As String ' Armazena o HTML do template de e-mail
Private pModifiedTemplate   As String ' Armazena o HTML do template de e-mail modificado

'==================================== Métodos Assessores Privados ====================================

Private Property Get outlookApp() As Outlook.Application: Set outlookApp = pOutlookApp: End Property

Private Property Set outlookApp(object As Outlook.Application): Set pOutlookApp = object: End Property

Private Property Get emailItem() As Outlook.MailItem: Set emailItem = pEmailItem: End Property

Private Property Set emailItem(object As Outlook.MailItem): Set pEmailItem = object: End Property

Private Property Get originalTemplate() As String: originalTemplate = pOriginalTemplate: End Property

Private Property Let originalTemplate(Value As String): pOriginalTemplate = Value: End Property

Private Property Get modifiedTemplate() As String: modifiedTemplate = pModifiedTemplate: End Property

Private Property Let modifiedTemplate(Value As String): pModifiedTemplate = Value: End Property

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
    originalTemplate = ""
End Sub

'=========================================== Métodos Públicos ===========================================

' Descrição: Importa um template HTML de e-mail
' Parâmetros:
'   - htmlTemplatePath: caminho completo do arquivo HTML
Public Sub ImportTemplate(htmlTemplatePath As String)
    Dim HtmlContent As String
    
    On Error GoTo ErrorHandler
    
    HtmlContent = FileManipulation.ReadFile(htmlTemplatePath)
    originalTemplate = AdjustTemplatePath(htmlTemplatePath, HtmlContent)
    modifiedTemplate = originalTemplate

    Exit Sub
ErrorHandler:
    Err.Raise ERR_FILE_NOT_FOUND, "ClsEmailOutlook", "Erro ao importar template HTML: " & Err.Description
End Sub

' Descrição: Substitui chaves no template HTML por valores do dicionário
' Parâmetros:
'   - dictKeysValues: objeto Scripting.Dictionary com pares chave-valor
Public Sub ReplaceKeys(dictKeysValues As Scripting.Dictionary)
    Dim key As Variant
    
    modifiedTemplate = originalTemplate
    
    On Error GoTo ErrorHandler
    
    If Not dictKeysValues Is Nothing Then ' Verifica se o dicionário não está vazio
        For Each key In dictKeysValues.Keys   ' Itera sobre todos as chaves do dicionário
            modifiedTemplate = Replace(modifiedTemplate, key, dictKeysValues(key)) ' Substitui os valores das respectivas chaves
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
'   - filePaths (Collection, opcional): Caminhos completos dos arquivos a serem anexados.
'   - showBeforeSend (Boolean, opcional): Define se o e-mail será exibido antes do envio. Esse parâmetro afeta diretamente a aplicação da assinatura padrão do Outlook.
'       - True: O e-mail é exibido antes do envio. A assinatura padrão do Outlook (ou do template) é aplicada completa, incluindo imagens e formatação.
'       - False com template HTML: O e-mail é enviado automaticamente. A assinatura padrão pode ser aplicada parcialmente, geralmente sem imagens.
'       - False sem template HTML: O e-mail é enviado automaticamente sem assinatura, a menos que o corpo (htmlEmailBody) inclua manualmente uma.
Public Sub SendEmail(sender As String, recipient As String, subject As String, _
                        Optional htmlEmailBody As String = "", _
                        Optional copy As String = "", _
                        Optional hiddenCopy As String = "", _
                        Optional filePaths As Collection = "", _
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
        
        If modifiedTemplate = "" Then
            .HTMLBody = htmlEmailBody & .HTMLBody ' Corpo + Assinatura padrão
        Else
            .HTMLBody = modifiedTemplate
        End If
        
        If Not filePaths Is Nothing Then
            Dim i As Integer
            For i = 1 To filePaths.Count
                If FileManipulation.FileExists(filePaths(i)) Then
                    .Attachments.Add filePaths(i)
                End If
            Next i
        End If
        
        .Send 'Enviar e-mail
    End With
    
    Exit Sub
ErrorHandler:
    Err.Raise ERR_EMAIL_SEND_FAILED, "ClsEmailOutlook", "Erro ao enviar e-mail: " & Err.Description
End Sub

' Descrição: Reseta o template HTML
Public Sub ResetTemplate()
    originalTemplate = ""
    modifiedTemplate = ""
End Sub

'=========================================== Métodos Privados ===========================================

' Descrição: Corrige caminhos relativos no HTML para caminhos absolutos
' Parâmetros:
'   - HtmlTemplatePath: caminho do arquivo HTML
'   - HtmlContent: conteúdo lido do arquivo
Private Function AdjustTemplatePath(htmlTemplatePath As String, HtmlContent As String) As String
    Dim templateName As String
    Dim templateFolderPath As String
    Dim templateNameModify As String
    
    On Error GoTo ErrorHandler
    templateFolderPath = Left(htmlTemplatePath, InStrRev(htmlTemplatePath, "\") - 1)
    templateName = Replace(Replace(Dir(htmlTemplatePath), ".html", "_arquivos"), ".htm", "_arquivos")
    templateNameModify = Replace(templateName, " ", "%20")
    AdjustTemplatePath = Replace(HtmlContent, templateNameModify & "/", templateFolderPath & "\" & templateName & "\")
    Exit Function
ErrorHandler:
    AdjustTemplatePath = HtmlContent ' Retorna o conteúdo original em caso de erro
End Function
