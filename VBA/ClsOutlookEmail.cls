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
' Versão: 5.0.0
' Autor: Lucas Wyllamys Carmo da Silva
' Criado em: 22/10/2025
' Atualizado em: 24/01/2026
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
'    email.ReplaceKeysBody chaveValor
'    email.SendEmail "remetente@exemplo.com", "destinatario@exemplo.com", "Assunto do E-mail", , , , "C:\Users\U350504\Downloads\Documentacao_Classe_Email_Outlook.pdf"
' Observações:
'   - A assinatura padrão do Outlook é preservada se não for usado um template.
'   - O método ImportTemplate deve ser chamado antes do envio para preparar o corpo do e-mail.
'   - Para anexar múltiplos arquivos, recomenda-se compactá-los em um único .zip.
'===========================================================================================================

Option Explicit

'====================================== Propriedades Privadas ========================================
 
Private outlookApp          As Outlook.Application ' Referência do Outlook
Private emailItem           As Outlook.MailItem ' Referência do E-mail
Private htmlContent         As String ' Armazena o conteúdo HTML template de e-mail HTML
Private tempPath            As String ' Armazena o caminho do template
Private templateType        As String ' Armazena o tipo de template selecionado

'=========================================== Inicialização ===========================================

Private Sub Class_Initialize()
    On Error GoTo ErrorHandler
    
    Set outlookApp = New Outlook.Application 'Instanciamento do Outlook
    
    Exit Sub
ErrorHandler:
    Err.Raise ERR_OUTLOOK_INIT_FAILED, "ClsEmailOutlook", "Erro ao inicializar Outlook: " & Err.Description ' Lançamento de erro
End Sub

Private Sub Class_Terminate()
    CleanTemplate
    Set outlookApp = Nothing
    Set emailItem = Nothing
End Sub

'=========================================== Métodos Públicos ===========================================

' Descrição: Importa um template de e-mail .oft ou .htm/.htm
' Parâmetros:
'   - TemplatePath: caminho completo do arquivo
Public Sub ImportTemplate(templatePath As String)
    If FileManipulation.GetFileExtension(templatePath) = "oft" Then
        tempPath = templatePath
        templateType = "oft"
    ElseIf FileManipulation.GetFileExtension(templatePath) = "htm" Or FileManipulation.GetFileExtension(templatePath) = "html" Then
        tempPath = templatePath
        htmlContent = FileManipulation.ReadFile(templatePath)
        htmlContent = AdjustTemplatePath(templatePath, htmlContent)
        templateType = "html"
    Else
        Err.Raise ERR_INVALID_FILE_TYPE, "ClsEmailOutlook", "Formato de template não suportado. Somente arquivos .oft ou .html são permitidos."
    End If
End Sub

' Descrição: Substitui chaves no template por valores do dicionário
' Parâmetros:
'   - dictKeysValues: objeto Scripting.Dictionary com pares chave-valor
Public Sub ReplaceKeysBody(keysValues As Scripting.Dictionary)
    Dim key As Variant
    
    If Not keysValues Is Nothing Then ' Verifica se o dicionário não está vazio
        If templateType = "html" Then ' Substitui as chaves pelos valores no conteúdo HTML
            For Each key In keysValues.Keys   ' Itera sobre todos as chaves do dicionário
                htmlContent = Replace(htmlContent, key, keysValues(key)) ' Substitui os valores das respectivas chaves
            Next key
        Else ' Substitui as chaves pelos valores direto no .HTMLBody
            For Each key In keysValues.Keys
                emailItem.HTMLBody = Replace(emailItem.HTMLBody, key, keysValues(key))
            Next key
        End If
    End If
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
                        Optional copy As String = "", _
                        Optional hiddenCopy As String = "", _
                        Optional htmlEmailBody As String = "", _
                        Optional keysValues As Scripting.Dictionary = "", _
                        Optional filePaths As Collection = "", _
                        Optional showBeforeSend As Boolean = False)
    Dim corpoEmail As String
    
    On Error GoTo ErrorHandler
    
    If templateType = "oft" Then
        Set emailItem = outlookApp.CreateItemFromTemplate(tempPath) 'Cria um item de e-mail a partir de um template .oft
    ElseIf templateType = "html" Then
        Set emailItem = outlookApp.CreateItem(olMailItem) 'Cria um novo item de e-mail
    End If
    
    With emailItem
        If showBeforeSend Then .Display 'Exibir e-mail
        .SentOnBehalfOfName = sender
        .To = recipient
        If copy <> "" Then .CC = copy
        If hiddenCopy <> "" Then .BCC = hiddenCopy
        .subject = subject
            
        If templateType = "oft" Then
            If Not keysValues Is Nothing Then
                ReplaceKeysBody keysValues
            End If
        ElseIf templateType = "html" Then
            If Not keysValues Is Nothing Then
                ReplaceKeysBody keysValues
            End If
            .HTMLBody = htmlContent
        Else
            .HTMLBody = htmlEmailBody & .HTMLBody ' Corpo + Assinatura padrão
            If Not keysValues Is Nothing Then
                ReplaceKeysBody keysValues
            End If
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
Public Sub CleanTemplate()
    htmlContent = ""
    tempPath = ""
    templateType = ""
End Sub

'=========================================== Métodos Privados ===========================================

' Descrição: Corrige caminhos relativos no HTML para caminhos absolutos
' Parâmetros:
'   - HtmlTemplatePath: caminho do arquivo HTML
'   - HtmlContent: conteúdo lido do arquivo
Private Function AdjustTemplatePath(htmlTemplatePath As String, htmlContent As String) As String
    Dim templateName As String
    Dim templateFolderPath As String
    Dim templateNameModify As String
    
    templateFolderPath = Left(htmlTemplatePath, InStrRev(htmlTemplatePath, "\") - 1)
    templateName = Replace(Replace(Dir(htmlTemplatePath), ".html", "_arquivos"), ".htm", "_arquivos")
    templateNameModify = Replace(templateName, " ", "%20")
    AdjustTemplatePath = Replace(htmlContent, templateNameModify & "/", templateFolderPath & "\" & templateName & "\")
End Function
