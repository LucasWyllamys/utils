VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WordTemplateProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================================================================================
' Classe VBA: WordTemplateProcessor
' Versão: 3.0.0
' Autor: Lucas Wyllamys Carmo da Silva
' Criada em: 04/11/2025
' Atualizada em: 09/03/2026
'
' Descrição geral:
'   Classe responsável por manipular documentos Microsoft Word baseados em templates (.doc ou .docx).
'   Permite:
'       • Carregar um arquivo modelo (template).
'       • Substituir chaves de texto por valores definidos pelo usuário.
'       • Exportar o documento final nos formatos DOCX ou PDF.
'       • Garantir o fechamento seguro do documento e da instância Word, evitando processos órfãos.
'
' Funcionalidades principais:
'   • Inicialização automática e silenciosa da aplicação Word (instância isolada).
'   • Carregamento de templates em modo somente leitura, evitando alterações indevidas.
'   • Substituição rápida no conteúdo principal (wordDoc.Content) usando pares chave/valor.
'   • Exportação para DOCX ou PDF com tratamento de exceções e sobrescrita controlada.
'   • Liberação segura de recursos e memória, finalizando corretamente o Word.
'
' Pré-requisitos:
'   • Bibliotecas necessárias (VBA -> Ferramentas > Referências):
'       - Microsoft Word xx.x Object Library
'       - Microsoft Scripting Runtime
'
' Dependências externas:
'   • Módulo utilitário "FolderAndFileManipulator"
'       https://github.com/LucasWyllamys/utils/blob/main/VBA/FolderAndFileManipulator.bas
'   • Módulo utilitário "PublicConstants"
'       https://github.com/LucasWyllamys/utils/blob/main/VBA/PublicConstants.bas
'
' Observações importantes:
'   • A propriedade privada `pWordApp` mantém a instância da aplicação Word.
'   • A propriedade privada `pWordDoc` referencia o documento atualmente carregado.
'   • Esta classe opera sempre com um documento por vez.
'   • A enumeração `ExportFileType` define os formatos suportados (DOCX ou PDF).
'   • A substituição de chaves neste processador ocorre somente no conteúdo principal do documento,
'     não abrangendo cabeçalhos, rodapés ou shapes — otimização proposital para desempenho.
'
' Exemplo de uso:
'
'    Dim templateProcessor As ITemplateProcessor   ' Programar contra a interface
'    Dim keysValues As Scripting.Dictionary
'    Dim i As Integer
'
'    Set templateProcessor = New WordTemplateProcessor   ' Implementação concreta
'    Set keysValues = New Scripting.Dictionary
'
'    For i = 1 To 10
'
'        With keysValues
'            .Add "{NOME CLIENTE}", "LUCAS WYLLAMYS CARMO DA SILVA - " & i
'            .Add "{INSTALAÇÃO}", "16278556 - " & i
'            .Add "{CONTA CONTRATO}", "7054321568 - " & i
'            .Add "{NÚMERO DO PN}", "100245684 - " & i
'            .Add "{TN}", "200245685 - " & i
'            .Add "{NOTA}", "9245687543 - " & i
'            .Add "{DATA}", "02/03/2026 - " & i
'        End With
'
'        templateProcessor.ImportTemplate _
'            "C:\Users\U350504\Documents\GitHub\projetos-em-vba\Olympus\templates\hephaestus\Contrato Instalação A.docx"
'
'        templateProcessor.Export _
'            "C:\Users\U350504\Documents\GitHub\projetos-em-vba\Olympus\files\hephaestus", _
'            "Teste - " & i, _
'            pdf, _
'            keysValues
'
'        keysValues.RemoveAll
'    Next i
'
'    templateProcessor.Finish      ' Encerra Word corretamente
'
'    Set keysValues = Nothing
'    Set templateProcessor = Nothing
'
'    MsgBox "Concluído!"
'
'===========================================================================================================

Option Explicit

Implements ITemplateProcessor

Private Const CLASS_NAME As String = "WordTemplateProcessor"

' Instâncias internas da aplicação Word e do documento aberto
Private pWordApp As word.Application
Private pWordDoc As word.Document

'============================= Assessores =============================

Private Property Get wordApp() As word.Application
    Set wordApp = pWordApp
End Property

Private Property Set wordApp(app As word.Application)
    Set pWordApp = app
End Property

Private Property Get wordDoc() As word.Document
    Set wordDoc = pWordDoc
End Property

Private Property Set wordDoc(doc As word.Document)
    Set pWordDoc = doc
End Property

' Controla se a janela do Word será visível ou não.
Public Property Let Visible(ByVal value As Boolean)
    If Not wordApp Is Nothing Then wordApp.Visible = value
End Property

'============================= Ciclo de vida ================================

' Construtor da classe (executa automaticamente ao criar New WordTemplateProcessor)
' Inicializa o Word de forma silenciosa
Private Sub Class_Initialize()
    On Error GoTo Fail

    ' Instancia o Word
    Set wordApp = New word.Application
    wordApp.Visible = False                        ' Word invisível
    wordApp.DisplayAlerts = word.WdAlertLevel.wdAlertsNone   ' Sem alerts

    Exit Sub

' Tratamento de erro da inicialização
Fail:
    Set wordApp = Nothing
    Err.Raise ERR_INIT_FILE_FAILED, CLASS_NAME & ".Class_Initialize", _
              "Erro ao inicializar o Word: " & Err.Description
End Sub

' Destrutor da classe — executa quando o objeto é destruído via Set obj = Nothing
Private Sub Class_Terminate()
    ITemplateProcessor_Finish
End Sub

'============================= Implementação da Interface ======================

' Importa um template Word (.doc ou .docx)
Private Sub ITemplateProcessor_ImportTemplate(ByVal wordTemplatePath As String)
    Dim ext As String

    ' Valida se o arquivo existe
    If Not FolderAndFileManipulator.FileExists(wordTemplatePath) Then
        Err.Raise ERR_FILE_NOT_FOUND, CLASS_NAME & ".ImportTemplate", "Arquivo de template não existe."
    End If

    ' Obtém e valida a extensão
    ext = FolderAndFileManipulator.GetFileExtension(wordTemplatePath)

    Select Case ext
        Case "docx", "doc"
            ' Abre o template em modo somente leitura
            Set wordDoc = wordApp.Documents.Open( _
                    fileName:=wordTemplatePath, _
                    ReadOnly:=True, _
                    AddToRecentFiles:=False)

            ' Caso excepcional: arquivo não abriu
            If wordDoc Is Nothing Then
                Err.Raise ERR_OPEN_FILE_FAILED, CLASS_NAME & ".ImportTemplate", "Falha ao carregar o template."
            End If

        Case Else
            ' Extensões inválidas
            Err.Raise ERR_INVALID_FILE_TYPE, CLASS_NAME & ".ImportTemplate", _
                      "Formato inválido: apenas .docx ou .doc são suportados."
    End Select
End Sub

' Exporta o documento substituindo as chaves (se houver)
Private Function ITemplateProcessor_Export( _
    ByVal targetFolderPath As String, _
    ByVal fileName As String, _
    ByVal fileType As ExportFileType, _
    Optional ByVal keysValues As Scripting.Dictionary = Nothing) As String

    Dim outPath As String
    Dim safeName As String
    
    ' Verifica o tipo de arquivo de exportação
    If Not (fileType = flDOCX Or fileType = flPDF) Then
        ITemplateProcessor_ResetTemplate
        Err.Raise ERR_INVALID_FILE_TYPE, CLASS_NAME & ".Export", "Tipo de exportação inválido. Tipos de arquivos suportados .docx ou .pdf"
    End If

    ' Verifica pasta destino
    If Not FolderAndFileManipulator.FolderExists(targetFolderPath) Then
        ITemplateProcessor_ResetTemplate
        Err.Raise ERR_FOLDER_NOT_FOUND, CLASS_NAME & ".Export", "A pasta destino não existe!"
    End If

    ' Verifica se existe documento carregado
    If wordDoc Is Nothing Then
        ITemplateProcessor_ResetTemplate
        Err.Raise ERR_FILE_NOT_FOUND, CLASS_NAME & ".Export", "Nenhum template carregado."
    End If

    
    ' Verifica se o nome do arquivo é válido
    safeName = FolderAndFileManipulator.SanitizeFileName(fileName) ' Sanitiza o nome do arquivo
    If LenB(safeName) = 0 Then
        ITemplateProcessor_ResetTemplate
        Err.Raise ERR_SAVE_FILE_FAILED, CLASS_NAME & ".Export", "Nome de arquivo inválido."
    End If

    ' Caso tenha chaves, substitui
    If Not keysValues Is Nothing Then
        ReplaceKeys keysValues, InContent
    End If

    On Error GoTo ErrHandler

    ' Decide o formato de saída
    Select Case fileType

        ' Exportação como DOCX
        Case flDOCX
            outPath = FolderAndFileManipulator.BuildPath(targetFolderPath, safeName, "docx")  ' Monta o caminho do arquivo que será exportado
            FolderAndFileManipulator.EnsureOverwrite outPath ' Se existir, remove arquivo antes de salvar/exportar
            wordDoc.SaveAs2 fileName:=outPath, FileFormat:=word.WdSaveFormat.wdFormatXMLDocument ' Exporta o arquivo

        ' Exportação como PDF
        Case flPDF
            outPath = FolderAndFileManipulator.BuildPath(targetFolderPath, safeName, "pdf")
            FolderAndFileManipulator.EnsureOverwrite outPath
            wordDoc.ExportAsFixedFormat _
                OutputFileName:=outPath, _
                ExportFormat:=word.WdExportFormat.wdExportFormatPDF, _
                OpenAfterExport:=False, _
                OptimizeFor:=word.WdExportOptimizeFor.wdExportOptimizeForPrint, _
                Range:=word.WdExportRange.wdExportAllDocument, _
                item:=word.WdExportItem.wdExportDocumentContent, _
                IncludeDocProps:=True, _
                KeepIRM:=True, _
                CreateBookmarks:=word.WdExportCreateBookmarks.wdExportCreateNoBookmarks, _
                DocStructureTags:=True, _
                BitmapMissingFonts:=True, _
                UseISO19005_1:=False
            
    End Select

    ' Retorna o caminho do arquivo gerado
    ITemplateProcessor_Export = outPath

    ' Fecha o documento para liberar memória
    ITemplateProcessor_ResetTemplate
    Exit Function

' Tratamento de erro
ErrHandler:
    ITemplateProcessor_ResetTemplate
    Err.Raise ERR_SAVE_FILE_FAILED, CLASS_NAME & ".Export", _
              "Falha ao salvar arquivo: " & Err.Description
End Function

' Fecha o documento atual (se existir)
Private Sub ITemplateProcessor_ResetTemplate()
    If Not wordDoc Is Nothing Then
        wordDoc.Close SaveChanges:=False
        Set wordDoc = Nothing
    End If
End Sub

' ' Finaliza completamente a instância do Word
Private Sub ITemplateProcessor_Finish()
    ITemplateProcessor_ResetTemplate
    If Not wordApp Is Nothing Then
        wordApp.Quit SaveChanges:=False     ' Fecha o Word sem salvar nada
        Set wordApp = Nothing               ' Libera memória
    End If
End Sub

'============================= Helpers privados =============================

Private Sub ReplaceKeys(ByVal keysValues As Scripting.Dictionary, ByVal replacementType As TypeOfReplacementInWord)
    If wordDoc Is Nothing Then Exit Sub
    If keysValues Is Nothing Then Exit Sub
    If keysValues.Count = 0 Then Exit Sub

    On Error GoTo ErrHandler
    
    If replacementType = InContent Then
        ReplaceKeysInContent keysValues ' Substituição rápida apenas no Content
    ElseIf replacementType = InAllStories Then
        ReplaceKeysInAllStories keysValues ' Substitui em TODO o documento (texto, headers, footers, shapes…)
    End If

    Exit Sub

ErrHandler:
    ITemplateProcessor_ResetTemplate
    Err.Raise ERR_WRITE_FILE_FAILED, CLASS_NAME & ".ReplaceKeys", _
              "Erro ao substituir chaves no template: " & Err.Description
End Sub

' Substituição rápida apenas no Content
Private Sub ReplaceKeysInContent(ByVal keysValues As Scripting.Dictionary)
    Dim key As Variant
    Dim findText As String
    Dim replaceText As String

    With wordDoc.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = word.WdFindWrap.wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        For Each key In keysValues.keys
            If Not IsNull(key) Then
                findText = Trim$(CStr(key))
                If Len(findText) > 0 Then
                    replaceText = CStr(keysValues(key))
                    .text = findText
                    .Replacement.text = replaceText
                    .Execute Replace:=word.WdReplace.wdReplaceAll
                End If
            End If
        Next key
    End With
End Sub

' Substitui texto em todas as áreas do documento (corpo, cabeçalho, rodapé e shapes)
Private Sub ReplaceKeysInAllStories(ByVal keysValues As Scripting.Dictionary)
    Dim story As word.Range
    Dim shp As word.Shape
    Dim key As Variant, findText As String, replaceText As String
    
    For Each key In keysValues.keys
        If Not IsNull(key) Then
            findText = Trim$(CStr(key))
            If Len(findText) > 0 Then
                replaceText = CStr(keysValues(key))
                ' Todas as StoryRanges (principal, cabeçalhos, rodapés, notas…)
                For Each story In wordDoc.StoryRanges
                    With story.Find
                        .ClearFormatting
                        .Replacement.ClearFormatting
                        .text = findText
                        .Replacement.text = replaceText
                        .Forward = True
                        .Wrap = word.WdFindWrap.wdFindContinue
                        .Format = False
                        .MatchCase = False
                        .MatchWholeWord = False
                        .MatchWildcards = False
                        .Execute Replace:=word.WdReplace.wdReplaceAll
                    End With
                Next story
                ' Todas as Shapes que possuem TextFrame (caixas de texto, WordArt, etc.)
                For Each shp In wordDoc.Shapes
                    If shp.TextFrame.HasText Then
                        With shp.TextFrame.TextRange.Find
                            .ClearFormatting
                            .Replacement.ClearFormatting
                            .text = findText
                            .Replacement.text = replaceText
                            .Forward = True
                            .Wrap = word.WdFindWrap.wdFindContinue
                            .Format = False
                            .MatchCase = False
                            .MatchWholeWord = False
                            .MatchWildcards = False
                            .Execute Replace:=word.WdReplace.wdReplaceAll
                        End With
                    End If
                Next shp
            End If
        End If
    Next key
End Sub
