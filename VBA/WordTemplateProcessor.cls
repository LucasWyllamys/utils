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
' Classe VBA: WordTemplateProcessor - Processador de Templates Word
' Versão: 1.5.0
' Autor: Lucas Wyllamys Carmo da Silva
' Criado em: 04/11/2025
' Atualizado em: 18/11/2025
'
' Descrição:
'   Esta classe encapsula a lógica de manipulação de documentos Word baseados em templates (.docx),
'   permitindo a importação, substituição de chaves por valores e exportação em formatos DOCX ou PDF.
'
' Recursos:
'   - Inicialização automática da aplicação Microsoft Word.
'   - Importação de arquivos .docx como templates.
'   - Substituição de chaves no documento com base em um dicionário de pares chave/valor.
'   - Exportação do documento em formato DOCX ou PDF.
'   - Encerramento seguro da aplicação Word e liberação de recursos.
'
' Pré-requisitos:
'   - Ativar bibliotecas
'       - Microsoft Scripting Runtime: Ferramentas > Referências > Microsoft Word xx.x Object Library
'       - Microsoft Word xx.x Object Library: Ferramentas > Referências > Microsoft Scripting Runtime
'   - Importar módulo: PublicConstants
'
' Observações:
'   - A propriedade `wordApp` representa a instância da aplicação Word.
'   - A propriedade `wordDoc` representa o documento Word carregado.
'   - A propriedade `templatePath` armazena o caminho do template importado.
'   - A enumeração `flType` define os formatos de exportação disponíveis: DOCX ou PDF.
'
' Exemplo de uso:
'   Dim wordProcessor As clsWordTemplateProcessor
'   Dim dictKeysValue As Scripting.Dictionary
'
'   Set wordProcessor = New clsWordTemplateProcessor
'
'   wordProcessor.ImportWordTemplate "C:\Templates\Contrato.docx"
'
'   dict.Add "{NOME}", "Lucas Wyllamys"
'   dict.Add "{DATA}", "04/11/2025"
'
'   wordProcessor.ReplaceKeys dictKeysValue
'   wordProcessor.Export "C:\Contratos", "Contrato_Lucas", pdf
'   wordProcessor.CloseWordApp
'===========================================================================================================

Option Explicit

'====================================== Propriedades Privadas ========================================

Private pWordApp As Word.Application
Private pWordDoc As Word.Document

Private pTemplatePath As String

'====================================== Constantes =========================================

Private Const CLASS_NAME As String = "clsWordTemplateProcessor"

'=============================================== Enum ===============================================

Public Enum FileType
    FileTypePDF = 1
    FileTypeDOCX = 2
End Enum

'========================================= Métodos Assessores ========================================

Private Property Get wordApp() As Word.Application: Set wordApp = pWordApp: End Property

Private Property Set wordApp(object As Word.Application): Set pWordApp = object: End Property

Private Property Get wordDoc() As Word.Document: Set wordDoc = pWordDoc: End Property

Private Property Set wordDoc(object As Word.Document): Set pWordDoc = object: End Property

Public Property Get templatePath() As String: templatePath = pTemplatePath: End Property

Public Property Let templatePath(value As String): pTemplatePath = value: End Property

'=========================================== Inicialização ===========================================

Private Sub Class_Initialize()
    On Error GoTo ErrHandler
    Set wordApp = New Word.Application
    wordApp.Visible = False
    wordApp.DisplayAlerts = False
    Exit Sub
ErrHandler:
    CloseWordApp
    Err.Raise ERR_WORD_INIT_FAILED, CLASS_NAME, "Erro ao inicializar."
End Sub

Private Sub Class_Terminate()
    CloseWordDoc
    CloseWordApp
End Sub

'=========================================== Métodos Públicos ===========================================

' Descrição: Importa um template Word
' Parâmetros:
'   - wordTemplatePath: caminho completo do arquivo Word
Public Sub ImportTemplate(wordTemplatePath As String)
    If LCase(Right(wordTemplatePath, 5)) <> ".docx" Then
        Err.Raise ERR_FILE_NOT_FOUND, CLASS_NAME, "Arquivo não é um .docx válido: " & wordTemplatePath
    End If
    On Error GoTo ErrHandler
    If Dir(wordTemplatePath) <> "" Then
        Set wordDoc = wordApp.Documents.Open(wordTemplatePath, ReadOnly:=True)
        templatePath = wordTemplatePath
    Else
        Err.Raise ERR_FILE_NOT_FOUND, CLASS_NAME, "Erro ao importar template Word."
    End If
    Exit Sub
ErrHandler:
    CloseWordApp
    Err.Raise ERR_FILE_NOT_FOUND, CLASS_NAME, "Erro ao importar template Word: " & Err.Description
End Sub

' Descrição: Verifica se um template Word está carregado
' Retorno: True se o documento estiver carregado, False caso contrário
Public Function IsTemplateLoaded() As Boolean
    IsTemplateLoaded = Not wordDoc Is Nothing
End Function

' Descrição: Reseta o caminho do template armazenado
Public Sub ResetTemplatePath()
    templatePath = ""
End Sub

' Descrição: Substitui chaves no documento Word com os valores fornecidos
' Parâmetros:
'   - dictKeysValues: dicionário contendo pares chave/valor para substituição
Public Sub ReplaceKeys(dictKeysValues As Scripting.Dictionary)
    Dim key As Variant
    If wordDoc Is Nothing Then
        CloseWordApp
        Err.Raise ERR_WRITE_FILE_FAILED, CLASS_NAME, "O template não foi importado."
    End If
    On Error GoTo ErrHandler
    With wordDoc.Content.Find
        .ClearFormatting ' Remove qualquer formatação do texto a ser buscado (negrito, itálico, fonte, etc.)
        .Replacement.ClearFormatting ' Remove qualquer formatação do texto que será usado como substituição
        .Forward = True ' Define que a busca será feita do início para o fim do documento
        .Wrap = wdFindContinue ' Faz com que a busca continue do início após atingir o final do documento
        .Format = False ' Indica que a busca não deve considerar formatação (estilos visuais)
        .MatchCase = False ' Ignora diferenças entre maiúsculas e minúsculas (ex: "Nome" = "nome")
        .MatchWholeWord = False ' Permite encontrar o texto mesmo se estiver dentro de outra palavra
        .MatchWildcards = False ' Desativa o uso de curingas como * ou ? na busca
        .MatchSoundsLike = False ' Desativa a busca por palavras que soam como a procurada (útil em inglês)
        .MatchAllWordForms = False ' Desativa a busca por variações da palavra (ex: singular/plural, tempos verbais)
        For Each key In dictKeysValues.keys
            If Trim(key) <> "" Then
                .text = key
                .Replacement.text = dictKeysValues(key)
                .Execute Replace:=wdReplaceAll
            End If
        Next key
    End With
    Exit Sub
ErrHandler:
    CloseWordApp
    Err.Raise ERR_WRITE_FILE_FAILED, CLASS_NAME, "Erro ao substituir chaves no template: " & Err.Description
End Sub

' Descrição: Exporta o documento Word como DOCX ou PDF
' Parâmetros:
'   - targetFolderPath: pasta de destino
'   - fileName: nome do arquivo sem extensão
'   - fileType: tipo de arquivo (docx ou pdf)
' Retorno: Caminho do arquivo exportado
Public Function Export(targetFolderPath As String, fileName As String, FileType As FileType) As String
    If wordDoc Is Nothing Then
        Err.Raise ERR_WRITE_FILE_FAILED, CLASS_NAME, "Nenhum documento Word carregado para exportação."
    End If
    If Dir(targetFolderPath, vbDirectory) = "" Then
        Err.Raise ERR_SAVE_FILE_FAILED, CLASS_NAME, "Pasta de destino não encontrada: " & targetFolderPath
    End If
    On Error GoTo ErrHandler
    Select Case FileType
        Case FileTypeDOCX
            wordDoc.SaveAs targetFolderPath & "\" & fileName & ".docx"
                        Export = targetFolderPath & "\" & fileName & ".docx"
        Case FileTypePDF
            wordDoc.ExportAsFixedFormat OutputFileName:=targetFolderPath & "\" & fileName & ".pdf", ExportFormat:=wdExportFormatPDF   'Exporta em PDF. 17 representa PDF
                        Export = targetFolderPath & "\" & fileName & ".pdf"
                Case Else
            Err.Raise ERR_INVALID_FILE_TYPE, CLASS_NAME, "Tipo de arquivo não suportado: " & FileType
    End Select
    Exit Sub
ErrHandler:
    CloseWordApp
    Err.Raise ERR_SAVE_FILE_FAILED, CLASS_NAME, "Falha ao salvar arquivo: " & Err.Description
End Sub

' Descrição: Exporta o documento e encerra a aplicação Word
Public Sub ExportAndFinalize(targetFolderPath As String, fileName As String, FileType As FileType)
    Export targetFolderPath, fileName, FileType
    CloseWordApp
End Sub

' Descrição: Fecha o documento Word sem salvar alterações
Public Sub CloseWordDoc()
    If Not wordDoc Is Nothing Then
        wordDoc.Close SaveChanges:=False
        Set wordDoc = Nothing
    End If
End Sub

' Descrição: Encerra a aplicação Word e libera recursos
Public Sub CloseWordApp()
    CloseWordDoc
    
    If Not wordApp Is Nothing Then
        wordApp.Quit
        Set wordApp = Nothing
    End If
End Sub
