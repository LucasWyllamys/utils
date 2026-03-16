VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelTemplateProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================================================================================
' Classe VBA: ExcelTemplateProcessor - Processador de Templates Excel
' Versão: 1.0.0
' Autor: Lucas Wyllamys Carmo da Silva
' Criado em: 19/11/2025
' Atualizado em: 19/11/2025
'
' Descrição:
'   Esta classe encapsula a lógica de manipulação de planilhas Excel baseadas em templates (.xlsx),
'   permitindo a importação, inserção de valores em células específicas e exportação em formatos XLSX ou PDF.
'
' Recursos:
'   - Inicialização automática com desativação de atualizações visuais para melhor desempenho.
'   - Importação de arquivos .xlsx como templates.
'   - Inserção de valores em células definidas por meio de um dicionário chave/valor.
'   - Exportação do documento em formato XLSX ou PDF.
'   - Encerramento seguro do workbook e restauração das configurações do Excel.
'
' Pré-requisitos:
'   - Ativar bibliotecas:
'       - Microsoft Excel xx.x Object Library: Ferramentas > Referências > Microsoft Excel xx.x Object Library
'       - Microsoft Scripting Runtime: Ferramentas > Referências > Microsoft Scripting Runtime
'   - Importar módulo: PublicConstants (para códigos de erro como ERR_FILE_NOT_FOUND, ERR_WRITE_FILE_FAILED, etc.)
'
' Observações:
'   - A propriedade `wb` representa a instância do Workbook carregado.
'   - A propriedade `ws` representa a planilha ativa do template.
'   - A propriedade `templatePath` armazena o caminho do template importado.
'   - A enumeração `FileType1` define os formatos de exportação disponíveis: XLSX ou PDF.
'
' Exemplo de uso:
'   Dim excelProcessor As clsExcelTemplateProcessor
'   Dim dictValues As Scripting.Dictionary
'
'   Set excelProcessor = New clsExcelTemplateProcessor
'
'   excelProcessor.ImportTemplate "C:\Templates\Relatorio.xlsx", "Planilha1"
'
'   Set dictValues = New Scripting.Dictionary
'   dictValues.Add "A1", "Lucas Wyllamys"
'   dictValues.Add "B2", "19/11/2025"
'
'   excelProcessor.InsertValues dictValues
'   excelProcessor.Export "C:\Relatorios", "Relatorio_Lucas", FileTypePDF
'   excelProcessor.CloseWb
'===========================================================================================================

Option Explicit

Implements ITemplateProcessor

Private Const CLASS_NAME As String = "ExcelTemplateProcessor"

Private pWb As Workbook

'========================================= Métodos Assessores ========================================

Private Property Get wb() As Workbook
    Set wb = pWb
End Property

Private Property Set wb(app As Workbook)
    Set pWb = app
End Property

'============================================= Ciclo de Vida =============================================

Private Sub Class_Terminate()
    ITemplateProcessor_Finish
End Sub

'============================= Implementação da Interface ======================

' Descrição: Importa um template Excel e define a planilha ativa.
' Parâmetros:
'   - excelTemplatePath: caminho completo do arquivo Excel.
Private Sub ITemplateProcessor_ImportTemplate(ByVal excelTemplatePath As String)
    Dim ext As String
    
    ' Valida se o arquivo existe
    If Not FolderAndFileManipulator.FileExists(excelTemplatePath) Then
        Err.Raise ERR_FILE_NOT_FOUND, CLASS_NAME & ".ImportTemplate", "Arquivo de template não existe."
    End If
    
    ' Obtém e valida a extensão
    ext = FolderAndFileManipulator.GetFileExtension(excelTemplatePath)
    Select Case ext
        Case "xlsx", "xls"
            ' Abre o template em modo somente leitura
            Set wb = Workbooks.Open(fileName:=excelTemplatePath, ReadOnly:=True)

            ' Caso excepcional: arquivo não abriu
            If wb Is Nothing Then
                Err.Raise ERR_OPEN_FILE_FAILED, CLASS_NAME & ".ImportTemplate", "Falha ao carregar o template."
            End If

        Case Else
            ' Extensões inválidas
            Err.Raise ERR_INVALID_FILE_TYPE, CLASS_NAME & ".ImportTemplate", _
                      "Formato inválido: apenas .xlsx ou .xls são suportados."
    End Select
End Sub

' Descrição: Exporta o documento Excel como .xlsx ou .pdf.
' Parâmetros:
'   - targetFolderPath: pasta de destino.
'   - fileName: nome do arquivo sem extensão.
'   - FileType1: tipo de arquivo (xlsx ou pdf).
' Retorno: String - Caminho completo do arquivo exportado.
Private Function ITemplateProcessor_Export( _
    ByVal targetFolderPath As String, _
    ByVal fileName As String, _
    ByVal fileType As ExportFileType, _
    Optional ByVal keysValues As Scripting.Dictionary = Nothing) As String
    
    Dim outPath As String
    Dim safeName As String
    
    ' Verifica o tipo de arquivo de exportação
    If Not (fileType = flXLSX Or fileType = flPDF) Then
        ITemplateProcessor_ResetTemplate
        Err.Raise ERR_INVALID_FILE_TYPE, CLASS_NAME & ".Export", "Tipo de exportação inválido. Tipos de arquivos suportados .xlsx ou .pdf"
    End If

    ' Verifica pasta destino
    If Not FolderAndFileManipulator.FolderExists(targetFolderPath) Then
        ITemplateProcessor_ResetTemplate
        Err.Raise ERR_FOLDER_NOT_FOUND, CLASS_NAME & ".Export", "A pasta destino não existe!"
    End If

    ' Verifica se existe documento carregado
    If wb Is Nothing Then
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
        ReplaceKeys keysValues
    End If
    
    On Error GoTo ErrHandler
    
  ' Decide o formato de saída
    Select Case fileType

        ' Exportação como xlsx
        Case flXLSX
            outPath = FolderAndFileManipulator.BuildPath(targetFolderPath, safeName, "xlsx") ' Monta o caminho do arquivo que será exportado
            FolderAndFileManipulator.EnsureOverwrite outPath ' Se existir, remove arquivo antes de salvar/exportar
            wb.SaveAs fileName:=outPath, FileFormat:=xlOpenXMLWorkbook ' Exporta o arquivo

        ' Exportação como pdf
        Case flPDF
            outPath = FolderAndFileManipulator.BuildPath(targetFolderPath, safeName, "pdf")
            FolderAndFileManipulator.EnsureOverwrite outPath
            wb.ExportAsFixedFormat _
                Type:=xlTypePDF, _
                fileName:=outPath, _
                Quality:=xlQualityStandard, _
                IncludeDocProperties:=True, _
                IgnorePrintAreas:=False, _
                OpenAfterPublish:=False
            
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

' Descrição: Reseta o caminho do template armazenado.
Private Sub ITemplateProcessor_ResetTemplate()
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
End Sub

' Descrição: Fecha o documento Excel sem salvar alterações.
Private Sub ITemplateProcessor_Finish()
    ITemplateProcessor_ResetTemplate
    Set wb = Nothing
End Sub

'============================= Helpers privados =============================

' Descrição: ...
' Parâmetros:
'   - keysValues: dicionário contendo pares celula/valor para substituição.
Public Sub ReplaceKeys(ByVal keysValues As Scripting.Dictionary)
    Dim key As Variant
    
    On Error GoTo ErrHandler

    wb.Activate

    For Each key In keysValues.keys
        Cells.Replace _
            What:=CStr(key), _
            Replacement:=keysValues(key), _
            LookAt:=xlPart, _
            SearchOrder:=xlByRows, _
            MatchCase:=True, _
            SearchFormat:=False, _
            ReplaceFormat:=False, _
            FormulaVersion:=xlReplaceFormula2
    Next key
    
    Exit Sub
    
ErrHandler:
    ITemplateProcessor_ResetTemplate
    Err.Raise ERR_WRITE_FILE_FAILED, CLASS_NAME & ".ReplaceKeys", _
              "Erro ao substituir chaves no template: " & Err.Description
End Sub
