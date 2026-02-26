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
'   - A enumeração `FileType` define os formatos de exportação disponíveis: XLSX ou PDF.
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

'====================================== Propriedades Privadas ========================================

Private pWb As Workbook
Private pWs As Worksheet
Private pTemplatePath As String

'====================================== Constantes =========================================

Private Const CLASS_NAME As String = "clsExcelTemplateProcessor"

'=============================================== Enum ===============================================

Public Enum FileType
    FileTypePDF = 1
    FileTypeXLSX = 2
End Enum

'========================================= Métodos Assessores ========================================

Private Property Get wb() As Workbook
    Set wb = pWb
End Property

Private Property Set wb(object As Workbook)
    Set pWb = object
End Property

Public Property Get ws() As Worksheet
    Set ws = pWs
End Property

Private Property Set ws(object As Worksheet)
    Set pWs = object
End Property

Public Property Get templatePath() As String
    templatePath = pTemplatePath
End Property

Public Property Let templatePath(Value As String)
    pTemplatePath = Value
End Property

'============================================= Inicialização =============================================

Private Sub Class_Terminate()
    CloseWb
    EnableExcelUpdates
End Sub

'=========================================== Métodos Públicos ===========================================

' Descrição: Importa um template Excel e define a planilha ativa.
' Parâmetros:
'   - excelTemplatePath: caminho completo do arquivo Excel.
'   - templateSheetName: nome da planilha dentro do template.
Public Sub ImportTemplate(excelTemplatePath As String, templateSheetName As String)
    On Error GoTo ErrHandler
    DisableExcelUpdates
    If Dir(excelTemplatePath) = "" Then
        Err.Raise ERR_FILE_NOT_FOUND, CLASS_NAME, "Arquivo não encontrado: " & excelTemplatePath
    End If
    templatePath = excelTemplatePath
    Set wb = Workbooks.Open(fileName:=templatePath, ReadOnly:=True)
    Set ws = wb.Worksheets(templateSheetName)
    EnableExcelUpdates
    Exit Sub
ErrHandler:
    EnableExcelUpdates
    CloseWb
    Err.Raise ERR_EXCEL_INIT_FAILED, CLASS_NAME, "Erro ao inicializar: " & Err.Description
End Sub

' Descrição: Verifica se um template Excel está carregado.
' Retorno: True se o documento estiver carregado, False caso contrário.
Public Function IsTemplateLoaded() As Boolean
    IsTemplateLoaded = Not ws Is Nothing
End Function

' Descrição: Reseta o caminho do template armazenado.
Public Sub ResetTemplatePath()
    templatePath = ""
End Sub

' Descrição: Insere os valores fornecidos nas células especificadas.
' Parâmetros:
'   - dictRangesValues: dicionário contendo pares celula/valor para substituição.
Public Sub InsertValues(dictRangesValues As Scripting.Dictionary)
    Dim key As Variant
    If ws Is Nothing Then
        CloseWb
        Err.Raise ERR_WRITE_FILE_FAILED, CLASS_NAME, "Template não importado."
    End If
    On Error GoTo ErrHandler
    For Each key In dictRangesValues.keys
        ws.Range(CStr(key)).Value = dictRangesValues(key)
    Next key
    Exit Sub
ErrHandler:
    EnableExcelUpdates
    CloseWb
    Err.Raise ERR_WRITE_FILE_FAILED, CLASS_NAME, "Erro ao inserir valores: " & Err.Description
End Sub

' Descrição: Exporta o documento Excel como .xlsx ou .pdf.
' Parâmetros:
'   - targetFolderPath: pasta de destino.
'   - fileName: nome do arquivo sem extensão.
'   - fileType: tipo de arquivo (xlsx ou pdf).
' Retorno: String - Caminho completo do arquivo exportado.
Public Function Export(targetFolderPath As String, fileName As String, FileType As FileType) As String
    If ws Is Nothing Then
        Err.Raise ERR_WRITE_FILE_FAILED, CLASS_NAME, "Nenhum documento carregado para exportação."
    End If
    If Dir(targetFolderPath, vbDirectory) = "" Then
        Err.Raise ERR_SAVE_FILE_FAILED, CLASS_NAME, "Pasta de destino não encontrada: " & targetFolderPath
    End If
    On Error GoTo ErrHandler
    
    Dim fullPath As String
    fullPath = targetFolderPath & Application.PathSeparator & fileName
    
    Select Case FileType
        Case FileTypePDF
            ws.ExportAsFixedFormat Type:=xlTypePDF, fileName:=fullPath & ".pdf"
            Export = fullPath & ".pdf"
        Case FileTypeXLSX
            wb.SaveAs fileName:=fullPath & ".xlsx", FileFormat:=xlOpenXMLWorkbook
            Export = fullPath & ".xlsx"
        Case Else
            Err.Raise ERR_INVALID_FILE_TYPE, CLASS_NAME, "Tipo de arquivo inválido."
    End Select
    Exit Function
ErrHandler:
    EnableExcelUpdates
    CloseWb
    Err.Raise ERR_SAVE_FILE_FAILED, CLASS_NAME, "Falha ao salvar arquivo: " & Err.Description
End Function

' Descrição: Exporta o documento e encerra a aplicação Excel.
' Parâmetros:
'   - targetFolderPath: pasta de destino.
'   - fileName: nome do arquivo sem extensão.
'   - fileType: tipo de arquivo (xlsx ou pdf).
Public Sub ExportAndFinalize(targetFolderPath As String, fileName As String, FileType As FileType)
    Export targetFolderPath, fileName, FileType
    CloseWb
End Sub

' Descrição: Fecha o documento Excel sem salvar alterações.
Public Sub CloseWb()
    On Error Resume Next
    Set ws = Nothing
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Set wb = Nothing
    EnableExcelUpdates
End Sub

'=========================================== Métodos Privados ===========================================

' Descrição: Desativa atualizações visuais e cálculos automáticos do Excel para melhorar desempenho.
Private Sub DisableExcelUpdates()
    Application.ScreenUpdating = False
    'Application.Calculation = xlCalculationManual
End Sub

' Descrição: Reativa atualizações visuais e cálculos automáticos do Excel.
Private Sub EnableExcelUpdates()
    Application.ScreenUpdating = True
    'Application.Calculation = xlCalculationAutomatic
End Sub
