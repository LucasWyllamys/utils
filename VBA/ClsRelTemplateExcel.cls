VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsRelTemplateExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Esta classe destina-se a realizar a tarefa de gerar um relatório em PDF usando uma planilha de outro arquivo Excel como template.

Private Type ClasseType
    CaminhoArquivo  As String
    Arquivo         As Workbook
    Planilha        As Worksheet
    'Campos:----------------------
    Protocolo       As String
    DataDenuncia    As Date
    DataMigracao    As Date
    Instalacao      As Long
    MultaRecisoria  As String
    NomePN          As String
    ContaContrato   As String
    CNPJ            As String
    ModTarif        As String
    Demanda         As String
    Consultor       As String
    EmailCopia      As String
    '-----------------------------
End Type

Private Self As ClasseType

Private Sub Class_Initialize()  'Evento
    CaminhoArquivo = ThisWorkbook.Path & "\arquivos\modeloCartaRespostaMigracaoACL.xlsx"                 'Caminho do arquivo base
    Arquivo = Workbooks.Open(Filename:=CaminhoArquivo, ReadOnly:=True)                                      'Instancia o arquivo base
    Planilha = Arquivo.Worksheets(1)                                                                        'Instancia a planilha base
End Sub

Private Sub Class_Terminate()   'Evento
    Arquivo.Close False
    Arquivo = Nothing
    Planilha = Nothing
End Sub

'Métodos acessores (propriedades):--------------------------------------------------------------------------------------------------------------------

Private Property Get CaminhoArquivo() As String    'Método Assessor usado para retornar "get: pegar" o valor do atributo "Nome" da classe "ClsPessoa"
    CaminhoArquivo = Self.CaminhoArquivo
End Property

Private Property Let CaminhoArquivo(Value As String)   'Método Assessor usado para modificar "let: deixar" o valor do atributo "Nome" da classe "ClsPessoa"
    Self.CaminhoArquivo = Value
End Property

Private Property Get Arquivo() As Workbook
    Set Arquivo = Self.Arquivo
End Property

Private Property Let Arquivo(Objeto As Workbook)
    Set Self.Arquivo = Objeto
End Property

Private Property Get Planilha() As Worksheet
    Planilha = Self.Planilha
End Property

Private Property Let Planilha(Objeto As Worksheet)
    Set Self.Planilha = Objeto
End Property

Private Property Get Protocolo() As String
    Protocolo = Self.Protocolo
End Property

Private Property Let Protocolo(Value As String)
    Self.Protocolo = Value
End Property

Private Property Get DataDenuncia() As Date
    DataDenuncia = Self.DataDenuncia
End Property

Private Property Let DataDenuncia(Value As Date)
    Self.DataDenuncia = Value
End Property

Private Property Get DataMigracao() As Date
    DataMigracao = Self.DataMigracao
End Property

Private Property Let DataMigracao(Value As Date)
    Self.DataMigracao = Value
End Property

Private Property Get Instalacao() As Long
    Instalacao = Self.Instalacao
End Property

Private Property Let Instalacao(Value As Long)
    Self.Instalacao = Value
End Property

Private Property Get MultaRecisoria() As Long
    MultaRecisoria = Self.MultaRecisoria
End Property

Private Property Let MultaRecisoria(Value As Long)
    Self.MultaRecisoria = Value
End Property

Private Property Get NomePN() As String
    NomePN = Self.NomePN
End Property

Private Property Let NomePN(Value As String)
    Self.NomePN = Value
End Property

Private Property Get ContaContrato() As String
    ContaContrato = Self.ContaContrato
End Property

Private Property Let ContaContrato(Value As String)
    Self.ContaContrato = Value
End Property

Private Property Get CNPJ() As String
    CNPJ = Self.CNPJ
End Property

Private Property Let CNPJ(Value As String)
    Self.CNPJ = Value
End Property

Private Property Get ModTarif() As String
    ModTarif = Self.ModTarif
End Property

Private Property Let ModTarif(Value As String)
    Self.ModTarif = Value
End Property

Private Property Get Demanda() As String
    Demanda = Self.Demanda
End Property

Private Property Let Demanda(Value As String)
    Self.Demanda = Value
End Property

Private Property Get Consultor() As String
    Consultor = Self.Consultor
End Property

Private Property Let Consultor(Value As String)
    Self.Consultor = Value
End Property

Private Property Get EmailCopia() As String
    EmailCopia = Self.EmailCopia
End Property

Private Property Let EmailCopia(Value As String)
    Self.EmailCopia = Value
End Property

'Métodos:--------------------------------------------------------------------------------------------------------------------------------------

Public Sub GerarRelatorio(Linha As Integer)
'Este método tem por objetivo gerar um relatório em PDF usando uma planilha de outro arquivo Excel como template.
'Os dados são passados de uma planilha deste arquivo para outro planilha de outro arquivo, e este é exportado em PDF.
    Dim CaminhoAnexo As String
    'Pega os dados da planilha neste arquivo e atribui às propriedades----------------
    ThisWorkbook.Activate
    With Planilha1
        Protocolo = .Cells(Linha, "D").Value
        DataDenuncia = .Cells(Linha, "E").Value
        DataMigracao = .Cells(Linha, "F").Value
        Instalacao = .Cells(Linha, "H").Value
        MultaRecisoria = .Cells(Linha, "I").Value
        NomePN = .Cells(Linha, "K").Value
        ContaContrato = .Cells(Linha, "L").Value
        CNPJ = .Cells(Linha, "M").Value
        ModTarif = .Cells(Linha, "N").Value
        Demanda = .Cells(Linha, "O").Value
    End With
    'Passa os valores das propriedades para o tamplate em Excel------------------------
    With Self.Planilha
        .Range("N15").Value = Protocolo
        .Range("N13").Value = DataDenuncia
        .Range("N14").Value = DataMigracao
        .Range("E13").Value = Instalacao & " / " & ContaContrato
        .Range("E14").Value = NomePN
        .Range("E15").Value = CNPJ
        .Range("K13").Value = MultaRecisoria
        .Range("K14").Value = ModTarif
        .Range("K15").Value = Demanda
    'Configurações do gráfico-----------------------------------------------------------
        .ChartObjects("Gráfico 52").Activate
        With ActiveChart.Axes(xlValue)
            .Select
            .MinimumScale = CLng(DataDenuncia)
            .MaximumScale = CLng(DataMigracao) + 1
            .MajorUnit = (CLng(DataMigracao) - CLng(DataDenuncia) + 1) / 5
        End With
        Range("A1").Select
    'Exportar em PDF--------------------------------------------------------------------
        CaminhoAnexo = ThisWorkbook.Path & "\anexos\" & UCase(Replace(NomePN, " ", "_")) & "_" & Instalacao & ".pdf"
        .ExportAsFixedFormat Type:=xlTypePDF, Filename:=CaminhoAnexo, Quality:= _
        xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
    End With
    'Salvar o caminho em uma célula------------------------------------------------------
    Planilha1.Cells(Linha, "R").Value = CaminhoAnexo
End Sub
