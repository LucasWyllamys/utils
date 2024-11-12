VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsCompilador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Esta classe destina-se a realizar a tarefa de passar informações de uma planilha de outro arquivo para as-
'  propriedades desta classe, e por conseguinte, para uma planilha deste arquivo usando uma chave.
    
Private Type ClassType
    CaminhoArquivo  As String
    Arquivo         As Workbook
    Planilha        As Worksheet
    'Campos:----------------------
    Instalacao      As Long
    NomePN          As String
    ContaContrato   As String
    CNPJ            As String
    ModTarif        As String
    Demanda         As String
    Consultor       As String
    EmailCopia      As String
    '-----------------------------
End Type

Private Self As ClassType

Private Sub Class_Initialize()  'Evento
    CaminhoArquivo = ThisWorkbook.Path & "\arquivos\baseClientesCGR.xlsx"  'Caminho do arquivo base
    Arquivo = Workbooks.Open(Filename:=CaminhoArquivo, ReadOnly:=True)                                      'Instancia o arquivo base
    Planilha = Arquivo.Worksheets(1)                                                                        'Instancia a planilha base
End Sub

Private Sub Class_Terminate()   'Evento
    Arquivo.Close False
    Arquivo = Nothing
    Planilha = Nothing
End Sub

' Métodos acessores (propriedades):--------------------------------------------------------------------------------------------------------------------

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
    Set Planilha = Self.Planilha
End Property

Private Property Let Planilha(Objeto As Worksheet)
    Set Self.Planilha = Objeto
End Property

Private Property Get Instalacao() As Long
    Instalacao = Self.Instalacao
End Property

Private Property Let Instalacao(Value As Long)
    Self.Instalacao = Value
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
    Self.CNPJ = Format(Value, "00\.000\.000\/0000\-00")
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

' Métodos:--------------------------------------------------------------------------------------------------------------------------------------

Public Sub Compilar(Chave As Long, LinhaCelulaDestino As Integer)
'Este método tem por objetivo passar as informações das células da planilha do arquivo instanciado para as propriedades-
'desta classe e das propriedades desta classe para as células definidas.
    Dim matrizTabela As String
    
    matrizTabela = "A:I"
    Instalacao = Chave
    
    On Error Resume Next
    With Application.WorksheetFunction
        Instalacao = .VLookup(Instalacao, Self.Planilha.Columns(matrizTabela), 4, False)
        NomePN = .VLookup(Instalacao, Self.Planilha.Columns(matrizTabela), 4, False)
        ContaContrato = .VLookup(Instalacao, Self.Planilha.Columns(matrizTabela), 3, False)
        CNPJ = .VLookup(Instalacao, Self.Planilha.Columns(matrizTabela), 5, False)
        ModTarif = .VLookup(Instalacao, Self.Planilha.Columns(matrizTabela), 6, False)
        Demanda = .VLookup(Instalacao, Self.Planilha.Columns(matrizTabela), 7, False)
        Consultor = .VLookup(Instalacao, Self.Planilha.Columns(matrizTabela), 8, False)
        EmailCopia = .VLookup(Instalacao, Self.Planilha.Columns(matrizTabela), 9, False)
    End With
    
    ThisWorkbook.Activate
    With Planilha1
        .Cells(LinhaCelulaDestino, "K").Value = NomePN
        .Cells(LinhaCelulaDestino, "L").Value = ContaContrato
        .Cells(LinhaCelulaDestino, "M").Value = CNPJ
        .Cells(LinhaCelulaDestino, "N").Value = ModTarif
        .Cells(LinhaCelulaDestino, "O").Value = Demanda
        .Cells(LinhaCelulaDestino, "P").Value = Consultor
        .Cells(LinhaCelulaDestino, "Q").Value = EmailCopia
    End With
End Sub
