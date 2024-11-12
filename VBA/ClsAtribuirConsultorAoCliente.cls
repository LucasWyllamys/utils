VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsAtribuirConsultorAoCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type ClasseType
    TipoCliente As String
    Municipio As String
    PN As String
    UTD As String
    Setor As String
    DemandaMax As Double
    NomePN As String
    Consultor As String
    DictDescricoesCNAE As New Scripting.Dictionary
    DictPNsPai As New Scripting.Dictionary
    DictCarteiras As New Scripting.Dictionary
End Type

Private Self As ClasseType

'Private Sub Class_Initialize()
'    With Self.DictCarteiras
'        .Add
'    End With
'"muni_salvador",            "LEA DANTAS DE OLIVEIRA AQUINO"
'"estd",                 "MICHELLE RIBEIRO BRASILEIRO E SILVA"
'"fedr",                 "ADNA GONCALVES DA SILVA FERNANDES"
'"telecomunicacoes",             "MARIA GABRIELA SOUSA CASTRO"
'"industriais_1",            "LUCAS DE ALMEIDA VIRGOLINO"
'"industriais_2",            "LUELE CHALEGRE SANTANA"
'End Sub

Private Property Get TipoCliente() As String
    TipoCliente = Self.TipoCliente
End Property

Private Property Let TipoCliente(Value As String)
    Self.TipoCliente = Value
End Property

Private Property Get Municipio() As String
    Municipio = Self.Municipio
End Property

Private Property Let Municipio(Value As String)
    Self.Municipio = Value
End Property

Private Property Get PN() As String
    PN = Self.PN
End Property

Private Property Let PN(Value As String)
    Self.PN = Value
End Property

Private Property Get UTD() As String
    UTD = Self.UTD
End Property

Private Property Let UTD(Value As String)
    Self.UTD = Value
End Property

Private Property Get Setor() As String
    Setor = Self.Setor
End Property

Private Property Let Setor(Value As String)
    Self.Setor = Value
End Property

Private Property Get DemandaMax() As Double
    DemandaMax = Self.DemandaMax
End Property

Private Property Let DemandaMax(Value As Double)
    Self.DemandaMax = Value
End Property

Private Property Get NomePN() As String
    NomePN = Self.NomePN
End Property

Private Property Let NomePN(Value As String)
    Self.NomePN = Value
End Property

Private Property Get Consultor() As String
    Consultor = Self.Consultor
End Property

Private Property Let Consultor(Value As String)
    Self.Consultor = Value
End Property

Private Property Get DictDescricoesCNAE() As Scripting.Dictionary
    Set DictDescricoesCNAE = Self.DictDescricoesCNAE
End Property

Private Property Set DictDescricoesCNAE(Value As Scripting.Dictionary)
    Set Self.DictDescricoesCNAE = Value
End Property

Private Property Get DictPNsPai() As Scripting.Dictionary
    Set DictPNsPai = Self.DictPNsPai
End Property

Private Property Set DictPNsPai(Value As Scripting.Dictionary)
    Set Self.DictPNsPai = Value
End Property

Private Property Get DictCarteiras() As Scripting.Dictionary
    Set DictCarteiras = Self.DictCarteiras
End Property

Private Property Set DictCarteiras(Value As Scripting.Dictionary)
    Set Self.DictCarteiras = Value
End Property

Private Function ObterValorDictPNsPai(Chave As String) As String
    If Self.DictPNsPai.Exists(Chave) Then
        ObterValorDictPNsPai = Self.DictPNsPai(Chave)
    End If
End Function

Private Sub InserirValorDictPNsPai(Chave As String, Valor As String)
    If Not Self.DictPNsPai.Exists(Chave) Then
        Self.DictPNsPai.Add Chave, Valor
    End If
End Sub

Public Function AtribuirConsultor(ParamTipoCliente As String, ParamMunicipio As String, ParamPN As String, ParamUTD As String, _
ParamSetor As String, ParamDemandaMax As Double, ParamNomePN As String, ParamDictDescricoesCNAE As Object) As String
    
    TipoCliente = ParamTipoCliente
    Municipio = ParamMunicipio
    PN = ParamPN
    UTD = ParamUTD
    Setor = ParamSetor
    DemandaMax = ParamDemandaMax
    NomePN = ParamNomePN
    Set DictDescricoesCNAE = ParamDictDescricoesCNAE
                                                ' 1ª instância (PNs específicos)
    Consultor = ObterValorDictPNsPai(PN)        ' 2ª instância
    If Consultor = "" Then AtribuirPorTipoPN    ' 3ª instância
    If Consultor = "" Then AtribuirPorCNAE      ' 4ª instância
    If Consultor = "" Then AtribuirPorUTD       ' 5ª instância

    AtribuirConsultor = Consultor
End Function

Private Sub AtribuirPorTipoPN()
    If TipoCliente = "MUNI" And Municipio = "SALVADOR" Then
        Consultor = "LEA DANTAS DE OLIVEIRA AQUINO"
    ElseIf TipoCliente = "ESTD" Then
        Consultor = "MICHELLE RIBEIRO BRASILEIRO E SILVA"
    ElseIf TipoCliente = "FEDR" Then
        Consultor = "ADNA GONCALVES DA SILVA FERNANDES"
    End If
End Sub

Private Sub AtribuirPorCNAE()
    If DictDescricoesCNAE("descricao_divisao") = "TELECOMUNICAÇÕES" Then
        Consultor = "MARIA GABRIELA SOUSA CASTRO"
    ElseIf DictDescricoesCNAE("descricao_grupo") = "HOTÉIS E SIMILARES" Then
        If Setor = "SALVADOR" Or Setor = "METROPOLINATA" Then
            If DemandaMax > 300 Then
                Consultor = "LUCAS DE ALMEIDA VIRGOLINO"
                Call InserirValorDictPNsPai(PN, Consultor)
            End If
        End If
    ElseIf DictDescricoesCNAE("descricao_grupo") = "ATIVIDADES DE ATENDIMENTO HOSPITALAR" Then
        If Setor = "SALVADOR" Or Setor = "METROPOLINATA" Then
            Consultor = "LUELE CHALEGRE SANTANA"
            Call InserirValorDictPNsPai(PN, Consultor)
        End If
    ElseIf DictDescricoesCNAE("denominacao_cnae_sap") = "INDUSTRIAL" Then
        If DemandaMax > 300 Then
            If Setor <> "OESTE" Then
                Consultor = "DEFINIR CONSULTOR INDUSTRIA"
                Call InserirValorDictPNsPai(PN, Consultor)
            End If
        Else
            If Setor = "METROPOLITANA" Or Setor = "SALVADOR" Then
                Consultor = "GRACE SOUSA SILVEIRA"
            End If
        End If
    ElseIf DictDescricoesCNAE("denominacao_cnae_sap") = "RURAL" Then
        If DemandaMax > 500 Then
            If Setor <> "OESTE" Then
                Consultor = "THELMO FRANCISCO PEREIRA DA SILVA"
                Call InserirValorDictPNsPai(PN, Consultor)
            End If
        End If
    End If
End Sub

Private Sub AtribuirPorUTD()
' COMERCIAL, RESIDENCIAL, SERVIÇO PÚBLICO, REVENDA, SISTEMA DE MIGRAÇÃO
    If Setor = "SALVADOR" Then
        If DemandaMax > 300 Then
            Consultor = "LUELE CHALEGRE SANTANA"
            Call InserirValorDictPNsPai(PN, Consultor)
        Else
            If UTD = "GRACA" Or UTD = "PITUBA" Then
                Consultor = "ITALO BISPO DE ALMEIDA"
            ElseIf UTD = "ITAPOAN" Or UTD = "PERIPERI" Or UTD = "PIRAJA" Then
                Consultor = "SILVANA MATIAS ATAIDE"
            End If
        End If
    ElseIf Setor = "METROPOLITANA" Then
        If DemandaMax > 300 Then
            Consultor = "LUELE CHALEGRE SANTANA"
            Call InserirValorDictPNsPai(PN, Consultor)
        Else
            Consultor = "SILVANA MATIAS ATAIDE"
        End If
    ElseIf UTD = "ALAGOINHAS" Or UTD = "ESPLANADA" Then
        If DemandaMax > 300 Then
            Consultor = "LUCAS DE ALMEIDA VIRGOLINO"
        Else
            Consultor = "GRACE SOUSA SILVEIRA"
        End If
    ElseIf UTD = "IRECE" Or UTD = "ITABERABA" Or UTD = "SEABRA" Then
        Consultor = "ANALISTA IRC"
    ElseIf UTD = "AMARGOSA" Or UTD = "IPIAU" Or UTD = "SANTO ANTONIO DE JESUS" Or UTD = "VALENCA" Then
        Consultor = "ANDERSON DA CRUZ OLIVEIRA"
    ElseIf UTD = "PAULO AFONSO" Or UTD = "JACOBINA" Or UTD = "JUAZEIRO" Or UTD = "REMANSO" Or UTD = "SENHOR DO BONFIM" Then
        Consultor = "EDUARDO VINICIUS DA HORA CEDRAZ"
    ElseIf UTD = "EUNAPOLIS" Or UTD = "ILHEUS" Or UTD = "ITABUNA" Then
        Consultor = "GEORGE MENEZES B NASCIMENTO"
    ElseIf UTD = "BARREIRAS" Or UTD = "BOM JESUS DA LAPA" Or UTD = "IBOTIRAMA" Or UTD = "LUIS EDUARDO MAGALHAES" Or UTD = "SANTA MARIA DA VITORIA" Then
        Consultor = "JESSICA SANTOS SILVA / ENGENHEIRO BRR"
    ElseIf UTD = "CONCEICAO DO COITE" Or UTD = "FEIRA DE SANTANA NORTE" Or UTD = "FEIRA DE SANTANA SUL" Or UTD = "SANTO AMARO" Or UTD = "SERRINHA" Or UTD = "RIBEIRA DO POMBAL" Then
        Consultor = "JOAO MATEUS SANTOS DIAS"
    ElseIf UTD = "JEQUIE" Or UTD = "LIVRAMENTO DE NOSSA SENHORA" Or UTD = "BRUMADO" Or UTD = "GUANAMBI" Or UTD = "ITAPETINGA" Or UTD = "VITORIA DA CONQUISTA" Then
        Consultor = "LUIZ GUSTAVO CALDEIRA DE CASTRO"
    ElseIf UTD = "PORTO SEGURO" Or UTD = "POSTO DA MATA" Or UTD = "TEIXEIRA DE FREITAS" Then
        Consultor = "MARCELO FERRAZ BOMFIM"
    End If
End Sub
