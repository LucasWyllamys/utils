VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsConexaoAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Option Explicit

Private Type ClassType
    ID As String
    statusAcao As String
    StatusGravacao As String
    Protocolo As String
    DataDenuncia As Date
    DataMigracao As Date
    MultaRecisoria As String
    EmailDestinatario As String
    Instalacao As String
    Observacoes As String
    NomePN As String
    ContaContrato As String
    CNPJ As String
    ModTarif As String
    Demanda As String
    Consultor As String
    EmailCopia As String
    Anexo As String
    EnviadoPor As String
    DataEnvio As Date
    CriadoPor As String
    DataCriacao As Date
End Type

Private Self As ClassType

'Métodos acessores (propriedades):--------------------------------------------------------------------------------------------------------------------

Public Property Get ID() As String
    ID = Self.ID
End Property

Public Property Let ID(Value As String)
    Self.ID = Value
End Property

Public Property Get statusAcao() As String
    statusAcao = Self.statusAcao
End Property

Public Property Let statusAcao(Value As String)
    Self.statusAcao = Value
End Property

Public Property Get StatusGravacao() As String
    StatusGravacao = Self.StatusGravacao
End Property

Public Property Let StatusGravacao(Value As String)
    Self.StatusGravacao = Value
End Property

Public Property Get Protocolo() As String
    Protocolo = Self.Protocolo
End Property

Public Property Let Protocolo(Value As String)
    Self.Protocolo = Value
End Property

Public Property Get DataDenuncia() As Date
    DataDenuncia = Self.DataDenuncia
End Property

Public Property Let DataDenuncia(Value As Date)
    Self.DataDenuncia = Value
End Property

Public Property Get DataMigracao() As Date
    DataMigracao = Self.DataMigracao
End Property

Public Property Let DataMigracao(Value As Date)
    Self.DataMigracao = Value
End Property

Public Property Get MultaRecisoria() As String
    MultaRecisoria = Self.MultaRecisoria
End Property

Public Property Let MultaRecisoria(Value As String)
    Self.MultaRecisoria = Value
End Property

Public Property Get EmailDestinatario() As String
    EmailDestinatario = Self.EmailDestinatario
End Property

Public Property Let EmailDestinatario(Value As String)
    Self.EmailDestinatario = Value
End Property

Public Property Get Instalacao() As String
    Instalacao = Self.Instalacao
End Property

Public Property Let Instalacao(Value As String)
    Self.Instalacao = Value
End Property

Public Property Get Observacoes() As String
    Observacoes = Self.Observacoes
End Property

Public Property Let Observacoes(Value As String)
    Self.Observacoes = Value
End Property

Public Property Get NomePN() As String
    NomePN = Self.NomePN
End Property

Public Property Let NomePN(Value As String)
    Self.NomePN = Value
End Property

Public Property Get ContaContrato() As String
    ContaContrato = Self.ContaContrato
End Property

Public Property Let ContaContrato(Value As String)
    Self.ContaContrato = Value
End Property

Public Property Get CNPJ() As String
    CNPJ = Self.CNPJ
End Property

Public Property Let CNPJ(Value As String)
    Self.CNPJ = Value
End Property

Public Property Get ModTarif() As String
    ModTarif = Self.ModTarif
End Property

Public Property Let ModTarif(Value As String)
    Self.ModTarif = Value
End Property

Public Property Get Demanda() As String
    Demanda = Self.Demanda
End Property

Public Property Let Demanda(Value As String)
    Self.Demanda = Value
End Property

Public Property Get Consultor() As String
    Consultor = Self.Consultor
End Property

Public Property Let Consultor(Value As String)
    Self.Consultor = Value
End Property

Public Property Get EmailCopia() As String
    EmailCopia = Self.EmailCopia
End Property

Public Property Let EmailCopia(Value As String)
    Self.EmailCopia = Value
End Property

Public Property Get Anexo() As String
    Anexo = Self.Anexo
End Property

Public Property Let Anexo(Value As String)
    Self.Anexo = Value
End Property

Public Property Get EnviadoPor() As String
    EnviadoPor = Self.EnviadoPor
End Property

Public Property Let EnviadoPor(Value As String)
    Self.EnviadoPor = Value
End Property

Public Property Get DataEnvio() As Date
    DataEnvio = Self.DataEnvio
End Property

Public Property Let DataEnvio(Value As Date)
    Self.DataEnvio = Value
End Property

Public Property Get CriadoPor() As String
    CriadoPor = Self.CriadoPor
End Property

Public Property Let CriadoPor(Value As String)
    Self.CriadoPor = Value
End Property

Public Property Get DataCriacao() As Date
    DataCriacao = Self.DataCriacao
End Property

Public Property Let DataCriacao(Value As Date)
    Self.DataCriacao = Value
End Property

'Métodos:--------------------------------------------------------------------------------------------------------------------------------------
