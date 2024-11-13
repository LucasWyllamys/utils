VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsAPIs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Esta classe tem por objetivo realizar operações com APIs.
' Importar Módulo: JsonConverter (https://github.com/VBA-tools/VBA-JSON)
' Importar Módulo: Utilidades
' Habilitar Biblioteca: Microsoft Scripting Runtime
' Habilitar Biblioteca: Microsoft WinHTTP Services

Private Type ClassType
    request             As New WinHttpRequest
    responseJSON        As String
    DictResponse        As New Scripting.Dictionary
    
    url                 As String
    cpf                 As String
    dataNascimento      As Date
    token               As String
    cnpj                As String
    cep                 As String
End Type

Private self As ClassType

' Métodos assessores:-----------------------------------------------------------------------------------
Private Property Get request() As WinHttpRequest
    Set request = self.request
End Property

Private Property Set request(obj As WinHttpRequest)
    Set self.request = obj
End Property

Private Property Get responseJSON() As String
    responseJSON = self.responseJSON
End Property

Private Property Let responseJSON(value As String)
    self.responseJSON = value
End Property

Private Property Get DictResponse() As Scripting.Dictionary
    Set DictResponse = self.DictResponse
End Property

Private Property Set DictResponse(obj As Scripting.Dictionary)
    Set self.DictResponse = obj
End Property

Private Property Get url() As String
    url = self.url
End Property

Private Property Let url(value As String)
    self.url = value
End Property

Private Property Get cpf() As String
    cpf = self.cpf
End Property

Private Property Let cpf(value As String)
    self.cpf = value
End Property

Private Property Get dataNascimento() As Date
    dataNascimento = self.dataNascimento
End Property

Private Property Let dataNascimento(value As Date)
    self.dataNascimento = value
End Property

Private Property Get cnpj() As String
    cnpj = self.cnpj
End Property

Private Property Let cnpj(value As String)
    self.cnpj = value
End Property

Private Property Get cep() As String
    cep = self.cep
End Property

Private Property Let cep(value As String)
    self.cep = value
End Property

Private Property Get token() As String
    token = self.token
End Property

Private Property Let token(value As String)
    self.token = value
End Property
'-------------------------------------------------------------------------------------------------------

Public Function ConsultaAPI(paramURL As String, paramCriterio As String) As Scripting.Dictionary
    ' Esta função retorna um dicionário.
    
    url = paramURL & "/" & paramCriterio
    ' request.SetProxy 2, "10.219.78.3:8080", ""                                      ' Servidor padrão do proxy
    request.Open "GET", url, False                                                  ' Abre uma conexão com a API
    request.SetRequestHeader "Accept", "application/json"
    ' request.SetRequestHeader "Content-Type", "application/json; charset=UTF-8"      ' Trata a requisição como UTF-8
    request.Send                                                                    ' Envia a requisição para a API

    If request.status = 200 Then
        responseJSON = request.responseText
        If responseJSON <> "[]" Then
            responseJSON = Utilidades.LerRespostaComoUTF8(request)                  ' Converte a requisição em UTF-8
            Set DictResponse = JsonConverter.ParseJson(responseJSON)                ' Converte o JSON em dicionário
            Set ConsultaAPI = DictResponse  ' Retorna um dicionário
        Else
            MsgBox "A consulta não retornou dados.", vbInformation
        End If
    Else
        MsgBox "Erro na requisição: " & request.status
    End If
End Function

Public Function ConsultaAPICNPJ(paramCNPJ As String) As Scripting.Dictionary
    Dim url As String
    
    cnpj = paramCNPJ
    url = "https://brasilapi.com.br/api/cnpj/v1/"
    
    Set ConsultaAPICNPJ = ConsultaAPI(url, cnpj)
End Function

Public Function ConsultaAPICPF(paramCPF As String, paramDataNascimento As Date) As Scripting.Dictionary
    Dim url As String, token As String
    
    token = "AceldjFJujoDwoB0O16bw4GY1kgaidWuvqaRnrWc"
    cpf = paramCPF
    dataNascimento = paramDataNascimento
    
    url = "https://api.infosimples.com/api/v2/consultas/receita-federal/cpf?" & _
    "cpf=" & cpf & _
    "&birthdate=" & data_nascimento & _
    "&origem=web&token=" & token & _
    "&timeout=300"
    
    Set ConsultaAPICPF = ConsultaAPI(url, cpf)
End Function

Public Function ConsultaAPICEP(paramCEP As String) As Scripting.Dictionary
    Dim url As String
    
    cep = paramCEP
    url = "https://brasilapi.com.br/api/cep/v1/"
    
    Set ConsultaAPICEP = ConsultaAPI(url, cep)
End Function
