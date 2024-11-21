Attribute VB_Name = "APIs"
Option Explicit

' Importar M�dulo: JsonConverter (https://github.com/VBA-tools/VBA-JSON)
' Importar M�dulo: Utilidades
' Habilitar Biblioteca: Microsoft Scripting Runtime
' Habilitar Biblioteca: Microsoft WinHTTP Services

Public Function ConsultaAPI(url As String, criterio As String) As Scripting.Dictionary
    ' Esta fun��o retorna um dicion�rio.
    ' Use um dicion�rio para armazenar o retorno desta fun��o.
    Dim request As New WinHttpRequest, dictResponse As New Scripting.Dictionary
    Dim responseJSON As String
    
    'request.SetProxy 2, "10.219.78.3:8080", ""                                      ' Servidor padr�o do proxy
    url = url & "/" & criterio
    request.Open "GET", url, False                                                  ' Abre uma conex�o com a API
    request.SetRequestHeader "Accept", "application/json"
    'request.SetRequestHeader "Content-Type", "application/json; charset=UTF-8"      ' Trata a requisi��o como UTF-8
    request.Send                                                                    ' Envia a requisi��o para a API

    If request.status = 200 Then
        responseJSON = request.responseText
        If responseJSON <> "[]" Then
            responseJSON = Utils.LerRespostaComoUTF8(request)                  ' Converte a requisi��o em UTF-8
            Set dictResponse = JsonConverter.ParseJson(responseJSON)                ' Converte o JSON em dicion�rio
            Set ConsultaAPI = dictResponse
        Else
            MsgBox "A consulta n�o retornou dados.", vbInformation
        End If
    Else
        MsgBox "Erro na requisi��o: " & request.status
    End If
End Function

Public Function ConsultaAPICNPJ(cnpj As String) As Scripting.Dictionary
    Dim url As String
    
    url = "https://brasilapi.com.br/api/cnpj/v1/"
    
    Set ConsultaAPICNPJ = ConsultaAPI(url, cnpj)
End Function

Public Function ConsultaAPICPF(cpf As String, dataNascimento As String) As Scripting.Dictionary
    Dim url As String, token As String
    
    token = "AceldjFJujoDwoB0O16bw4GY1kgaidWuvqaRnrWc"

    url = "https://api.infosimples.com/api/v2/consultas/receita-federal/cpf?" & _
    "cpf=" & cpf & _
    "&birthdate=" & dataNascimento & _
    "&origem=web&token=" & token & _
    "&timeout=300"
    
    Set ConsultaAPICPF = ConsultaAPI(url, cpf)
End Function

Public Function ConsultaAPICEP(cep As String) As Scripting.Dictionary
    Dim url As String
    
    url = "https://brasilapi.com.br/api/cep/v1/"
    
    Set ConsultaAPICEP = ConsultaAPI(url, cep)
End Function
