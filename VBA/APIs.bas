Attribute VB_Name = "APIs"
Option Explicit

' Importar M�dulo: JsonConverter (https://github.com/VBA-tools/VBA-JSON)
' Importar M�dulo: Utilidades
' Habilitar Biblioteca: Microsoft Scripting Runtime
' Habilitar Biblioteca: Microsoft WinHTTP Services

Public Function urlAPIs(palavraChave As String) As String
    ' API para consultar CNAE: https://servicodados.ibge.gov.br/api/v2/cnae/classes
    ' API para consultar CNPJ: https://brasilapi.com.br/api/cnpj/v1/{" & cnpj & "}
    ' API para consultar CEP: https://brasilapi.com.br/api/cep/v1/{" & cep & "}
    ' API para consultar CPF:
    '    token = "AceldjFJujoDwoB0O16bw4GY1kgaidWuvqaRnrWc"
    '    url = "https://api.infosimples.com/api/v2/consultas/receita-federal/cpf?" & _
    '    "cpf=" & cpf & _
    '    "&birthdate=" & data_nascimento & _
    '    "&origem=web&token=" & token & _
    '    "&timeout=300"
    If palavraChave = "cnae" Then
        urlAPIs
    ElseIf palavraChave = "cnpj" Then
        urlAPIs
    ElseIf palavraChave = "cep" Then
        urlAPIs
    ElseIf palavraChave = "cpf" Then
        urlAPIs
    End If
End Function

Public Function ConsultaAPI(url As String, criterio As String) As Scripting.Dictionary
    ' Esta fun��o retorna um dicion�rio.
    ' Use um dicion�rio para armazenar o retorno desta fun��o.
    
    Dim request As New WinHttpRequest, dictResponse As Object
    Dim responseJSON As String
    
    Set dictResponse = CreateObject("Scripting.Dictionary")
    
    'request.SetProxy 2, "10.219.78.3:8080", ""                                      ' Servidor padr�o do proxy
    url = url & "/" & criterio
    request.Open "GET", url, False                                                  ' Abre uma conex�o com a API
    request.SetRequestHeader "Accept", "application/json"
    'request.SetRequestHeader "Content-Type", "application/json; charset=UTF-8"      ' Trata a requisi��o como UTF-8
    request.Send                                                                    ' Envia a requisi��o para a API

    If request.status = 200 Then
        responseJSON = request.responseText
        If responseJSON <> "[]" Then
            responseJSON = Utilidades.LerRespostaComoUTF8(request)                  ' Converte a requisi��o em UTF-8
            Set dictResponse = JsonConverter.ParseJson(responseJSON)                ' Converte o JSON em dicion�rio
            Set ConsultaAPI = dictResponse
        Else
            MsgBox "A consulta n�o retornou dados.", vbInformation
        End If
    Else
        MsgBox "Erro na requisi��o: " & request.status
    End If
    
    Set request = Nothing
    Set dictResponse = Nothing
End Function

Public Function APIConsultaCNPJ(cnpj As String)
    Dim url As String, xmlhttp As New WinHttpRequest, Resposta As Object
    'Constr�i a URL da API com os par�metros desejados
    url = "https://brasilapi.com.br/api/cnpj/v1/{" & cnpj & "}"
    '-------------------------------------------------
    xmlhttp.Open "GET", url, False      'Abre uma conex�o com a API
    xmlhttp.Send                        'Envia a solicita��o para a API
    If xmlhttp.status <> 200 Then    'Tratar de erros
        APIConsultaCNPJ = xmlhttp.status   'Mensagem de erro
    Else
        'Converter o JSON
        Set Resposta = JsonConverter.ParseJson(xmlhttp.responseText)
        
        With Form_STGMusic
            .Txt_RazaoSocial_0101.Value = Resposta("razao_social")
            .Txt_NomeFantasia_0101.Value = Resposta("nome_fantasia")
            .Txt_NaturezaJuridica_0101.Value = Resposta("codigo_natureza_juridica")
            .Txt_AtividadePrincipal_0101.Value = Resposta("cnae_fiscal") & " - " & Resposta("cnae_fiscal_descricao")
            .CmbBox_SituacaoCadCNPJ_0101.Value = Resposta("descricao_situacao_cadastral")
            .Txt_RuaCNPJ_0101.Value = Resposta("descricao_tipo_logradouro") & " " & Resposta("logradouro")
            .Txt_NumCNPJ_0101.Value = Resposta("numero")
            .Txt_BairroCNPJ_0101.Value = Resposta("bairro")
            .Txt_CompCNPJ_0101.Value = Resposta("complemento")
            .Txt_CidadeCNPJ_0101.Value = Resposta("municipio")
            .Txt_EstadoCNPJ_0101.Value = Resposta("uf")
            .Txt_CEPCNPJ_0101.Value = Resposta("cep")
        End With
        APIConsultaCNPJ = xmlhttp.status   'Mensagem de erro
    End If
    'Debug.Print xmlhttp.responseText    'Exibe a resposta da API
    Set xmlhttp = Nothing
End Function

Public Function APIConsultaCPF(cpf As String, data_nascimento As String) As String
    Dim url As String, xmlhttp As New WinHttpRequest, token As String, Resposta As Object, data As Object
    On Error Resume Next
    token = "AceldjFJujoDwoB0O16bw4GY1kgaidWuvqaRnrWc"
    'Constr�i a URL da API com os par�metros desejados
    url = "https://api.infosimples.com/api/v2/consultas/receita-federal/cpf?" & _
    "cpf=" & cpf & _
    "&birthdate=" & data_nascimento & _
    "&origem=web&token=" & token & _
    "&timeout=300"
    '-------------------------------------------------
    xmlhttp.Open "GET", url, False      'Abre uma conex�o com a API
    xmlhttp.Send                        'Envia a solicita��o para a API
    If xmlhttp.status <> 200 Then    'Tratar de erros
        APIConsultaCPF = xmlhttp.status   'Mensagem de erro
    Else
        'Converter o JSON
        Set Resposta = JsonConverter.ParseJson(xmlhttp.responseText)
        
        With Form_STGMusic
            Set data = Resposta("data")(1)
            .Txt_Nome_0101.Value = data("nome")
            .CmbBox_SituacaoCadCPF_0101.Value = data("situacao_cadastral")
            .CmdBtn_ConfirmacaoAutencidade_0101.Caption = data("site_receipt")
        End With
        APIConsultaCPF = xmlhttp.status   'Mensagem de erro
    End If
    'Debug.Print xmlhttp.responseText    'Exibe a resposta da API
    Set xmlhttp = Nothing               'Limpa o objeto XMLHttpRequest
    On Error GoTo 0
End Function

Public Function APIConsultaCEP(cep As String)
    Dim url As String, xmlhttp As New WinHttpRequest, Resposta As Object
    'Constr�i a URL da API com os par�metros desejados
    On Error Resume Next
    url = "https://brasilapi.com.br/api/cep/v1/{" & cep & "}"
    '-------------------------------------------------
    xmlhttp.Open "GET", url, False      'Abre uma conex�o com a API
    xmlhttp.Send                        'Envia a solicita��o para a API
    If xmlhttp.status <> 200 Then    'Tratar de erros
        APIConsultaCEP = xmlhttp.status   'Mensagem de erro
    Else
        'Converter o JSON
        Set Resposta = JsonConverter.ParseJson(xmlhttp.responseText)
        
        With Form_STGMusic
            .Txt_Rua_0101.Value = Resposta("street")
            .Txt_Bairro_0101.Value = Resposta("neighborhood")
            .Txt_Cidade_0101.Value = Resposta("city")
            .Txt_Estado_0101.Value = Resposta("state")
        End With
    End If
    'Debug.Print xmlhttp.responseText    'Exibe a resposta da API
    Set xmlhttp = Nothing               'Limpa o objeto XMLHttpRequest
    On Error GoTo 0
End Function
