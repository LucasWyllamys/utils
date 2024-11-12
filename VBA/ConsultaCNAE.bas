Attribute VB_Name = "ConsultaCNAE"
Option Explicit

' Importar M�dulo (JsonConverter): https://github.com/VBA-tools/VBA-JSON
' Habilitar Biblioteca: Microsoft Scripting Runtime
' Habilitar Biblioteca: Microsoft WinHTTP Services

' Este m�dulo destina-se consultar dados de um CNAE via API do IBGE.
' A API usada neste m�dulo retorna requisi��es apenas de classes de CNAEs, ou seja, apenas os 5 primeiros digitos de um CNAE pode ser usado como par�metro.
' Para saber mais sobre esta API (documenta��o): https://servicodados.ibge.gov.br/api/docs/CNAE?versao=2
'-------------------------------------------------------------------------------------------
    '1. Se��o
    '   2. Divis�o
    '     3. Grupo
    '       4. Classe (!)
    '         5. Subclasse
    '           6. Atividade econ�mica
    
    'Exemplo:
    'Se��o      A           Agricultura, pecu�ria, produ��o fl orestal, pesca e aq�icultura
    'Divis�o    01          Agricultura, pecu�ria e servi�os relacionados
    'Grupo      01.1        Produ��o de lavouras tempor�rias
    'Classe     01.11-3     Cultivo de cereais (!)
    'Subclasse  0111-3/01   Cultivo de arroz
'-------------------------------------------------------------------------------------------
'Se��es da CNAE:----------------------------------------------------------------------------
    'Se��o      Denomina��o
    'A          Agricultura, pecu�ria, produ��o fl orestal, pesca e aq�icultura
    'B          Ind�strias extrativas
    'C          Ind�strias de transforma��o
    'D          Eletricidade e g�s
    'E          �gua, esgoto, atividades de gest�o de res�duos e descontamina��o
    'F          Constru��o
    'G          Com�rcio; repara��o de ve�culos automotores e motocicletas
    'H          Transporte, armazenagem e correio
    'I          Alojamento e alimenta��o
    'J          Informa��o e comunica��o
    'K          Atividades fi nanceiras, de seguros e servi�os relacionados
    'L          Atividades imobili�rias
    'M          Atividades profi ssionais, cient�fi cas e t�cnicas
    'N          Atividades administrativas e servi�os complementares
    'O          Administra��o p�blica, defesa e seguridade social
    'P          Educa��o
    'Q          Sa�de humana e servi�os sociais
    'R          Artes, cultura, esporte e recrea��o
    'S          Outras atividades de servi�os
    'T          Servi�os dom�sticos
    'U          Organismos internacionais e outras institui��es extraterritoriais
'-------------------------------------------------------------------------------------------
Public Function ConsultarCNAE(CNAE As String) As Scripting.Dictionary
    Dim Requisicao As New WinHttpRequest, DictDescricoesCNAE As Object, Resposta As Object, _
    Grupo As Object, Divisao As Object, Secao As Object
    Dim url As String, ResponseText As String
    
    Set DictDescricoesCNAE = CreateObject("Scripting.Dictionary")
    
    url = "https://servicodados.ibge.gov.br/api/v2/cnae/classes/" & CNAE

    Requisicao.SetProxy 2, "10.219.78.3:8080", ""                                   ' Servidor padr�o do proxy COELBA
    Requisicao.Open "GET", url, False                                               ' Abre uma conex�o com a API
    Requisicao.SetRequestHeader "Accept", "application/json"
    'Requisicao.SetRequestHeader "Content-Type", "application/json; charset=UTF-8"   ' Trata a requisi��o como UTF-8
    Requisicao.Send                                                                 ' Envia a requisi��o para a API

    If Requisicao.Status = 200 Then
        With DictDescricoesCNAE
            If Requisicao.ResponseText <> "[]" Then
                ResponseText = LerRespostaComoUTF8.LerRespostaComoUTF8(Requisicao)   ' Converte a requisi��o em UTF-8
                Set Resposta = JsonConverter.ParseJson(ResponseText)                 ' Converte o JSON em dicion�rio
                    .Add "id_classe", Resposta("id")
                    .Add "descricao_classe", Resposta("descricao")
                'Grupo---------------------------------
                Set Grupo = Resposta("grupo")
                    .Add "id_grupo", Grupo("id")
                    .Add "descricao_grupo", Grupo("descricao")
                'Divis�o-------------------------------
                Set Divisao = Grupo("divisao")
                    .Add "id_divisao", Divisao("id")
                    .Add "descricao_divisao", Divisao("descricao")
                'Se��o---------------------------------
                    Set Secao = Divisao("secao")
                    .Add "id_secao", Secao("id")
                    .Add "descricao_secao", Secao("descricao")
                '--------------------------------------
                Set ConsultarCNAE = DictDescricoesCNAE
            End If
        End With
    Else
        Stop
        MsgBox "Erro na requisi��o: " & Requisicao.Status
    End If
    ' Limpar objetos
    Set Requisicao = Nothing
    Set DictDescricoesCNAE = Nothing
    Set Resposta = Nothing
    Set Grupo = Nothing
    Set Divisao = Nothing
    Set Secao = Nothing
End Function
