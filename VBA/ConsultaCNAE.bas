Attribute VB_Name = "ConsultaCNAE"
Option Explicit

' Importar Módulo (JsonConverter): https://github.com/VBA-tools/VBA-JSON
' Habilitar Biblioteca: Microsoft Scripting Runtime
' Habilitar Biblioteca: Microsoft WinHTTP Services

' Este módulo destina-se consultar dados de um CNAE via API do IBGE.
' A API usada neste módulo retorna requisições apenas de classes de CNAEs, ou seja, apenas os 5 primeiros digitos de um CNAE pode ser usado como parâmetro.
' Para saber mais sobre esta API (documentação): https://servicodados.ibge.gov.br/api/docs/CNAE?versao=2
'-------------------------------------------------------------------------------------------
    '1. Seção
    '   2. Divisão
    '     3. Grupo
    '       4. Classe (!)
    '         5. Subclasse
    '           6. Atividade econômica
    
    'Exemplo:
    'Seção      A           Agricultura, pecuária, produção fl orestal, pesca e aqüicultura
    'Divisão    01          Agricultura, pecuária e serviços relacionados
    'Grupo      01.1        Produção de lavouras temporárias
    'Classe     01.11-3     Cultivo de cereais (!)
    'Subclasse  0111-3/01   Cultivo de arroz
'-------------------------------------------------------------------------------------------
'Seções da CNAE:----------------------------------------------------------------------------
    'Seção      Denominação
    'A          Agricultura, pecuária, produção fl orestal, pesca e aqüicultura
    'B          Indústrias extrativas
    'C          Indústrias de transformação
    'D          Eletricidade e gás
    'E          Água, esgoto, atividades de gestão de resíduos e descontaminação
    'F          Construção
    'G          Comércio; reparação de veículos automotores e motocicletas
    'H          Transporte, armazenagem e correio
    'I          Alojamento e alimentação
    'J          Informação e comunicação
    'K          Atividades fi nanceiras, de seguros e serviços relacionados
    'L          Atividades imobiliárias
    'M          Atividades profi ssionais, científi cas e técnicas
    'N          Atividades administrativas e serviços complementares
    'O          Administração pública, defesa e seguridade social
    'P          Educação
    'Q          Saúde humana e serviços sociais
    'R          Artes, cultura, esporte e recreação
    'S          Outras atividades de serviços
    'T          Serviços domésticos
    'U          Organismos internacionais e outras instituições extraterritoriais
'-------------------------------------------------------------------------------------------
Public Function ConsultarCNAE(CNAE As String) As Scripting.Dictionary
    Dim Requisicao As New WinHttpRequest, DictDescricoesCNAE As Object, Resposta As Object, _
    Grupo As Object, Divisao As Object, Secao As Object
    Dim url As String, ResponseText As String
    
    Set DictDescricoesCNAE = CreateObject("Scripting.Dictionary")
    
    url = "https://servicodados.ibge.gov.br/api/v2/cnae/classes/" & CNAE

    Requisicao.SetProxy 2, "10.219.78.3:8080", ""                                   ' Servidor padrão do proxy COELBA
    Requisicao.Open "GET", url, False                                               ' Abre uma conexão com a API
    Requisicao.SetRequestHeader "Accept", "application/json"
    'Requisicao.SetRequestHeader "Content-Type", "application/json; charset=UTF-8"   ' Trata a requisição como UTF-8
    Requisicao.Send                                                                 ' Envia a requisição para a API

    If Requisicao.Status = 200 Then
        With DictDescricoesCNAE
            If Requisicao.ResponseText <> "[]" Then
                ResponseText = LerRespostaComoUTF8.LerRespostaComoUTF8(Requisicao)   ' Converte a requisição em UTF-8
                Set Resposta = JsonConverter.ParseJson(ResponseText)                 ' Converte o JSON em dicionário
                    .Add "id_classe", Resposta("id")
                    .Add "descricao_classe", Resposta("descricao")
                'Grupo---------------------------------
                Set Grupo = Resposta("grupo")
                    .Add "id_grupo", Grupo("id")
                    .Add "descricao_grupo", Grupo("descricao")
                'Divisão-------------------------------
                Set Divisao = Grupo("divisao")
                    .Add "id_divisao", Divisao("id")
                    .Add "descricao_divisao", Divisao("descricao")
                'Seção---------------------------------
                    Set Secao = Divisao("secao")
                    .Add "id_secao", Secao("id")
                    .Add "descricao_secao", Secao("descricao")
                '--------------------------------------
                Set ConsultarCNAE = DictDescricoesCNAE
            End If
        End With
    Else
        Stop
        MsgBox "Erro na requisição: " & Requisicao.Status
    End If
    ' Limpar objetos
    Set Requisicao = Nothing
    Set DictDescricoesCNAE = Nothing
    Set Resposta = Nothing
    Set Grupo = Nothing
    Set Divisao = Nothing
    Set Secao = Nothing
End Function
