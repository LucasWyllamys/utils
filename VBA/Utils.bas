Attribute VB_Name = "Utils"
Option Explicit

'================================================================================
' Módulo VBA: Utils
' Autor: Lucas Wyllamys Carmo da Silva
' Criado em: 20/01/2026
' Atualizado em: 27/02/2026
' Versão: 1.3.0
' Habilitar bibliotecas:
'   Microsoft Scripting Runtime
'   Microsoft WinHTTP Services, version 5.1
'================================================================================

'=============================================== Enum ===============================================

Public Enum StringType
    cep = 1
    cnpj = 2
    cpf = 3
End Enum

'========================================= Funções Públicas ==========================================

Public Function AbrirArquivoPowerShell(caminhoArquivo As String)
    Shell "powershell.exe -File " & caminhoArquivo, vbNormalFocus
End Function

' ------------------------------------------------------------------------------------
' Descrição:
'   Percorre todas as colunas de uma linha específica e monta um dicionário
'   onde cada chave corresponde ao título localizado na linha de cabeçalhos
'   (linhaChaves) e cada valor corresponde ao conteúdo encontrado na linha
'   informada (linha). Apenas células não vazias são consideradas.
'
' Parâmetros:
'   - planilha (Worksheet): Worksheet onde os dados serão lidos.
'   - linhaChaves (Integer): Número da linha que contém os nomes das chaves (cabeçalhos das colunas).
'   - linha (Long): Número da linha de onde serão lidos os valores associados às chaves.
'   - colunaInicio (Long): Coluna inicial do intervalo a ser varrido.
'   - ultimaColuna (Integer): Última coluna do intervalo a ser varrido.
'
' Retorno: Um objeto Dictionary contendo pares (Chave, Valor) obtidos nas colunas entre colunaInicio e ultimaColuna.
'
' Observações:
'   - As chaves são lidas da linhaChaves.
'   - Os valores são lidos da linha.
'   - Somente células não vazias na linha de valores são adicionadas ao dicionário.
' ------------------------------------------------------------------------------------
Public Function ObterChavesValores( _
    planilha As Worksheet, _
    ByVal linhaChaves As Integer, _
    ByVal linha As Long, _
    ByVal colunaInicio As Long, _
    ByVal ultimaColuna As Integer) As Scripting.dictionary
    
    Dim coluna As Long
    Dim chaveValor As Scripting.dictionary
    
    Set chaveValor = New Scripting.dictionary
    
    With planilha
        For coluna = colunaInicio To ultimaColuna
            ' Verifica se a célula da linha de valores não está vazia
            If .Cells(linha, coluna).Value <> "" Then
                ' Adiciona ao dicionário:
                '   chave = valor da célula na linha de cabeçalhos
                '   valor = conteúdo da célula na linha de dados
                chaveValor.Add .Cells(linhaChaves, coluna).Value, .Cells(linha, coluna).Value
            End If
        Next coluna
    End With
    
    Set ObterChavesValores = chaveValor
End Function

' ------------------------------------------------------------------------------------
' Descrição:
'   Busca uma chave em uma tabela estruturada (ListObject) e retorna o valor de uma
'   coluna específica, deslocada em relação à coluna onde a chave foi encontrada.
'
' Parâmetros:
'   - chave (String): Valor que será procurado na tabela.
'   - planilha (Worksheet): Planilha que contém a tabela onde será feita a busca.
'   - nomeTabela (String): Nome da tabela (ListObject) onde a chave será pesquisada.
'   - colunaOffset (Integer):
'       Quantidade de colunas a partir da coluna onde a chave foi encontrada para
'       obter o valor de retorno. Pode ser positiva ou negativa.
'   - colunaBusca (String) [Opcional]:
'       Nome da coluna onde a chave será procurada.
'       Caso informado, a busca é feita somente nessa coluna, tornando-a mais eficiente.
'
' Retorno (String):
'   Retorna o valor encontrado na posição especificada pelo deslocamento da coluna
'   onde a chave foi encontrada. Retorna vazio se a chave não for localizada.
'
' Observações:
'   - A busca é exata (LookAt:=xlWhole).
'   - Apenas células com valores são consideradas (LookIn:=xlValues).
' ------------------------------------------------------------------------------------
Public Function ObterValorTabela( _
    ByVal chave As String, _
    ByVal planilha As Worksheet, _
    ByVal nomeTabela As String, _
    ByVal colunaOffset As Integer, _
    Optional ByVal colunaBusca As String) As String
    
    Dim lo As ListObject
    Dim rng As Range

    If chave <> "" Then
        Set lo = planilha.ListObjects(nomeTabela) ' Retorna o objeto da tabela
        
        If colunaBusca <> "" Then
            Set rng = lo.ListColumns(colunaBusca).DataBodyRange  ' só o corpo da coluna (sem cabeçalho)
        End If
        
        Set rng = rng.Find(what:=chave, LookIn:=xlValues, LookAt:=xlWhole) 'Retorna a célula encontrada
        If Not rng Is Nothing Then ObterValorTabela = rng.Offset(0, colunaOffset) 'Retorna o caminho do template
    End If
End Function

' Descrição: Substitui chaves no texto por valores do dicionário
' Parâmetros:
'   - text: Texto no qual as chaves serão substituídas pelos valores
'   - keysValues: objeto Scripting.Dictionary com pares chave-valor
Public Function ReplaceKeys(text As String, keysValues As Scripting.dictionary) As String
    Dim key As Variant
    
    If Not keysValues Is Nothing Then ' Verifica se o dicionário não está vazio
        If text <> "" Then
            For Each key In keysValues.keys   ' Itera sobre todos as chaves do dicionário
                text = Replace(text, key, keysValues(key)) ' Substitui os valores das respectivas chaves
            Next key
            ReplaceKeys = text
        End If
    End If
End Function

Public Function DividirTextoEmColecao(ByVal texto As String, Optional ByVal delimitador As String = ";") As Collection
    Dim partes() As String
    Dim resultado As New Collection
    Dim i As Long
    Dim item As String

    texto = Trim(texto)
    delimitador = CStr(delimitador)

    ' Se texto estiver vazio, retorna coleção vazia
    If Len(texto) = 0 Then
        Set DividirTextoEmColecao = resultado
        Exit Function
    End If

    ' Divide o texto pelo delimitador
    partes = Split(texto, delimitador)

    ' Percorre o array e adiciona os itens limpos
    For i = LBound(partes) To UBound(partes)
        item = Trim(partes(i))
        If Len(item) > 0 Then
            resultado.Add item
        End If
    Next i

    Set DividirTextoEmColecao = resultado
End Function

' Esta função valida os tipos de dados de acordo com o tipo informado.
Public Function ValidaValor(valor As String, tipo As StringType) As Boolean
    Dim tamanho As Integer
    
    valor = LimparFormatacao(valor)
    tamanho = Len(valor)
    
    Select Case tipo
        Case cpf And tamanho = 11
            ValidaValor = True
        Case cep And tamanho = 8
            ValidaValor = True
        Case cnpj And tamanho = 14
            ValidaValor = True
        Case Else
            ValidaValor = False
    End Select
End Function

' Esta função formata os dados de acordo com o tipo informado.
Public Function FormataValor(valor As String, tipo As StringType) As String
    valor = LimparFormatacao(valor)
    
    Select Case tipo
        Case cpf
            FormataValor = Format(valor, "000\.000\.000\-00")
        Case cep
            FormataValor = Format(valor, "00\.000\-000")
        Case cnpj
            FormataValor = Format(valor, "00\.000\.000/0000\-00")
    End Select
End Function

'Formato tempoespera: 00:00:00
Public Function Aguardar(tempoEspera As String)
    Dim tempo As Double
    tempo = Now + TimeValue(tempoEspera)
    Application.Wait tempo
End Function

Public Function LimparFormatacao(valor As String) As String
    LimparFormatacao = Trim(Replace(Replace(Replace(valor, ".", ""), "-", ""), "/", ""))
End Function

Public Function AbrirSite(url As String)
    Shell "cmd /c start " & url, vbHide     ' Abre o link do site no navegador padrão.
End Function

Public Function GetUsuario() As String
    GetUsuario = Environ("USERNAME")
    ' GetUsuario = CreateObject("WScript.Network").UserName
End Function

' Esta função usa ADODB.Stream para ler a resposta  de uma requisição HTTP como UTF-8.
Public Function LerRespostaComoUTF8(request As WinHttpRequest) As String
    Dim stream As Object
    Dim responseText As String

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1                             ' adTypeBinary
    stream.Open
    stream.Write request.ResponseBody
    stream.Position = 0
    stream.Type = 2                             ' adTypeText
    stream.Charset = "utf-8"
    responseText = stream.ReadText
    stream.Close
    Set stream = Nothing

    LerRespostaComoUTF8 = responseText
End Function
