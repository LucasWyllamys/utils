Attribute VB_Name = "Mensagens"
Option Explicit

Public Function MsgOperacaoConcluida() As Integer
    MsgOperacaoConcluida = MsgBox("Opera��o conclu�da!", vbInformation)
End Function

Public Function MsgOperacaoConcluidaID(id As String) As Integer
    MsgOperacaoConcluidaID = MsgBox("Opera��o conclu�da!" & vbNewLine + vbNewLine & "id: " & id, vbInformation)
End Function

Public Function MsgConfirmacao() As Integer
    MsgConfirmacao = MsgBox("Deseja prosseguir com a opera��o?", vbQuestion + vbYesNo)
End Function

Public Function MsgCamposObrigatorios() As Integer
    MsgCamposObrigatorios = MsgBox("Preencha os campos obrigat�rios!", vbExclamation)
End Function

Public Function MsgUsuarioIncorreto() As Integer
    MsgUsuarioIncorreto = MsgBox("Usu�rio incorreto!", vbExclamation)
End Function

Public Function MsgSenhaIncorreta() As Integer
    MsgSenhaIncorreta = MsgBox("Senha incorreta!", vbExclamation)
End Function

Public Function MsgRegistroInexistente() As Integer
    MsgRegistroInexistente = MsgBox("Nenhum registro encontrado!", vbExclamation)
End Function

Public Function MsgAPInaoRetornou() As Integer
    MsgAPInaoRetornou = MsgBox("A consulta n�o retornou valores!", vbExclamation)
End Function

Public Function MsgDadosInvalidos() As Integer
    MsgDadosInvalidos = MsgBox("Dados inv�lidos!", vbExclamation)
End Function

Public Function MsgErro() As Integer
    MsgErro = MsgBox("Ocerreu um erro: " & Err.Description, vbCritical)
End Function

Public Function MsgArquivoNaoEncontrado() As Integer
    MsgArquivoNaoEncontrado = MsgBox("O arquivo n�o foi encontrado ou voc� n�o tem acesso!", vbExclamation)
End Function

Public Function MsgValidacaoDados() As Integer
    MsgValidacaoDados = MsgBox("H� inconsist�ncia de dados!", vbExclamation)
End Function
