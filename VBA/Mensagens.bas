Attribute VB_Name = "Mensagens"
Option Explicit

Public Function MsgOperacaoConcluida() As Integer
    MsgOperacaoConcluida = MsgBox("Operação concluída!", vbInformation)
End Function

Public Function MsgOperacaoConcluidaID(id As String) As Integer
    MsgOperacaoConcluidaID = MsgBox("Operação concluída!" & vbNewLine + vbNewLine & "id: " & id, vbInformation)
End Function

Public Function MsgConfirmacao() As Integer
    MsgConfirmacao = MsgBox("Deseja prosseguir com a operação?", vbQuestion + vbYesNo)
End Function

Public Function MsgCamposObrigatorios() As Integer
    MsgCamposObrigatorios = MsgBox("Preencha os campos obrigatórios!", vbExclamation)
End Function

Public Function MsgUsuarioIncorreto() As Integer
    MsgUsuarioIncorreto = MsgBox("Usuário incorreto!", vbExclamation)
End Function

Public Function MsgSenhaIncorreta() As Integer
    MsgSenhaIncorreta = MsgBox("Senha incorreta!", vbExclamation)
End Function

Public Function MsgRegistroInexistente() As Integer
    MsgRegistroInexistente = MsgBox("Nenhum registro encontrado!", vbExclamation)
End Function

Public Function MsgAPInaoRetornou() As Integer
    MsgAPInaoRetornou = MsgBox("A consulta não retornou valores!", vbExclamation)
End Function

Public Function MsgDadosInvalidos() As Integer
    MsgDadosInvalidos = MsgBox("Dados inválidos!", vbExclamation)
End Function

Public Function MsgErro() As Integer
    MsgErro = MsgBox("Ocerreu um erro: " & Err.Description, vbCritical)
End Function

Public Function MsgArquivoNaoEncontrado() As Integer
    MsgArquivoNaoEncontrado = MsgBox("O arquivo não foi encontrado ou você não tem acesso!", vbExclamation)
End Function

Public Function MsgValidacaoDados() As Integer
    MsgValidacaoDados = MsgBox("Há inconsistência de dados!", vbExclamation)
End Function
