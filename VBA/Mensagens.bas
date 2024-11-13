Attribute VB_Name = "Mensagens"
Option Explicit

Function msg_operacao_concluida()
    msg_operacao_concluida = MsgBox("Operação concluída!", vbInformation)
End Function

Function msg_operacao_concluida_id(id As String)
    msg_operacao_concluida_id = MsgBox("Operação concluída!" & vbNewLine + vbNewLine & "id: " & id, vbInformation)
End Function

Function msg_confirmacao() As Integer
    msg_confirmacao = MsgBox("Deseja prosseguir com a operação?", vbQuestion + vbYesNo)
End Function

Function msg_campos_obrigatorios()
    msg_campos_obrigatorios = MsgBox("Preencha todos os campos obrigatórios!", vbExclamation)
End Function

Function msg_usuario_incorreto()
    msg_usuario_incorreto = MsgBox("Usuário incorreto!", vbExclamation)
End Function

Function msg_senha_incorreta()
    msg_senha_incorreta = MsgBox("Senha incorreta!", vbExclamation)
End Function

Function msg_registro_inexistente()
    msg_registro_inexistente = MsgBox("Nenhum registro encontrado!", vbExclamation)
End Function

Function msg_api_nao_retornou()
    msg_api_nao_retornou = MsgBox("A consulta não retornou valores!", vbExclamation)
End Function

Function msg_dados_invalidos()
    msg_dados_invalidos = MsgBox("Dados inválidos!", vbExclamation)
End Function

Function msg_erro_anexo()
    msg_erro_anexo = MsgBox("Houve um erro ao anexar o arquivo, provavelmente porque ele está aberto!" & vbNewLine + vbNewLine & _
    "Você deve fechar o arquivo e tentar anexá-lo novamente.", vbCritical)
End Function

Function msg_arquivo_nao_encontrado()
    msg_arquivo_nao_encontrado = MsgBox("O arquivo não foi encontrado ou você não tem este acesso!", vbExclamation)
End Function

Function msg_validacao_tipos_campos()
    msg_validacao_tipos_campos = MsgBox("Há inconsistências no(s) tipo(s) de dado(s)!", vbExclamation)
End Function
