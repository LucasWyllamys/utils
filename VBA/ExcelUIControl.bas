Attribute VB_Name = "ExcelUIControl"
Option Explicit

' Desliga recursos que deixam o Excel lento durante automações,
' como atualização de tela, eventos automáticos e alertas.
' Esse método deixa suas macros mais rápidas e sem travamentos.
Public Sub EnablePerformanceMode()
    Application.ScreenUpdating = False ' Para de atualizar a tela a cada ação ? acelera muito.
    Application.DisplayAlerts = False ' Impede janelas como "Deseja salvar?" durante o processamento.
    Application.EnableEvents = False ' Desliga eventos automáticos que poderiam disparar macros indesejadas.
    Application.Calculation = xlCalculationManual ' Desliga recálculo automático ? evita quedas de performance.
End Sub

' Reativa os recursos desativados no modo sistema.
' Volta o Excel ao comportamento normal de atualização/cálculo.
Public Sub DisablePerformanceMode()
    Application.ScreenUpdating = True ' Atualiza a tela novamente.
    Application.DisplayAlerts = True ' Volta a exibir mensagens e alertas.
    Application.EnableEvents = True ' Eventos automáticos são reativados.
    Application.Calculation = xlCalculationAutomatic ' Cálculo automático ativado novamente.
End Sub

' Oculta elementos visuais do Excel (Ribbon, barras, abas, grades),
' deixando o visual totalmente limpo — como se fosse um software próprio.
Public Sub HideUI()
    Application.DisplayFullScreen = True ' Coloca o Excel em tela cheia.
    Application.DisplayFormulaBar = False ' Remove a barra de fórmulas.
    Application.DisplayStatusBar = False ' Oculta a barra inferior de status.

    Application.ExecuteExcel4Macro "Show.ToolBar(""Ribbon"", False)" ' Oculta a Ribbon (comando Excel 4.0, único que funciona sempre).

    With Application.ActiveWindow ' A janela ativa contém elementos visuais da planilha.
        .DisplayWorkbookTabs = False ' Oculta as abas das planilhas.
        .DisplayHeadings = False ' Oculta cabeçalhos (A/B/1/2).
        .DisplayGridlines = False ' Oculta grades da planilha.
        .DisplayHorizontalScrollBar = False ' Some com barra de rolagem horizontal.
        .DisplayVerticalScrollBar = False ' Some com barra de rolagem vertical.
    End With
End Sub

' Restaura todos os elementos visuais do Excel, retornando o aplicativo ao estado original.
Public Sub ShowUI()
    Application.DisplayFullScreen = False ' Sai do modo tela cheia.
    Application.DisplayFormulaBar = True ' Mostra barra de fórmulas.
    Application.DisplayStatusBar = True ' Mostra barra de status.

    Application.ExecuteExcel4Macro _
        "Show.ToolBar(""Ribbon"", True)" ' Traz a Ribbon de volta.

    With Application.ActiveWindow ' Restaura elementos de interface da janela ativa.
        .DisplayWorkbookTabs = True ' Exibe abas das planilhas.
        .DisplayHeadings = True ' Exibe cabeçalhos.
        .DisplayGridlines = True ' Exibe grade.
        .DisplayHorizontalScrollBar = True ' Exibe barra horizontal.
        .DisplayVerticalScrollBar = True ' Exibe barra vertical.
    End With
End Sub

' Ativa todo o modo sistema melhorando a performance e aplicando aparência de sistema no Excel.
Public Sub EnterSystemMode()
    Call DesativarPerformance ' Melhora execução das macros.
    Call OcultarUI ' Deixa o Excel com aparência de sistema.
End Sub

' Desfaz tudo que o modo sistema modificou.
Public Sub ExitSystemMode()
    RestaurarPerformance ' Reativa cálculo, eventos e alertas.
    ExibirUI ' Restaura aparência padrão do Excel.
End Sub
