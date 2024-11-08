Attribute VB_Name = "Inicializacao"
'
'+------------------------------------------------------+
'|   Projeto de Automatização de Cadastro de Patrimonio |
'| da Fatec Carapícuiba.                                |
'| Carapícuida, 14 de Agosto de 2023.                   |
'+------------{ Desenvolvido por: Nataia de Morais }----+
'
'
'
' Inicio da funcionalidade da macro
Sub MaxTelaInicial()
    ' Desbloqueia a planilia "Padrimonio"
    Sheets("Patrimonio").Unprotect
    Password = "F@tec#2023"
    ' Bloqueia a planilia "HOME" e senha  para poder bloquear
    Sheets("HOME").Protect
    Password = "F@tec#2023"
    
    ' Seleciona a Planili home
    Sheets("HOME").Select
    
    ' ~~{ Inicio das inst de aplicação }~~
    With Application
        
        ' Desabilita a atualização de tela e Desabilitar eventos
        .ScreenUpdating = False
        .EnableEvents = False
    
        .ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
        
        ' Oculta a barra de formula ebarra de status
        .DisplayFormulaBar = False
        .DisplayStatusBar = False
        
        ' Altera o titulo da janela
        .Caption = "~~{ Cadatro de Patrimonio }~~"
        .SendKeys "{HOME}"
    
        ' ~~{ Inicio das intruções da janela }~~
        With ActiveWindow
            ' ~~{ Desabilita partes da planilia que não serão necessarias }~~
            .DisplayHorizontalScrollBar = False
            .DisplayVerticalScrollBar = False
            .DisplayWorkbookTabs = False
            .DisplayHeadings = False
            .DisplayGridlines = False
        
        End With
    End With
End Sub

' ~~{ Inicio da funcionalidade macro
Sub MinTelaInicial()
    ' ~~{ Inicio das inst de aplicação }~~
    With Application
        ' Habilita a atualização de tela e Habilitar eventos
        .ScreenUpdating = True
        .EnableEvents = True
    
        .ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
        
        ' Exibe a barra de formula e Exibe a barra de status
        .DisplayFormulaBar = True
        .DisplayStatusBar = True
        
        ' Altera o titulo da janela
        .Caption = ""
        .SendKeys "{HOME}"
        
        ' ~~{ Inicio das intruções da janela }~~
        With ActiveWindow
            ' ~~{ Habilita partes da planilia que não serão necessarias }~~
            .DisplayHorizontalScrollBar = True
            .DisplayVerticalScrollBar = True
            .DisplayWorkbookTabs = True
            .DisplayHeadings = True
            .DisplayGridlines = True
        
        End With
    End With
    
    ' Bloqueia a planilia "Patrimonio"
    Sheets("Patrimonio").Protect
    Password = "F@tec#2023"
    Sheets("Patrimonio").Select
    Range("A3").Select
End Sub
