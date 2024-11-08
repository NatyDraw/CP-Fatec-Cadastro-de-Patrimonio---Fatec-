Attribute VB_Name = "Macros"
'
'+------------------------------------------------------+
'|   Projeto de Automatização de Cadastro de Patrimonio |
'| da Fatec Carapícuiba.                                |
'| Carapícuida, 14 de Agosto de 2023.                   |
'+------------{ Desenvolvido por: Nataia de Morais }----+
'
'
'
' ~~{ Inicio da funcionalidade macro da auto_open }~~
Public Sub auto_open()
    ' Seleciona a macro "TelaInicial"
    Call MaxTelaInicial
End Sub

' ~~{ Inicio da funcionalidade macro da Excluir }~~
Sub Excluir()
    ' Exibe o formulario de Cadastro "FRM_ExcluirDados"
   FRM_ExcluirDados.Show
End Sub

' ~~{ Inicio da funcionalidade macro da Cadastrar }~~
Sub Cadastrar()
    ' Exibe o formulario de Cadastro "FRM_Cadastro"
    FRM_Cadastro.Show
End Sub

' ~~{ Inicio da funcionalidade macro da Pesquisar }~~
Sub Pesquisar()
    ' Exibe o formulario de Cadastro "FRM_Pesquisar"
    FRM_Pesquisar.Show
End Sub

' ~~{ Inicio da funcionalidade macro da ExibirTab }~~
Sub Editar()
    ' Exibe o formulario de Cadastro "FRM_Pesquisar"
    FRM_EditarDados.Show
End Sub

' ~~{ Inicio da funcionalidade macro da ExibirTab }~~
Sub ExibirTab()
    ' ~~{ Seleciona a tabela "Patrimonio" e Selecionar celula }~~
    Sheets("Patrimonio").Select
    Range("A3").Select
    Sheets("Patrimonio").Protect
    Password = "F@tec#2023"
    
    ' ~~{ Inicio das instruções }~~
    With ActiveWindow
        ' Habilita a barra Vertical da planilia e a barra Horizontal da planilia
        .DisplayVerticalScrollBar = True
        .DisplayHorizontalScrollBar = True
    End With
End Sub

' ~~{ Inicio da funcionalidade macro da Help }~~
Sub Help()
    ' Seleciona a tabela "Ajuda" e Bloqueia a planilia Ajuda
    Sheets("Ajuda").Select
    Range("L4").Select
    Sheets("Ajuda").Protect
    Password = "F@tec#2023"
    
    ' ~~{ Inicio das intruões da jenale }~~
    With ActiveWindow
    ' ~~{ Desabilita partes da planilia que não serão necessarias }~~
        .DisplayWorkbookTabs = True
        .DisplayVerticalScrollBar = True
    End With
End Sub

' ~~{ Inicio da funcionalidade macro da Page_1 }~~
Sub Page_1()
    ' Seleciona a tabela "Page_1"
    Sheets("1").Select
    Range("L4").Select
    Sheets("1").Protect
    Password = "F@tec#2023"
End Sub

' ~~{ Inicio da funcionalidade macro da Page_2 }~~
Sub Page_2()
    ' Seleciona a tabela "Page_2"
    Sheets("2").Select
    Range("L4").Select
    Sheets("2").Protect
    Password = "F@tec#2023"
End Sub

' ~~{ Inicio da funcionalidade macro da Page_3 }~~
Sub Page_3()
    ' Seleciona a tabela "Page_3"
    Sheets("3").Select
    Range("L4").Select
    Sheets("3").Protect
    Password = "F@tec#2023"
End Sub
