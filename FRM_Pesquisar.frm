VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FRM_Pesquisar 
   Caption         =   "Pesquisar Patrimônio"
   ClientHeight    =   13530
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19185
   OleObjectBlob   =   "FRM_Pesquisar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FRM_Pesquisar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'  +------------------------------------------------------+
'  |   Projeto de Automatização de Cadastro de Patrimonio |
'  | da Fatec Carapícuiba.                                |
'  | Carapícuida, 14 de Agosto de 2023.                   |
'  +------------{ Desenvolvido por: Nataia de Morais }----+
'
'
'
Public linha As Integer ' <---- Variaval publica
' Inicio do evento 'btn_Apagar_Click()'
Private Sub btn_Apagar_Click()
    ' --{ Seleciona as Caixas de Texto }--
    Me.txt_NumBem.SetFocus
    Me.txt_NumBem.SelText = ""
    
    ' --{ Limpar caixa de texto }--
    txt_NumBem.Value = ""
    txt_Grupo.Value = ""
    txt_DescrBem.Value = ""
    txt_Cor.Value = ""
    txt_Marca.Value = ""
    txt_Modelo.Value = ""
    txt_NumSala.Value = ""
    txt_NumSerie.Value = ""
    txt_Local.Value = ""
    txt_Processo.Value = ""
    opt_Ativo.Value = False
    opt_Desativado.Value = False
    txt_DataCadas.Value = ""
    txt_Valor.Value = ""
    
End Sub

' Inicio do evento 'btn_Voltar_Click()'
Private Sub btn_Voltar_Click()
    ' ~~{ Fecha o formulario e seleciona a Planilia HOME }~~
    Unload FRM_Pesquisar
    Sheets("HOME").Select
End Sub

' Inicio do evento 'txt_NumBem_AfterUpdate()'
Private Sub txt_NumBem_AfterUpdate()
    
    Me.txt_NumBem.SetFocus
    Me.txt_NumBem.SelText = ""
    ' --{ Declarando variavel }
    Dim numBem, i As Integer
    Dim op As String
    
    ' 'linha' recebe o numero de linha planilia
    linha = Sheets("Patrimonio").Range("B2").End(xlDown).Row
    
    ' ~~{ Inicio do Loop For: Faça ate que o 'i' maior ou igual a 'linha' }~~
    For i = 3 To linha
        
        op = Sheets("Patrimonio").Range("B" & i).Value
        numBem = txt_NumBem.Value
        
        ' ~~{ Tratamento de ERRO }~~
        If numBem = op Then
            ' ~~{ Exibe os dados da tabela nas caixa de texto }~~
            txt_Grupo.Value = Sheets("Patrimonio").Range("C" & i).Value
            txt_DescrBem.Value = Sheets("Patrimonio").Range("D" & i).Value
            txt_Cor.Value = Sheets("Patrimonio").Range("E" & i).Value
            txt_Marca.Value = Sheets("Patrimonio").Range("F" & i).Value
            txt_Modelo.Value = Sheets("Patrimonio").Range("G" & i).Value
            txt_NumSala.Value = Sheets("Patrimonio").Range("H" & i).Value
            txt_NumSerie.Value = Sheets("Patrimonio").Range("I" & i).Value
            txt_Local.Value = Sheets("Patrimonio").Range("J" & i).Value
            
            If Sheets("Patrimonio").Range("L" & i).Value = "Ativo" Then
                ' ~~{ Condição (opt_Ativo = Verdadeiro) = 'Ativado' }~~
                opt_Ativo.Value = True
            Else
                If Sheets("Patrimonio").Range("L" & i).Value = "Desativado" Then
                    ' ~~{ Condição (opt_Desativo = Verdadeiro) = 'Desativado' }~~
                    opt_Desativado.Value = True
                End If
            End If
            
            ' ~~{ Exibe os dados da tabela nas caixa de texto }~~
            txt_Processo.Value = Sheets("Patrimonio").Range("K" & i).Value
            txt_DataCadas.Value = Sheets("Patrimonio").Range("M" & i).Value
            txt_Valor.Value = Sheets("Patrimonio").Range("N" & i).Value
            Exit For
        End If
    Next
    
    ' ~~{ Se o 'i' for maior que Linha faca }~~
    If i > linha And Not txt_NumBem = "" Then
        ' Mostre a mensagem que a cadastro com esse codigo
        msg = MsgBox("Não foi posivel efetuar a pesquisa, patrimônio invalido", _
            vbOKOnly + vbCritical, "ERROR")
    End If
End Sub

' Inicio do evento do Formulario (evento Ativar)
Private Sub UserForm_Activate()
    ' --{ Salva o tamanho do UserForm }--
    Dim alt, lar, zoom As Integer
    
    If Application.Height < 700 Then 'Analisar
    ' ~~{ Se o tamanho janela (altura) for menor que 700 faça }~~
        alt = 97
        lar = 70
        zoom = 75
    Else
        If Application.Height > 700 Then 'rever com monitor maior
        ' ~~{ Se o tamanho da janela (altura) for maior que 700 }~~
            alt = 80
            lar = 60
            zoom = 90
        End If
    End If

    ' --{ Diminue a altura e largura do formulario }--
    Me.Height = (Application.Height * alt) / 100
    Me.Width = (Application.Width * lar) / 100
    ' Ajusta o tamanho e a posição do UserForm
    Me.zoom = zoom
    ' Tag recebe a altura do UserForm
    Tag = Height
    ' Renomeia o titulo do formulário
    FRM_Pesquisar.Caption = "{ Pesquisar Patrimônio }"
End Sub

' Inicio do evento evento Clicar do Formulario
Private Sub UserForm_Click()
    ' --{ Declarando variavel }--
    Dim NewHeight As Single
    Dim alt, lar As String
    ' --{ Salva o tamanho do UserForm }--
    alt = 705.75
    lar = 971.25
    
    ' Tamanho atual do UserForm
    NewHeight = Height
    ' ~~{ Se 'NewHeigh' for igual a 'Tag' diminua a tela do formulario }~~
    If NewHeight = Tag Then
    
        ' --{ Reajusta o formulario para o tamanho padrao }--
        Me.Height = Val(alt)
        Me.Width = Val(lar)
        Me.zoom = 100
        
    Else
        ' ~~{ Eeculta o evento UserForm_Activate }~~
        Call UserForm_Activate
    End If
End Sub

' Inicio do evento Redimencionardo Formulario
Private Sub UserForm_Resize()
    ' ~~{ Renomeia o titulo do formulario }~~
    FRM_Pesquisar.Caption = "{ Pesquisar Patrimônio: Click para reajustar! }"
End Sub

