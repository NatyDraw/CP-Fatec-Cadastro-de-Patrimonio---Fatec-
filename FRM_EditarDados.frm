VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FRM_EditarDados 
   Caption         =   "Editar Dado do Patrimônio"
   ClientHeight    =   13530
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19185
   OleObjectBlob   =   "FRM_EditarDados.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FRM_EditarDados"
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

' Inicio do evento 'btm_Editar_Click()'
Private Sub btn_Editar_Click()
    ' ~~{ Seleciona a planilia, valiavel 'linha' recebe um valor da quantidade de linhas da planilia 'Patrimonio' }~~
    Sheets("Patrimonio").Select
    linha = Cells.Find(txt_NumBem.Value).Row
    
    ' --{ Salvana planilia os dados digitados nas caixas de texto}
    Sheets("Patrimonio").Range("H" & linha).Value = txt_NumSala.Value
    Sheets("Patrimonio").Range("I" & linha).Value = txt_NumSerie.Value
    Sheets("Patrimonio").Range("J" & linha).Value = txt_Local.Value
        
    If opt_Ativo.Value = True Then
        Sheets("Patrimonio").Range("L" & linha).Value = "Ativo"
    Else
        If opt_Desativado.Value = True Then
            Sheets("Patrimonio").Range("L" & linha).Value = "Desativado"
        End If
    End If
    ' ~~{ Mensagem exibida ao usuario sobre a edição efetuado }~~
    msg = MsgBox("Dados do patrimônio editado com sucesso.", vbOK + vbInformation, "Editando...")
    
    ' Chama a Macro que limpa as caixas de texto e retorna para a planilia HOME
    Call btn_Apagar_Click
    Sheets("HOME").Select
End Sub

' Inicio do evento 'btn_Apagar_Click()'
Private Sub btn_Apagar_Click()
    ' --{ Desbloqueia a caisxa de texto e Torna o botão visivel }--
    txt_NumBem.Locked = False
    btn_Editar.Visible = False
    
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
    ' ~~{ Fecha o formulario e seleciona a planilia HOME }~~
    Unload FRM_EditarDados
    Sheets("HOME").Select
End Sub

' Inicio do evento 'txt_DataCadas_Change()'
Private Sub txt_DataCadas_Change()
    ' ~~{ Para poder armazenar  10 digitos, Ex:. 10/10/2014 }~~
    txt_DataCadas.MaxLength = 10
    
    ' ~~{ Neste If, ele adiciona uma "/" Quando for digitado 2 digitos }~~
    If Len(txt_DataCadas.Text) = 2 Then
        txt_DataCadas = txt_DataCadas + "/"
    End If
    
    ' ~~{ Neste If, ele adiciona uma "/" Quando for digitado +2 digitos }~~
    If Len(txt_DataCadas.Text) = 5 Then
        txt_DataCadas = txt_DataCadas + "/"
    End If
        
    ' ~~{ Neste If, ele adiciona uma "/" Quando for digitado +4 digitos }~~
    If Len(txt_DataCadas.Text) = 10 Then
        Application.SendKeys "<TAB>"
    End If
End Sub

' ~~{ Inicio do evento 'txt_NumBem_AfterUpdate()' }~~
Private Sub txt_NumBem_AfterUpdate()
    ' --{ Declarando variavel }--
    Dim i As Integer
    Dim op As String
    
    ' 'linha' recebe o numero de linha planilia
    linha = Sheets("Patrimonio").Range("B2").End(xlDown).Row
    
    ' ~~{ Inicio do Loop For: Faça ate que o "i" maior ou igual a 'linha' }~~
    For i = 3 To linha
        op = Sheets("Patrimonio").Range("B" & i).Value
        
        ' ~~{ Tratamento de ERRO }~~
        If Me.txt_NumBem.Value = op Then
            ' --{ Exibe os dados da tabela nas caixa de texto }--
            Me.txt_Grupo.Value = Sheets("Patrimonio").Range("C" & i).Value
            Me.txt_DescrBem.Value = Sheets("Patrimonio").Range("D" & i).Value
            Me.txt_Cor.Value = Sheets("Patrimonio").Range("E" & i).Value
            Me.txt_Marca.Value = Sheets("Patrimonio").Range("F" & i).Value
            Me.txt_Modelo.Value = Sheets("Patrimonio").Range("G" & i).Value
            Me.txt_NumSala.Value = Sheets("Patrimonio").Range("H" & i).Value
            Me.txt_NumSerie.Value = Sheets("Patrimonio").Range("I" & i).Value
            Me.txt_Local.Value = Sheets("Patrimonio").Range("J" & i).Value

            If Sheets("Patrimonio").Range("L" & i).Value = "Ativo" Then
                ' ~~{ Condição (opt_Ativo = Verdadeiro) = 'Ativado' }~~
                Me.opt_Ativo.Value = True
            Else
                If Sheets("Patrimonio").Range("L" & i).Value = "Desativado" Then
                    ' ~~{Condição (opt_Desativo = Verdadeiro) = 'Desativado' }~~
                    Me.opt_Desativado.Value = True
                End If
            End If

            Me.txt_Processo.Value = Sheets("Patrimonio").Range("K" & i).Value
            Me.txt_DataCadas.Value = Sheets("Patrimonio").Range("M" & i).Value
            Me.txt_Valor.Value = Sheets("Patrimonio").Range("N" & i).Value
            
            ' Chama a Macro que desbloqueia as caixa de texto e torna visivel o botão editar
            Call desbloqueia
            Exit For
        End If
    Next
    
    If i > linha And Not txt_NumBem = "" Then
        ' ~~{ Mensagem exibida ao usuario sobre o Cadastro não existente }~~
        msg = MsgBox("Não foi posivel efetuar a pesquisa. Patrimônio invalido!", _
            vbOKOnly + vbCritical, "ERRO")
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
    FRM_EditarDados.Caption = "{ Editar Patrimônio }"
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
    FRM_EditarDados.Caption = "{ Editar Patrimônio: Click para reajustar! }"
End Sub

' Macro que desbloqueia as caixas de texto
Private Sub desbloqueia()

    If Not Me.txt_NumBem.Value = "" Then
        ' --{ BLOQUEIA A CAIXA DE TEXTO DO NUMERO DE BEM }--
        Me.txt_NumBem.Locked = True
        ' --{ TORNA VISIVEL O BOTAO EDITAR }--
        btn_Editar.Visible = True
        ' --{ DESBLOQUEIA AS DEMAIS CAIXA DE TEXTO }--
        Me.txt_NumSala.Locked = False
        Me.txt_Local.Locked = False
        Me.txt_NumSerie.Locked = False
        Me.opt_Ativo.Locked = False
        Me.opt_Desativado.Locked = False
        
    Else
        ' --{ DESBLOQUEIA A CAIXA DE TEXTO DO NUMERO DE BEM }--
        Me.txt_NumBem.Locked = False
        ' --{ TORNA INVISIVEL O BOTAO EDITAR }--
        btn_Editar.Visible = False
        ' --{ BLOQUEIA AS DEMAIS CAIXA DE TEXTO }--
        Me.txt_NumSala.Locked = True
        Me.txt_Local.Locked = True
        Me.txt_NumSerie.Locked = True
        Me.opt_Ativo.Locked = True
        Me.opt_Desativado.Locked = True
        
    End If
End Sub
