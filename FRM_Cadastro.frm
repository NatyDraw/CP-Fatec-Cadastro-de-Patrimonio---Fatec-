VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FRM_Cadastro 
   Caption         =   "Cadastrar Patrimônio"
   ClientHeight    =   13530
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19185
   OleObjectBlob   =   "FRM_Cadastro.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FRM_Cadastro"
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
' Inicio do evento 'btn_Cadastrar_Click'
Private Sub btn_Cadastrar_Click()
    ' --{ Declarando variavel }--
    Dim i, id As Integer
    Dim op As String
    
    'linha' = ao numero de linha da planilha
    linha = Sheets("Patrimonio").Range("B2").End(xlDown).Row
    
    For i = 3 To linha
        ' ~~{ Inicio do Loop: Faça ate que o 'i' maior ou igual a 'linha' }~~
        op = Sheets("Patrimonio").Range("B" & i).Value
        
        ' ~~{ Tratamento de ERRO: Se o 'numBem' For igual a 'op' faca }~~
        If Me.txt_NumBem.Value = op Then
            ' --{ INSERE NO BANCO DE DADOS }--
            Sheets("Patrimonio").Range("F" & linha).Value = Me.txt_NumSala.Value
            Sheets("Patrimonio").Range("G" & linha).Value = Me.txt_NumSerie.Value
            Sheets("Patrimonio").Range("H" & linha).Value = Me.txt_Local.Value
            
            ' ~~{ Se o Opt_Ativo for igual a verdadeiro faça }~~
            If opt_Ativo.Value = True Then
                Sheets("Patrimonio").Range("J" & linha).Value = "Ativo"
            Else
                ' ~~{ Se o Opt_Desativo for igual a verdadeiro faça }~~
                If opt_Desativado.Value = True Then
                    Sheets("Patrimonio").Range("J" & linha).Value = "Desativado"
                End If
                
            End If
            ' Break do Loop
            Exit For
            
        End If
    ' Proximo Loop i++
    Next
    ' ~~{ Se o 'i' for maior que 'Linha' faca }~~
    If i > linha Then
        
        ' ~~{ ARMAZENA DADOS NOVOS }~~
        id = Sheets("Patrimonio").Range("A" & i - 1).Value + 1
        Sheets("Patrimonio").Range("A" & i).Value = id
        Sheets("Patrimonio").Range("B" & i).Value = txt_NumBem.Value
        Sheets("Patrimonio").Range("C" & i).Value = txt_Grupo.Value
        Sheets("Patrimonio").Range("D" & i).Value = txt_DescrBem.Value
        Sheets("Patrimonio").Range("E" & i).Value = txt_Cor.Value
        Sheets("Patrimonio").Range("F" & i).Value = txt_Marca.Value
        Sheets("Patrimonio").Range("G" & i).Value = txt_Modelo.Value
        Sheets("Patrimonio").Range("H" & i).Value = txt_NumSala.Value
        Sheets("Patrimonio").Range("I" & i).Value = txt_NumSerie.Value
        Sheets("Patrimonio").Range("J" & i).Value = txt_Local.Value
        Sheets("Patrimonio").Range("K" & i).Value = txt_Processo.Value
        
        ' ~~{ Se o Opt_Ativo for igual a verdadeiro faça }~~
        If opt_Ativo.Value = True Then
                Sheets("Patrimonio").Range("L" & i).Value = "Ativo"
        Else
            ' ~~{ Se o Opt_Desativo for igual a verdadeiro faça }~~
            If opt_Desativado.Value = True Then
                Sheets("Patrimonio").Range("L" & i).Value = "Desativado"
            End If
        
        End If
        
        Sheets("Patrimonio").Range("M" & i).Value = txt_DataCadas.Value
        Sheets("Patrimonio").Range("N" & i).Value = txt_Valor.Value
        ' Vai para macro que formata o texto
        Call Formatar(i)
        
    End If
    
    ' Usar o codigo do evento btn_Apagar_Click'
    Call btn_Apagar_Click
    Call locked_text
    ' Mensagem exibida ao usuario sobre o cadastro efetuado
    msg = MsgBox("Patrimônio atualizado com sucesso.", vbOKOnly + vbInformation, "Cadastrando...")
    
    ' Volta para o inicio
    Sheets("HOME").Select
End Sub

' Inicio do evento 'btn_Apagar_Click'
Private Sub btn_Apagar_Click()
    ' ~~{ Desbloqueia a caisxa de texto e Torna o botão visivel }~~
    txt_NumBem.Locked = False
    btn_Cadastrar.Visible = False
    
    ' ~~{ Seleciona as Caixas de Texto }~~
    Me.txt_NumBem.SetFocus
    Me.txt_NumBem.SelText = ""
    
    ' ~~{ Limpar caixa de texto }~~
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
    txt_Valor = ""
End Sub

' Inicio da macro privada 'locked_text()'
Private Sub locked_text()
    ' Declarando variavel tipo boolean
    Dim op As Boolean
    op = False
    
    If txt_NumBem.Value = "" Then
        ' ~~{ Se a caixa de texto for igual a nulo, o op resebera o valor verdadeiro }~~
        op = True
    End If
    
    'mudando o estatus de bloqueado
    txt_NumBem.Locked = Not op
    txt_Grupo.Locked = op
    txt_DescrBem.Locked = op
    txt_Cor.Locked = op
    txt_Marca.Locked = op
    txt_Modelo.Locked = op
    txt_NumSala.Locked = op
    txt_NumSerie.Locked = op
    txt_Local.Locked = op
    txt_Processo.Locked = op
    opt_Ativo.Locked = op
    opt_Desativado.Locked = op
    txt_DataCadas.Locked = op
    txt_Valor.Locked = op
    btn_Cadastrar.Visible = Not op
    
End Sub

' Inicio do evento 'btn_Voltar_Click'
Private Sub btn_Voltar_Click()
    ' ~~{ Fecha o formulario e Seleciona a Tabela Inicial }~~
    Unload FRM_Cadastro
    Sheets("HOME").Select
End Sub

' Inicio do evento 'txt_DataCadas_Change'
Private Sub txt_DataCadas_Change()
    ' Para poder armazenar  10 digitos, Ex:. 01/02/2023
    txt_DataCadas.MaxLength = 10

    If Len(txt_DataCadas.Text) = 2 Then
        ' ~~{ Neste If, ele adiciona uma "/" Quando for digitado 2 digitos }~~
        txt_DataCadas = txt_DataCadas + "/"
        ' Ex:. 01/
    End If

    If Len(txt_DataCadas.Text) = 5 Then
        ' ~~{ Neste If, ele adiciona uma "/" Quando for digitado +2 digitos }~~
        txt_DataCadas = txt_DataCadas + "/"
        ' Ex:. 01/02/ Contando com os outros 3 digitos anteriores
    End If
        
    If Len(txt_DataCadas.Text) = 10 Then
        ' ~~{ Neste If, ele adiciona uma "/" Quando for digitado +4 digitos }~~
        Application.SendKeys "<TAB>"
        ' Emula o Tab, Ex:. 01/02/2023 <TAB>
    End If
End Sub

' Inicio do evento 'txt_NumBem_AfterUpdate'
Private Sub txt_NumBem_AfterUpdate()
    ' ~~{ Declarando variavel }~~
    Dim i As Integer
    Dim op As String
        
    'linha' recebe o numero de linha planilia
    linha = Sheets("Patrimonio").Range("B2").End(xlDown).Row
    numBem = txt_NumBem.Value
    
    ' ~~{ Tratamento de Erro: Se for diferente de "" faca }~~
    If Not txt_NumBem.Value = "" Then
            
        Call locked_text
        ' ~~{ Inicio do Loop For: Faça ate que o 'i' maior ou igual a 'linha' }~~
        For i = 3 To linha
        
            op = Sheets("Patrimonio").Range("B" & i).Value
            
            ' ~~{ Tratamento de ERRO }~~
            If Me.txt_NumBem.Value = op Then
            
                ' --{ Exibe os Dados nas Caixas de Texto }--
                txt_Grupo.Value = Sheets("Patrimonio").Range("C" & i).Value
                txt_DescrBem.Value = Sheets("Patrimonio").Range("D" & i).Value
                txt_Cor.Value = Sheets("Patrimonio").Range("E" & i).Value
                txt_Marca.Value = Sheets("Patrimonio").Range("F" & i).Value
                txt_Modelo.Value = Sheets("Patrimonio").Range("G" & i).Value
                txt_NumSala.Value = Sheets("Patrimonio").Range("H" & i).Value
                txt_NumSerie.Value = Sheets("Patrimonio").Range("I" & i).Value
                txt_Local.Value = Sheets("Patrimonio").Range("J" & i).Value
                
                ' ~~{ Se o dado da tabela for igual a 'Ativo' faca }~~
                If Sheets("Patrimonio").Range("L" & i).Value = "Ativo" Then
                    opt_Ativo.Value = True
                Else
                    ' ~~{ Se o dado da tabela for igual a 'Desativado' faca }~~
                    If Sheets("Patrimonio").Range("L" & i).Value = "Desativado" Then
                        opt_Desativado.Value = True
                    End If
                    
                End If
    
                txt_Processo.Value = Sheets("Patrimonio").Range("K" & i).Value
                txt_DataCadas.Value = Sheets("Patrimonio").Range("M" & i).Value
                txt_Valor.Value = Sheets("Patrimonio").Range("N" & i).Value
                
                ' --{ Desbloqueia os outras caixas de texto editaveis }--
                
                Exit For
                ' Break do LOOP FOR
                
            End If
        ' Proximo LOOP i++
        Next
        
        ' ~~{ Se o 'i' for maior que linha e 'txt_NumBem' for diferente de "" faca }~~
        If i > linha Then
            ' --{ Mensagem exibida ao usuario sobre cadastro não existente }--
            msg = MsgBox("Patrimônio não existente, deseja cadastrar", _
                vbYesNo + vbInformation, "ATENÇÃO")
             
            If msg = 6 Then
                ' ~~{ Se msg for = 6 faca }~~
                ' ~~{ Desbloqueia os outras caixas de texto editaveis }~~
                Call locked_text
                ' Torna o botão invisivel
            Else
                ' --{ Mensagem exibida ao usuario sobre cadastro não existente }--
                btn_Cadastrar.Visible = False
                msg = MsgBox("Digite o número ou escaneie o codigo do patrimônio novamento", _
                    vbOKOnly + vbInformation, "Atenção")
            
            End If
        End If
    End If
End Sub

' Inicio da macro 'Formatar(ByVal linha2 As Integer)'
Sub Formatar(ByVal linha2 As Integer)
    
    Sheets("Patrimonio").Select
    ' Seleciona a planilia
    Range("A3:N3").Select
    ' Seleciona a primeira linha da planilia
    
    Selection.Copy
    ' Copia a formatação da linha
    Range("A" & linha2).Select
    ' Seleciona a linha editada
    
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        ' Cola a formatação da primeira linha
    Application.CutCopyMode = False
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
    FRM_Cadastro.Caption = "{ Cadastrar Patrimônio }"
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
    FRM_Cadastro.Caption = "{ Cadastrar Patrimônio: Click para reajustar! }"
End Sub
