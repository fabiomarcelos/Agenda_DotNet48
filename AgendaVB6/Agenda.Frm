VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Agenda Eletr�nica - Conhecendo e Trabalhando com o VB-C�rtes & Scherr"
   ClientHeight    =   6030
   ClientLeft      =   2220
   ClientTop       =   1470
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6030
   ScaleWidth      =   7575
   Begin VB.CommandButton Nome_Completo_Btn 
      BackColor       =   &H0080C0FF&
      Caption         =   "Nome Completo"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton Nome_Iniciando_Btn 
      BackColor       =   &H0080C0FF&
      Caption         =   "Nomes Iniciando Com"
      Height          =   375
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton Letras_Nome_Btn 
      BackColor       =   &H0080C0FF&
      Caption         =   "Letras do Nome"
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton Refresh_Btn 
      BackColor       =   &H000000FF&
      Caption         =   "Atualizar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   5040
      Width           =   1335
   End
   Begin VB.TextBox Text9 
      DataField       =   "CEP"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   5640
      TabIndex        =   26
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0FFC0&
      Caption         =   "�ltimo"
      Height          =   375
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Pr�ximo"
      Height          =   375
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Anterior"
      Height          =   375
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Primeiro"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0080C0FF&
      Caption         =   "Fim"
      Height          =   375
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      DataField       =   "Data de Nascimento"
      DataSource      =   "Data1"
      Height          =   405
      Left            =   5880
      TabIndex        =   19
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      DataField       =   "Fax"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   4560
      TabIndex        =   17
      Top             =   1440
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Exclus�o "
      Height          =   375
      Left            =   3120
      MaskColor       =   &H8000000A&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Inclus�o "
      Height          =   375
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Pesquisa"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      DataField       =   "Endere�oNoCorreioEletr�nico"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   1200
      TabIndex        =   10
      Top             =   3480
      Width           =   2775
   End
   Begin VB.TextBox Text6 
      DataField       =   "EstadoOuProv�ncia"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   4920
      TabIndex        =   9
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox Text5 
      DataField       =   "Cidade"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      DataField       =   "Endere�o"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   2040
      Width           =   4815
   End
   Begin VB.TextBox Text2 
      DataField       =   "Telefone"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   1440
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      DataField       =   "Nome"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   840
      Width           =   3735
   End
   Begin VB.Label Label10 
      Caption         =   "CEP"
      Height          =   255
      Left            =   4920
      TabIndex        =   27
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label Label9 
      Caption         =   "Data Nasc."
      Height          =   495
      Left            =   5280
      TabIndex        =   18
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Fax"
      Height          =   375
      Left            =   4080
      TabIndex        =   16
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Agenda Eletr�nica - Banco de Dados com Acesso Direto"
      Height          =   255
      Left            =   2040
      TabIndex        =   15
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label7 
      Caption         =   "E-Mail "
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Estado"
      Height          =   255
      Left            =   4200
      TabIndex        =   8
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "Cidade"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Endere�o"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Telefone"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Nome"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
   'Define uma vari�vel global como
   'sendo do tipo Database. A vari�vel
   'Agenda ser� utilizada para indicar o
   'nome e o caminho (drive e diret�rio)
   'do arquivo de banco de dados. No nosso
   'caso, utilizaremos o arquivo agenda.mdb
   'localizado em c:\livro.vb\agenda4\
   Public Agenda As Database
   
   'Define uma vari�vel global como
   'sendo do tipo Recordset. Neste programa
   'o Recordset Tabela_Agenda ser� utilizado
   'para referenciar a tabela utilizada.
   'O arquivo agenda.mdb possui uma tabela
   'denominada Agenda_Eletronica. Lembre-se que
   'um arquivo do tipo Access pode possuir mais
   'de uma tabela.
   Public Tabela_Agenda As Recordset
   
   'Este Recordset ser� utilizado em
   'pesquisas de dados utilizando
   'estruturas em SQL
   Public Pesquisa As Recordset
       
   
   Function Atualiza_Campos_Pesquisa_SQL()
   
    'Atualizamos as caixas de texto com os campos
    'do banco de dados. Antes, por�m, verificamos
    'se os mesmos realment armazenam dados
    'atrav�s do teste da propriedade Value > 0 .
    'Caso n�o realizassemos esse teste, o programa
    'poderia dar erro ao tentarmos passar um valor
    'nulo (campo sem dado) a uma caixa de texto.
    'Lembre-se que nem todos os campos s�o de
    'preenchimento obrigat�rio, o que pode levar
    '� ocorr�ncia desse erro.
   
   If Pesquisa("Nome").Value > 0 Then
        Text1.Text = Pesquisa("Nome")
    Else
        Text1.Text = ""
    End If
    
    If Pesquisa("Telefone").Value > 0 Then
        Text2.Text = Pesquisa("Telefone")
    Else
        Text2.Text = ""
    End If
    
    If Pesquisa("Endere�o").Value > 0 Then
        Text3.Text = Pesquisa("Endere�o")
    Else
        Text3.Text = ""
    End If
     
    If Pesquisa("Fax").Value > 0 Then
        Text4.Text = Pesquisa("Fax")
    Else
        Text4.Text = ""
    End If
            
    If Pesquisa("Cidade").Value > 0 Then
        Text5.Text = Pesquisa("Cidade")
    Else
        Text5.Text = ""
    End If
    
    If Pesquisa("EstadoOuProv�ncia").Value > 0 Then
        Text6.Text = Pesquisa("EstadoOuProv�ncia")
    Else
        Text6.Text = ""
    End If
    
    If Pesquisa("Endere�oNoCorreioEletr�nico").Value > 0 Then
        Text7.Text = Pesquisa("Endere�oNoCorreioEletr�nico")
    Else
        Text7.Text = ""
    End If
    
    If Pesquisa("Data de Nascimento").Value > 0 Then
        Text8.Text = Pesquisa("Data de Nascimento")
    Else
        Text8.Text = ""
    End If
    
    If Pesquisa("Cep").Value > 0 Then
        Text9.Text = Pesquisa("Cep")
    Else
        Text9.Text = ""
    End If
   
   End Function


    
  
   Function Atualiza_Campos()
       
    'Para cada campo da tabela � verificado
    'o valor do campo. Caso o valor seja
    'maior do que zero, a caixa de texto
    'apresenta o conte�do do campo especificado
    'entre parenteses. Se o valor n�o for maior
    'do que zero, o conte�do da caixa de texto
    '� "zerado". Esse teste � necess�rio pois
    'os campos da tabela "Agenda_Eletronica" n�o
    's�o de preenchimento obrigat�rio. Dessa forma,
    'se tentarmos atualizar uma caixa de texto
    'utilizando um campo nulo, o programa acusar� erro.
    'O mesmo teste � feito para cada uma dos campos.
        
    
    If Tabela_Agenda("Nome").Value > 0 Then
        Text1.Text = Tabela_Agenda("Nome")
    Else
        Text1.Text = ""
    End If
    
    If Tabela_Agenda("Telefone").Value > 0 Then
        Text2.Text = Tabela_Agenda("Telefone")
    Else
        Text2.Text = ""
    End If
    
    If Tabela_Agenda("Endere�o").Value > 0 Then
        Text3.Text = Tabela_Agenda("Endere�o")
    Else
        Text3.Text = ""
    End If
     
    If Tabela_Agenda("Fax").Value > 0 Then
        Text4.Text = Tabela_Agenda("Fax")
    Else
        Text4.Text = ""
    End If
            
    If Tabela_Agenda("Cidade").Value > 0 Then
        Text5.Text = Tabela_Agenda("Cidade")
    Else
        Text5.Text = ""
    End If
    
    If Tabela_Agenda("EstadoOuProv�ncia").Value > 0 Then
        Text6.Text = Tabela_Agenda("EstadoOuProv�ncia")
    Else
        Text6.Text = ""
    End If
    
    If Tabela_Agenda("Endere�oNoCorreioEletr�nico").Value > 0 Then
        Text7.Text = Tabela_Agenda("Endere�oNoCorreioEletr�nico")
    Else
        Text7.Text = ""
    End If
    
    If Tabela_Agenda("Data de Nascimento").Value > 0 Then
        Text8.Text = Tabela_Agenda("Data de Nascimento")
    Else
        Text8.Text = ""
    End If
    
    If Tabela_Agenda("Cep").Value > 0 Then
        Text9.Text = Tabela_Agenda("Cep")
    Else
        Text9.Text = ""
    End If
   
   End Function
     
    
Function Limpa_Campos()
    
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""
    Text9.Text = ""

End Function

Private Sub Command1_Click()
 'Faz a pesquisa por Nome no Banco de Dados.
 'Se n�o localizar o registro,
 'mostra a mensagem de "N�o Existe".
 
 'O padr�o de pesquisa � o mesmo utilizado para
 'o objeto Data, apenas com a seguinte diferen�a:
 'quando utilizamos o objeto Data a pesquisa foi
 'realizada com  Data1.Recordset.FindFirst .
 'Agora, utilizamos Tabela_Agenda.FindFirst ,
 'lembrando que Tabela_Agenda foi definida como
 'Recordset em General Delcarations.
 'Como objeto Data � utilizado um Recordset
 'associado a esse componente, ao passo que nesta
 'vers�o da agenda n�s definimos um Recordset.
 
 Dim Dado As String, Busca As String
 
 'Recebe o nome completo a ser pesquisado.
 Dado = InputBox("Informe o Nome Completo", " Nome ")
 
 'Por exemplo: se desejarmos pesquisar o nome
 'Ramon, devemos ter informado esse nome
 'quando da execu��o do comando anterior InputBox.
 'O nome Ramon ficar� armazenado na vari�vel
 'Dado. Em seguida, utilizaremos essa vari�vel para
 'montar o string de pesquisa (Busca), o qual ficar�
 'assim: " Nome='Ramon' " (veja pr�ximo comando).
 Busca = " Nome='" & Dado & "'"
 
 'Em seguida, pediremos para pesquisar a primeira
 'ocorr�ncia do nome Ramon no campo Nome
 'da Tabela_Agenda, conforme o pr�ximo comando, o
 'qual poderia ser interpretado da seguinte maneira
 'Tabela_Agenda.FindFirst Nome='Ramon' ou
 'ent�o: na tabela Agenda_Eletronica (indicado pelo
 'Recordset Tabela_Agenda) procure a primeira
 'ocorr�ncia (FindFirst) de Ramon no campo
 'Nome (Nome='Ramon').
 Tabela_Agenda.FindFirst Busca
 
 'Caso ele n�o encontre o nome especificado
 'ele d� uma mensagem de eroo.
 If Tabela_Agenda.NoMatch Then
    MsgBox (" O Nome " & Dado & " n�o existe! ")
 Else
    'Caso ele encontre o nome, os
    'campos s�o atualizados.
    Atualiza_Campos
 End If
    
End Sub

Private Sub Command10_Click()

    'Vai para o �ltimo registro v�lido
    Tabela_Agenda.MoveLast
        
    Atualiza_Campos
 
End Sub


Private Sub Command2_Click()

    Dim Dado As String, Texto As String
    
    'Solicita ao usu�rio que informe
    'o novo nome a ser cadastrado.
    Texto = "Informe o Nome e depois entre com os dados." & _
            Chr$(13) & "Depois, clique em Atualizar."
    Dado = InputBox(Texto, "Nome")
    
    'Cria um novo registro utilizando AddNew
    Tabela_Agenda.AddNew
    
    'O nome informado � colocado na caixa
    'de texto Text1. Em seguida, o usu�rio
    'dever� colocar os restantes dos dados
    'nas outras caixas de texto. Todos esses
    'novos dados ser�o incorporados ao
    'registro rec�m criado
    Text1.Text = Dado
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""
    Text9.Text = ""
    
    'Habilita o bot�o Refresh_Btn
    '(Atualizar) para que o usu�rio
    'possa acion�-lo ap�s preencher
    'os campos (caixa de texto) do
    'formul�rio.
    Refresh_Btn.Enabled = True
    
    

End Sub

Private Sub Command3_Click()

    'Exclui o registro, sendo que antes da exclus�o,
    'confirma a exclus�o com o usu�rio
    'atrav�s do m�todo MSGBOX.

    Dado = Text1.Text
    
    'Quando utilizamos o MsgBox, logo a seguir, o n�mero 4
    'chama a aten��o dos programadores. Ao utilizarmos esse
    'n�mero, estamos informando que desejamos um quadro com
    'os bot�es Yes(Sim) / No(N�o). Caso utiliz�ssemos o
    'n�mero 0, apareceria apenas um bot�o OK. Com o n�mero 1
    'apareceria os bot�es OK e Cancel, e assim por diante.
    Choice = MsgBox(" Excluir o Nome " & Dado & " ? ", 4, "Excluir")
    
    'A estrutura a seguir, serve para identificar qual foi
    'o bot�o selecionado no MsgBox anteriormente executado.
    'O n�mero 6 � o valor retornado pelo bot�o Yes.
    'Conseq�entemente, estamos testanto se o usu�rio
    'acionou esse bot�o no MsgBox.
    If Choice = 6 Then
         Tabela_Agenda.Delete
    End If
    
    'A seguir, procuramos ir para o pr�ximo registro.
    'Caso isso n�o seja poss�vel, vamos para o �ltimo
    'registro v�lido. Essa rotina � melhor explicada
    'no bot�o que leva ao pr�ximo registro.
 
    Tabela_Agenda.MoveNext
    
    If Tabela_Agenda.EOF Then
        Tabela_Agenda.MoveLast
    End If
    
    Atualiza_Campos
 

End Sub

Private Sub Command4_Click()
    
    'Imprime o conte�do das caixas de texto.
    
    Printer.Print Text1.Text
    Printer.Print Text2.Text
    Printer.Print Text3.Text
    Printer.Print Text4.Text
    Printer.Print Text5.Text
    Printer.Print Text6.Text
    Printer.Print Text7.Text
    Printer.Print Text8.Text
    Printer.Print Text9.Text
    
    Printer.EndDoc
    

End Sub

Private Sub Command5_Click()
    
    'Finaliza o programa.
    
    End
    
End Sub

Private Sub Command6_Click()
    
    'Nos leva ao primeiro registro
    'da tabela amarrada pela
    'vari�vel Tabela_Agenda.
    'Essa vari�vel foi declarada
    'em General - Declarations.
    Tabela_Agenda.MoveFirst
                
    'Chama a fun��o que
    'atualiza os campos
    '(ou conte�do das
    'caixas de texto)
    Atualiza_Campos
    
 End Sub

Private Sub Command8_Click()
    
    'Nos leva ao registro anterior da
    'tabela especificada (ou amarrada)
    '� vari�vel Tabela_Agenda, a qual
    'foi declarada em General-Declarations
    Tabela_Agenda.MovePrevious
    
    'Testa se atingimos um ponto antes
    'do in�cio dos registros. BOF significa
    'Beginning Of File, que pode ser traduzido
    'por Antes do Arquivo. Ao atingirmos essa
    'posi��o, n�o poderemos atualizar os campos,
    'dessa forma, utilizamos MoveFirst para ir
    'ao primeiro (First) registro v�lido.
    If Tabela_Agenda.BOF Then
        Tabela_Agenda.MoveFirst
    End If
         
    'Caso tent�ssemos atualizar os campos (caixas
    'do tipo texto)sem a verifica��o anterior e
    'estiv�ssemos em BOF, o programa dar�a erro.
    'A seguir, � chamada a fun��o para atualiza��o
    'dos campos.
    Atualiza_Campos
  
End Sub

Private Sub Command9_Click()

    'Vamos para o pr�ximo registro da
    'tabela amarrada � vari�vel Tabela_Agenda.
    Tabela_Agenda.MoveNext
    
    'Verificamos se foi atingido o final
    'do arquivo (EOF que significa
    'End Of File e pode ser traduzido
    'exatamente como Final de Arquivo).
    'Caso essa posi��o tenha sido atingida,
    'vamos para o �ltimo registro v�lido,
    'utilizando MoveLast.
    If Tabela_Agenda.EOF Then
        Tabela_Agenda.MoveLast
    End If
        
    'O pr�ximo comando executa a rotina que
    'atualiza as caixas de texto do formul�rio
    'com o conte�do do registro atual resultante
    'de uma das opera��es anteriores (MoveNext ou
    'MoveLast). Caso seja atingido a posi��o EOF
    'e fosse solicitada a atualiza��o dos campos,
    'o programa acusaria um erro, pois o registro
    'n�o seria v�lido. Com o teste anterior,
    'fica contornado o problema, sendo apresentado
    'um registro v�lido.
    Atualiza_Campos
 
End Sub

Private Sub Form_Load()
    
 'Quando o formul�rio � carregado, o arquivo
 'c:\livro.vb\agenda4\agenda.mdb � lido e
 'amarrado � vari�vel Agenda, conforme indicado
 'a seguir. Essa vari�vel foi declarada em
 'General - Declarations. O par�mtero False indica que
 'N�O estamos abrindo o banco de dados apenas para leitura
 'ou consulta. Dessa forma, poderemos atualizar os dados
 'apresentados ou incluir outros registros.
 Set Agenda = OpenDatabase(App.Path & "\agenda.mdb", _
                            False)
 
 
 
 
 'Ap�s a abertura do arquivo, vamos amarrar a
 'vari�vel Tabela_Agenda �. Tabela Agenda_Eletronica.
 'A op��o dbOpenDynaset possibilita a abertura
 'de uma tabela para adi��o, edi��o e exclus�o de dados.
 'Essa vari�vel foi declarada em General - Declarations.
 Set Tabela_Agenda = Agenda.OpenRecordset("Agenda_Eletronica", _
                     dbOpenDynaset)
     
 'Levamos a tabela para o
 'primeiro registro utilizando o
 'MoveFirst, mostrado a seguir.
 Tabela_Agenda.MoveFirst
 
 Atualiza_Campos
 
 Refresh_Btn.Enabled = False
    
End Sub

Private Sub Letras_Nome_Btn_Click()
  'Faz a pesquisa por Nome no Banco de Dados,
  'utilizando instru��o SQL. Se n�o localizar
  'o registro, mostra a mensagem de "N�o Existe".
  
  Dim Dado As String, Busca As String, Mensagem As String
  Dim Consulta_SQL As String, Resultado As String
  Dim Contador As Integer, Total_encontrado As Integer
  
  'As vari�veis Contador e Total_encontrado ser�o
  'utiizadas para mostrar os registros encontrados
  'quando da realiza��o da pesquisa. A seguir,
  'essas vari�veis s�o zeradas.
  Contador = 0
  Total_encontrado = 0
                 
  'Montamos a mensagem a ser apresentada ao usu�rio, orientando-o.
  'A utiliza��o de Chr$ (13) faz com que o texto em seguida seja
  'apresentado em uma nova linha. O caractere _ indica que o comando
  'continua na pr�xima linha.
  Mensagem = "Informe uma letra ou uma seq��ncia de letras do nome." & _
             Chr$(13) & "Por exemplo: ao especificar ana voc� estar�" & _
             Chr$(13) & "pesquisando Ana, Mariana, Juliana, etc."
             
  
  'Apresentamos a mensagem na tela e recebemos a resposta do usu�rio,
  'a qual � armazenada na String Dado (ela ser� utilizada em seguida
  'para a montagem de uma instru��o SQL.
  Dado = InputBox(Mensagem)
  
  'A seguir montamos a nossa instru��o SQL. A sua estrutura � facil de
  'ser entendida e pode ser "traduzida" da seguinte forma:
  'SELECIONE campo Nome DA Agenda_Eletr�nica ONDE Campo Nome IGUAL a
  '*Dado* ORDENADOS PELO Nome. O aster�stico ao redor de DADO � utilizado
  'para indicar que as letras especificadas pelo usu�rio (armazenadas na
  'vari�vel Dado) podem conter outras letras antes (*Dado) ou depois (Dado*).
  'A ordena��o pelo nome � necess�ria para organizar os dados e a pesquisa.
  'A instru��o a seguir foi colocada como um coment�rio pois essa pesquisa
  'seleciona apenas o campo Nome, n�o fornecendo resultados dos outros
  'campos da tabela Agenda_Eletronica.
  'Consulta_SQL="SELECT Nome FROM Agenda_Eletronica WHERE Nome LIKE '*" _
                    + Dado + "*' ORDER BY Nome;"
  
                    
  'No programa foi utilizada a instru��o a seguir. Note que ela
  'apresenta uma diferen�a em rela��o � instru��o anterior (colocada
  'como coment�rio): em lugar do campo Nome, utilizamos um aster�stico
  '(*), o qual possibilita � sele��o de todos os campos de nossa tabela.
  Consulta_SQL = "SELECT * FROM Agenda_Eletronica WHERE Nome LIKE '*" _
                  + Dado + "*' ORDER BY Nome;"
    
  
  'A instru��o SQL � aplicada, atrav�s da abertura de um Recordset.
  'O resultado � armazenado na vari�vel Pesquisa, declarada em
  'General - Declarations.
  Set Pesquisa = Agenda.OpenRecordset(Consulta_SQL, dbOpenSnapshot)
  
  
  '� montada uma mensagem com o n�mero de registros
  'que satisfizeram a pesquisa anterior.
  Mensagem = Consulta_SQL + (Chr(13) & Chr(10)) + _
  "Apresentou total de registros = " + Str(Pesquisa.RecordCount)
  
  'Essa mensagem � apresentada ao usu�rio.
  Resultado = MsgBox(Mensagem, vbInformation)
          
  'A seguir, a vari�vel Total_encontrado recebe
  'o total de registros encontrados.
  Total_encontrado = Pesquisa.RecordCount
  
  'Limpa as caixas de texto,
  'executando a fun��o a seguir
  Limpa_Campos
     
  
  'O loop Do...Loop Until � utilizado para mostrar os
  'registros encontrados.
  Do
      '� feito um teste para verificar a posi��o
      'absoluta do registro que ser� mostrado.
      'Esse teste � apenas uma precau��o para que
      'n�o se tente mostrar um registro inexistente.
      If Pesquisa.AbsolutePosition > -1 Then
          
          'Atualiza o conte�do das caixas de texto
          'do nosso formul�rio, utilizando o Recordset
          'Pesquisa (declarado em General - Declarations).
          'Lembre-se que Atualiza_Campos_Pesquisa_SQL
          '� uma fun��o definida em General, a qual pode
          'ser chamada a partir de outras procedures.
          Atualiza_Campos_Pesquisa_SQL
            
          'Montamos uma mensagem mostrando o nome encontrado.
          Mensagem = Str(Contador + 1) + "� Registro encontrado: " + _
          Pesquisa("nome")
          
          'Essa mensagem � apresentada ao usu�rio.
          Resultado = MsgBox(Mensagem, vbInformation)
          
          'A vari�vel Contador (zerada no in�cio desta procedure)
          '� incrementada. Ela est� sendo utilizada para contar o
          'total de registros apresentados aos usu�rios.
          Contador = Contador + 1
      
          'Como a pesquisa SQL foi ordenada por Nome
          '(veja o trecho final da instru��o SQL ORDER BY Nome;)
          'agora vamos para o pr�ximo registro,
          'utilizando a instru��o MoveNext.
          Pesquisa.MoveNext
      
      Else
          'Caso n�o seja encontrado um registro v�lido as caixas
          'de texto s�o limpas e uma mensagem indicativa �
          'apresentada na tela.
          Limpa_Campos
          Resultado = MsgBox("Nome n�o encontrado.", vbInformation)
      End If
      
  
  'O loop Do...Loop Until � executado enquanto o
  'Contador for menor ou igual ao total de resgistros
  'encontrados (utilizando a vari�vel Total_encontrado)
  Loop Until Contador >= Total_encontrado
          
End Sub

Private Sub Nome_Completo_Btn_Click()
 'Faz a pesquisa por Nome no Banco de Dados utilizando SQL.
 'Se n�o localizar o registro, mostra a mensagem de "N�o Existe".

 Dim Dado As String, Mensagem As String
 Dim Consulta_SQL As String
 Dim Resultado As String
           
 'Solicita o nome completo a ser pesquisado.
 Dado = InputBox("Informe o Nome Completo")
 
 'Monta o comando SQL e o armazena na String Consulta_SQL
 Consulta_SQL = "SELECT * FROM Agenda_Eletronica " + _
                 "WHERE Nome = '" + Dado + "';"
 'O comando SQL anterior pode ser traduzido como
 'SELECIONE *(todos os registros) DA Agenda_Eletronica
 'ONDE Nome IGUAL nome especificado.
    
 'A vari�vel Pesquisa (definida como sendo do tipo Recordset
 '(no in�cio desta procedure) recebe o resultado da pesquisa
 'SQL montada em Consulta_SQL .
 '
 Set Pesquisa = Agenda.OpenRecordset(Consulta_SQL, dbOpenSnapshot)

 'Caso a propriedade AbsolutePosition seja maior do que -1
 'isso indica que a pesquisa encontrou o nome especificado.
 'Para indicar o valor de AbsolutePosition, montamos a mensagem
 'a seguir e a exibimos em um MsgBox, independente do seu valor.
 Mensagem = "Posi��o retornada por Pesquisa.AbsolutePosition = " + _
           Str(Pesquisa.AbsolutePosition)
           
 Resultado = MsgBox(Mensagem, vbInformation)
  
 'Limpa as caixas de texto,
 'executando a fun��o a seguir
 Limpa_Campos

 'Agora, testamos o valor da propriedade AbsolutePosition
 'a qual nos indica o sucesso ou n�o da pesquisa, conforme
 'comentado anteriormente.
 If Pesquisa.AbsolutePosition > -1 Then
    
    'Mostramos o nome encontrado
    Mensagem = "Nome encontrado: " + Pesquisa("nome")
    Resultado = MsgBox(Mensagem, vbInformation)
    
    Atualiza_Campos_Pesquisa_SQL
    
 Else
 
    Resultado = MsgBox("Nome n�o encontrado.", vbInformation)
    
 End If


End Sub

Private Sub Nome_Iniciando_Btn_Click()
 'Faz a pesquisa pela primeira parte do nome
 'no Banco de Dados utilizando uma instru��o em SQL. Por exemplo,
 'caso ele deseje verificar todas as pessoas com o nome Mariana
 '(Mariana Lima, Mariana Fernandez, Mariana Silveira, etc.),
 'Basta digitar apenas Mariana, que todos os registros com esse
 'primeiro nome ser�o mostrados. Se n�o localizar o registro,
 'mostra a mensagem de "N�o Existe". Os procedimentos s�o
 'muito similares aos apresentados para a procedure de pesquisa
 'de nomes com determinadas letras (Letras_Nome_Btn_Click).
 'A diferen�a est� na instru��o SQL (Dado + "*' ORDER BY Nome;)
 'a qual apresenta aster�stico apenas de um dos lados do
 'nome especificado pelo usu�rio. Por exemplo, ao pesquisar
 'o nome Paulo, a instru��o SQL ir� procurar os nomes Paulo*.
 'Lembre-se de que o aster�stico funciona como um caractere
 '(ou conjunto de caracteres) coringa, mais ou menos como
 'acontece com o MS-DOS, quando executamos o comando
 'Dir P* o qual ir� mostrar todos os programas que come�am
 'com a letra P (no diret�rio ativo).
 
 
 Dim Dado As String, Busca As String, Mensagem As String
 Dim Consulta_SQL As String, Resultado As String
 Dim Contador As Integer, Total_encontrado As Integer
 
     
 Contador = 0
 Total_encontrado = 0
                
 Dado = InputBox("Informe a Primeira Parte do Nome ")
 
 'Verifique a instru��o SQL montada a seguir:
 Consulta_SQL = "SELECT * FROM Agenda_Eletronica WHERE Nome LIKE '" + Dado + "*' ORDER BY Nome;"
 'Ela pode ser "traduzida" como da seguinte forma:
 'SELECIONE todos os campos (*) DA Agenda_Eletronica ONDE Nome IGUAL
 'nome especificado (seguido de aster�stico).
    
 
 Set Pesquisa = Agenda.OpenRecordset(Consulta_SQL, dbOpenSnapshot)
 
 Mensagem = Consulta_SQL + (Chr(13) & Chr(10)) + _
 "Apresentou total de registros = " + Str(Pesquisa.RecordCount)
 
 Resultado = MsgBox(Mensagem, vbInformation)
         
 Total_encontrado = Pesquisa.RecordCount
     
 
 'A seguir, os registros encontrados s�o
 'apresentados, um a um.
 Do
     If Pesquisa.AbsolutePosition > -1 Then
     
         Atualiza_Campos_Pesquisa_SQL
         
         Mensagem = Str(Contador + 1) + "� Registro encontrado: " + _
         Pesquisa("nome")
         
         Resultado = MsgBox(Mensagem, vbInformation)
         
         Contador = Contador + 1
     
         Pesquisa.MoveNext
     
     Else
         Resultado = MsgBox("Nome n�o encontrado.", vbInformation)
     End If
 
 Loop Until Contador >= Total_encontrado
 
End Sub

Private Sub Refresh_Btn_Click()
    
    'Desabilita o bot�o Refresh (Atualizar)
    Refresh_Btn.Enabled = False
        
    'Verifica o texto em cada caixa de texto.
    'Caseo seja igual a zero "", � acrescentado
    'um espa�o em branco, simplesmente para que
    'n�s possamos atribuir esses conte�dos aos
    'campos do banco de dados. No caso do
    'conte�do de uma caixa de texto ser igual
    'a zero "" e n�s tentarmos atribu�-lo a
    'um campo do banco de dados, ocorreria um
    'erro.
    
    If Text1.Text = "" Then
       Text1.Text = " "
    End If
        
    If Text2.Text = "" Then
       Text2.Text = " "
    End If
    
    If Text3.Text = "" Then
       Text3.Text = " "
    End If
    
    If Text4.Text = "" Then
       Text4.Text = " "
    End If
    
    If Text5.Text = "" Then
       Text5.Text = " "
    End If
    
    If Text6.Text = "" Then
       Text6.Text = " "
    End If
    
    If Text7.Text = "" Then
       Text7.Text = " "
    End If
    
    'No caso data de nascimento, como ela �
    'um campo espec�fico, n�o basta simplesmente
    'colocarmos um espa�o em branco para que
    'a atribui��o seja poss�vel. Nesse caso, temos
    'de colocar a data correta, motivo pelo qual
    'solicitamos isso ao usu�rio no caso do campo
    'estar vazio.
    
    If Text8.Text = "" Then
       Texto = "Voc� deve digitar a data " & Chr$(13) & _
                "de nascimento para prosseguir:"
       Text8.Text = InputBox(Texto, "Data de Nascimento")
    End If
    
    If Text9.Text = "" Then
       Text9.Text = " "
    End If
    
    'Faz a atribui��o do conte�do das
    'caixas de texto aos respectivos
    'campos no banco de dados.
    Tabela_Agenda("Nome") = Text1.Text
    Tabela_Agenda("Telefone") = Text2.Text
    Tabela_Agenda("Endere�o") = Text3.Text
    Tabela_Agenda("Fax") = Text4.Text
    Tabela_Agenda("Cidade") = Text5.Text
    Tabela_Agenda("EstadoOuProv�ncia") = Text6.Text
    Tabela_Agenda("Endere�oNoCorreioEletr�nico") = Text7.Text
    Tabela_Agenda("Data de Nascimento") = Text8.Text
    Tabela_Agenda("Cep") = Text9.Text
           
    'Finalmente, atualiza o banco
    'de dados utilizando Update
    Tabela_Agenda.Update
       
    
End Sub
