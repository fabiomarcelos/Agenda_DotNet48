VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Agenda Eletrônica - Conhecendo e Trabalhando com o VB-Côrtes & Scherr"
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
      Caption         =   "Último"
      Height          =   375
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Próximo"
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
      Caption         =   "Exclusão "
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
      Caption         =   "Inclusão "
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
      DataField       =   "EndereçoNoCorreioEletrônico"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   1200
      TabIndex        =   10
      Top             =   3480
      Width           =   2775
   End
   Begin VB.TextBox Text6 
      DataField       =   "EstadoOuProvíncia"
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
      DataField       =   "Endereço"
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
      Caption         =   "Agenda Eletrônica - Banco de Dados com Acesso Direto"
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
      Caption         =   "Endereço"
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
 
   'Define uma variável global como
   'sendo do tipo Database. A variável
   'Agenda será utilizada para indicar o
   'nome e o caminho (drive e diretório)
   'do arquivo de banco de dados. No nosso
   'caso, utilizaremos o arquivo agenda.mdb
   'localizado em c:\livro.vb\agenda4\
   Public Agenda As Database
   
   'Define uma variável global como
   'sendo do tipo Recordset. Neste programa
   'o Recordset Tabela_Agenda será utilizado
   'para referenciar a tabela utilizada.
   'O arquivo agenda.mdb possui uma tabela
   'denominada Agenda_Eletronica. Lembre-se que
   'um arquivo do tipo Access pode possuir mais
   'de uma tabela.
   Public Tabela_Agenda As Recordset
   
   'Este Recordset será utilizado em
   'pesquisas de dados utilizando
   'estruturas em SQL
   Public Pesquisa As Recordset
       
   
   Function Atualiza_Campos_Pesquisa_SQL()
   
    'Atualizamos as caixas de texto com os campos
    'do banco de dados. Antes, porém, verificamos
    'se os mesmos realment armazenam dados
    'através do teste da propriedade Value > 0 .
    'Caso não realizassemos esse teste, o programa
    'poderia dar erro ao tentarmos passar um valor
    'nulo (campo sem dado) a uma caixa de texto.
    'Lembre-se que nem todos os campos são de
    'preenchimento obrigatório, o que pode levar
    'à ocorrência desse erro.
   
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
    
    If Pesquisa("Endereço").Value > 0 Then
        Text3.Text = Pesquisa("Endereço")
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
    
    If Pesquisa("EstadoOuProvíncia").Value > 0 Then
        Text6.Text = Pesquisa("EstadoOuProvíncia")
    Else
        Text6.Text = ""
    End If
    
    If Pesquisa("EndereçoNoCorreioEletrônico").Value > 0 Then
        Text7.Text = Pesquisa("EndereçoNoCorreioEletrônico")
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
       
    'Para cada campo da tabela é verificado
    'o valor do campo. Caso o valor seja
    'maior do que zero, a caixa de texto
    'apresenta o conteúdo do campo especificado
    'entre parenteses. Se o valor não for maior
    'do que zero, o conteúdo da caixa de texto
    'é "zerado". Esse teste é necessário pois
    'os campos da tabela "Agenda_Eletronica" não
    'são de preenchimento obrigatório. Dessa forma,
    'se tentarmos atualizar uma caixa de texto
    'utilizando um campo nulo, o programa acusará erro.
    'O mesmo teste é feito para cada uma dos campos.
        
    
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
    
    If Tabela_Agenda("Endereço").Value > 0 Then
        Text3.Text = Tabela_Agenda("Endereço")
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
    
    If Tabela_Agenda("EstadoOuProvíncia").Value > 0 Then
        Text6.Text = Tabela_Agenda("EstadoOuProvíncia")
    Else
        Text6.Text = ""
    End If
    
    If Tabela_Agenda("EndereçoNoCorreioEletrônico").Value > 0 Then
        Text7.Text = Tabela_Agenda("EndereçoNoCorreioEletrônico")
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
 'Se não localizar o registro,
 'mostra a mensagem de "Não Existe".
 
 'O padrão de pesquisa é o mesmo utilizado para
 'o objeto Data, apenas com a seguinte diferença:
 'quando utilizamos o objeto Data a pesquisa foi
 'realizada com  Data1.Recordset.FindFirst .
 'Agora, utilizamos Tabela_Agenda.FindFirst ,
 'lembrando que Tabela_Agenda foi definida como
 'Recordset em General Delcarations.
 'Como objeto Data é utilizado um Recordset
 'associado a esse componente, ao passo que nesta
 'versão da agenda nós definimos um Recordset.
 
 Dim Dado As String, Busca As String
 
 'Recebe o nome completo a ser pesquisado.
 Dado = InputBox("Informe o Nome Completo", " Nome ")
 
 'Por exemplo: se desejarmos pesquisar o nome
 'Ramon, devemos ter informado esse nome
 'quando da execução do comando anterior InputBox.
 'O nome Ramon ficará armazenado na variável
 'Dado. Em seguida, utilizaremos essa variável para
 'montar o string de pesquisa (Busca), o qual ficará
 'assim: " Nome='Ramon' " (veja próximo comando).
 Busca = " Nome='" & Dado & "'"
 
 'Em seguida, pediremos para pesquisar a primeira
 'ocorrência do nome Ramon no campo Nome
 'da Tabela_Agenda, conforme o próximo comando, o
 'qual poderia ser interpretado da seguinte maneira
 'Tabela_Agenda.FindFirst Nome='Ramon' ou
 'então: na tabela Agenda_Eletronica (indicado pelo
 'Recordset Tabela_Agenda) procure a primeira
 'ocorrência (FindFirst) de Ramon no campo
 'Nome (Nome='Ramon').
 Tabela_Agenda.FindFirst Busca
 
 'Caso ele não encontre o nome especificado
 'ele dá uma mensagem de eroo.
 If Tabela_Agenda.NoMatch Then
    MsgBox (" O Nome " & Dado & " não existe! ")
 Else
    'Caso ele encontre o nome, os
    'campos são atualizados.
    Atualiza_Campos
 End If
    
End Sub

Private Sub Command10_Click()

    'Vai para o último registro válido
    Tabela_Agenda.MoveLast
        
    Atualiza_Campos
 
End Sub


Private Sub Command2_Click()

    Dim Dado As String, Texto As String
    
    'Solicita ao usuário que informe
    'o novo nome a ser cadastrado.
    Texto = "Informe o Nome e depois entre com os dados." & _
            Chr$(13) & "Depois, clique em Atualizar."
    Dado = InputBox(Texto, "Nome")
    
    'Cria um novo registro utilizando AddNew
    Tabela_Agenda.AddNew
    
    'O nome informado é colocado na caixa
    'de texto Text1. Em seguida, o usuário
    'deverá colocar os restantes dos dados
    'nas outras caixas de texto. Todos esses
    'novos dados serão incorporados ao
    'registro recém criado
    Text1.Text = Dado
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""
    Text9.Text = ""
    
    'Habilita o botão Refresh_Btn
    '(Atualizar) para que o usuário
    'possa acioná-lo após preencher
    'os campos (caixa de texto) do
    'formulário.
    Refresh_Btn.Enabled = True
    
    

End Sub

Private Sub Command3_Click()

    'Exclui o registro, sendo que antes da exclusão,
    'confirma a exclusão com o usuário
    'através do método MSGBOX.

    Dado = Text1.Text
    
    'Quando utilizamos o MsgBox, logo a seguir, o número 4
    'chama a atenção dos programadores. Ao utilizarmos esse
    'número, estamos informando que desejamos um quadro com
    'os botões Yes(Sim) / No(Não). Caso utilizássemos o
    'número 0, apareceria apenas um botão OK. Com o número 1
    'apareceria os botões OK e Cancel, e assim por diante.
    Choice = MsgBox(" Excluir o Nome " & Dado & " ? ", 4, "Excluir")
    
    'A estrutura a seguir, serve para identificar qual foi
    'o botão selecionado no MsgBox anteriormente executado.
    'O número 6 é o valor retornado pelo botão Yes.
    'Conseqüentemente, estamos testanto se o usuário
    'acionou esse botão no MsgBox.
    If Choice = 6 Then
         Tabela_Agenda.Delete
    End If
    
    'A seguir, procuramos ir para o próximo registro.
    'Caso isso não seja possível, vamos para o último
    'registro válido. Essa rotina é melhor explicada
    'no botão que leva ao próximo registro.
 
    Tabela_Agenda.MoveNext
    
    If Tabela_Agenda.EOF Then
        Tabela_Agenda.MoveLast
    End If
    
    Atualiza_Campos
 

End Sub

Private Sub Command4_Click()
    
    'Imprime o conteúdo das caixas de texto.
    
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
    'variável Tabela_Agenda.
    'Essa variável foi declarada
    'em General - Declarations.
    Tabela_Agenda.MoveFirst
                
    'Chama a função que
    'atualiza os campos
    '(ou conteúdo das
    'caixas de texto)
    Atualiza_Campos
    
 End Sub

Private Sub Command8_Click()
    
    'Nos leva ao registro anterior da
    'tabela especificada (ou amarrada)
    'à variável Tabela_Agenda, a qual
    'foi declarada em General-Declarations
    Tabela_Agenda.MovePrevious
    
    'Testa se atingimos um ponto antes
    'do início dos registros. BOF significa
    'Beginning Of File, que pode ser traduzido
    'por Antes do Arquivo. Ao atingirmos essa
    'posição, não poderemos atualizar os campos,
    'dessa forma, utilizamos MoveFirst para ir
    'ao primeiro (First) registro válido.
    If Tabela_Agenda.BOF Then
        Tabela_Agenda.MoveFirst
    End If
         
    'Caso tentássemos atualizar os campos (caixas
    'do tipo texto)sem a verificação anterior e
    'estivéssemos em BOF, o programa daría erro.
    'A seguir, é chamada a função para atualização
    'dos campos.
    Atualiza_Campos
  
End Sub

Private Sub Command9_Click()

    'Vamos para o próximo registro da
    'tabela amarrada à variável Tabela_Agenda.
    Tabela_Agenda.MoveNext
    
    'Verificamos se foi atingido o final
    'do arquivo (EOF que significa
    'End Of File e pode ser traduzido
    'exatamente como Final de Arquivo).
    'Caso essa posição tenha sido atingida,
    'vamos para o último registro válido,
    'utilizando MoveLast.
    If Tabela_Agenda.EOF Then
        Tabela_Agenda.MoveLast
    End If
        
    'O próximo comando executa a rotina que
    'atualiza as caixas de texto do formulário
    'com o conteúdo do registro atual resultante
    'de uma das operações anteriores (MoveNext ou
    'MoveLast). Caso seja atingido a posição EOF
    'e fosse solicitada a atualização dos campos,
    'o programa acusaria um erro, pois o registro
    'não seria válido. Com o teste anterior,
    'fica contornado o problema, sendo apresentado
    'um registro válido.
    Atualiza_Campos
 
End Sub

Private Sub Form_Load()
    
 'Quando o formulário é carregado, o arquivo
 'c:\livro.vb\agenda4\agenda.mdb é lido e
 'amarrado à variável Agenda, conforme indicado
 'a seguir. Essa variável foi declarada em
 'General - Declarations. O parâmtero False indica que
 'NÃO estamos abrindo o banco de dados apenas para leitura
 'ou consulta. Dessa forma, poderemos atualizar os dados
 'apresentados ou incluir outros registros.
 Set Agenda = OpenDatabase(App.Path & "\agenda.mdb", _
                            False)
 
 
 
 
 'Após a abertura do arquivo, vamos amarrar a
 'variável Tabela_Agenda à. Tabela Agenda_Eletronica.
 'A opção dbOpenDynaset possibilita a abertura
 'de uma tabela para adição, edição e exclusão de dados.
 'Essa variável foi declarada em General - Declarations.
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
  'utilizando instrução SQL. Se não localizar
  'o registro, mostra a mensagem de "Não Existe".
  
  Dim Dado As String, Busca As String, Mensagem As String
  Dim Consulta_SQL As String, Resultado As String
  Dim Contador As Integer, Total_encontrado As Integer
  
  'As variáveis Contador e Total_encontrado serão
  'utiizadas para mostrar os registros encontrados
  'quando da realização da pesquisa. A seguir,
  'essas variáveis são zeradas.
  Contador = 0
  Total_encontrado = 0
                 
  'Montamos a mensagem a ser apresentada ao usuário, orientando-o.
  'A utilização de Chr$ (13) faz com que o texto em seguida seja
  'apresentado em uma nova linha. O caractere _ indica que o comando
  'continua na próxima linha.
  Mensagem = "Informe uma letra ou uma seqüência de letras do nome." & _
             Chr$(13) & "Por exemplo: ao especificar ana você estará" & _
             Chr$(13) & "pesquisando Ana, Mariana, Juliana, etc."
             
  
  'Apresentamos a mensagem na tela e recebemos a resposta do usuário,
  'a qual é armazenada na String Dado (ela será utilizada em seguida
  'para a montagem de uma instrução SQL.
  Dado = InputBox(Mensagem)
  
  'A seguir montamos a nossa instrução SQL. A sua estrutura é facil de
  'ser entendida e pode ser "traduzida" da seguinte forma:
  'SELECIONE campo Nome DA Agenda_Eletrônica ONDE Campo Nome IGUAL a
  '*Dado* ORDENADOS PELO Nome. O asterístico ao redor de DADO é utilizado
  'para indicar que as letras especificadas pelo usuário (armazenadas na
  'variável Dado) podem conter outras letras antes (*Dado) ou depois (Dado*).
  'A ordenação pelo nome é necessária para organizar os dados e a pesquisa.
  'A instrução a seguir foi colocada como um comentário pois essa pesquisa
  'seleciona apenas o campo Nome, não fornecendo resultados dos outros
  'campos da tabela Agenda_Eletronica.
  'Consulta_SQL="SELECT Nome FROM Agenda_Eletronica WHERE Nome LIKE '*" _
                    + Dado + "*' ORDER BY Nome;"
  
                    
  'No programa foi utilizada a instrução a seguir. Note que ela
  'apresenta uma diferença em relação à instrução anterior (colocada
  'como comentário): em lugar do campo Nome, utilizamos um asterístico
  '(*), o qual possibilita à seleção de todos os campos de nossa tabela.
  Consulta_SQL = "SELECT * FROM Agenda_Eletronica WHERE Nome LIKE '*" _
                  + Dado + "*' ORDER BY Nome;"
    
  
  'A instrução SQL é aplicada, através da abertura de um Recordset.
  'O resultado é armazenado na variável Pesquisa, declarada em
  'General - Declarations.
  Set Pesquisa = Agenda.OpenRecordset(Consulta_SQL, dbOpenSnapshot)
  
  
  'É montada uma mensagem com o número de registros
  'que satisfizeram a pesquisa anterior.
  Mensagem = Consulta_SQL + (Chr(13) & Chr(10)) + _
  "Apresentou total de registros = " + Str(Pesquisa.RecordCount)
  
  'Essa mensagem é apresentada ao usuário.
  Resultado = MsgBox(Mensagem, vbInformation)
          
  'A seguir, a variável Total_encontrado recebe
  'o total de registros encontrados.
  Total_encontrado = Pesquisa.RecordCount
  
  'Limpa as caixas de texto,
  'executando a função a seguir
  Limpa_Campos
     
  
  'O loop Do...Loop Until é utilizado para mostrar os
  'registros encontrados.
  Do
      'É feito um teste para verificar a posição
      'absoluta do registro que será mostrado.
      'Esse teste é apenas uma precaução para que
      'não se tente mostrar um registro inexistente.
      If Pesquisa.AbsolutePosition > -1 Then
          
          'Atualiza o conteúdo das caixas de texto
          'do nosso formulário, utilizando o Recordset
          'Pesquisa (declarado em General - Declarations).
          'Lembre-se que Atualiza_Campos_Pesquisa_SQL
          'é uma função definida em General, a qual pode
          'ser chamada a partir de outras procedures.
          Atualiza_Campos_Pesquisa_SQL
            
          'Montamos uma mensagem mostrando o nome encontrado.
          Mensagem = Str(Contador + 1) + "º Registro encontrado: " + _
          Pesquisa("nome")
          
          'Essa mensagem é apresentada ao usuário.
          Resultado = MsgBox(Mensagem, vbInformation)
          
          'A variável Contador (zerada no início desta procedure)
          'é incrementada. Ela está sendo utilizada para contar o
          'total de registros apresentados aos usuários.
          Contador = Contador + 1
      
          'Como a pesquisa SQL foi ordenada por Nome
          '(veja o trecho final da instrução SQL ORDER BY Nome;)
          'agora vamos para o próximo registro,
          'utilizando a instrução MoveNext.
          Pesquisa.MoveNext
      
      Else
          'Caso não seja encontrado um registro válido as caixas
          'de texto são limpas e uma mensagem indicativa é
          'apresentada na tela.
          Limpa_Campos
          Resultado = MsgBox("Nome não encontrado.", vbInformation)
      End If
      
  
  'O loop Do...Loop Until é executado enquanto o
  'Contador for menor ou igual ao total de resgistros
  'encontrados (utilizando a variável Total_encontrado)
  Loop Until Contador >= Total_encontrado
          
End Sub

Private Sub Nome_Completo_Btn_Click()
 'Faz a pesquisa por Nome no Banco de Dados utilizando SQL.
 'Se não localizar o registro, mostra a mensagem de "Não Existe".

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
    
 'A variável Pesquisa (definida como sendo do tipo Recordset
 '(no início desta procedure) recebe o resultado da pesquisa
 'SQL montada em Consulta_SQL .
 '
 Set Pesquisa = Agenda.OpenRecordset(Consulta_SQL, dbOpenSnapshot)

 'Caso a propriedade AbsolutePosition seja maior do que -1
 'isso indica que a pesquisa encontrou o nome especificado.
 'Para indicar o valor de AbsolutePosition, montamos a mensagem
 'a seguir e a exibimos em um MsgBox, independente do seu valor.
 Mensagem = "Posição retornada por Pesquisa.AbsolutePosition = " + _
           Str(Pesquisa.AbsolutePosition)
           
 Resultado = MsgBox(Mensagem, vbInformation)
  
 'Limpa as caixas de texto,
 'executando a função a seguir
 Limpa_Campos

 'Agora, testamos o valor da propriedade AbsolutePosition
 'a qual nos indica o sucesso ou não da pesquisa, conforme
 'comentado anteriormente.
 If Pesquisa.AbsolutePosition > -1 Then
    
    'Mostramos o nome encontrado
    Mensagem = "Nome encontrado: " + Pesquisa("nome")
    Resultado = MsgBox(Mensagem, vbInformation)
    
    Atualiza_Campos_Pesquisa_SQL
    
 Else
 
    Resultado = MsgBox("Nome não encontrado.", vbInformation)
    
 End If


End Sub

Private Sub Nome_Iniciando_Btn_Click()
 'Faz a pesquisa pela primeira parte do nome
 'no Banco de Dados utilizando uma instrução em SQL. Por exemplo,
 'caso ele deseje verificar todas as pessoas com o nome Mariana
 '(Mariana Lima, Mariana Fernandez, Mariana Silveira, etc.),
 'Basta digitar apenas Mariana, que todos os registros com esse
 'primeiro nome serão mostrados. Se não localizar o registro,
 'mostra a mensagem de "Não Existe". Os procedimentos são
 'muito similares aos apresentados para a procedure de pesquisa
 'de nomes com determinadas letras (Letras_Nome_Btn_Click).
 'A diferença está na instrução SQL (Dado + "*' ORDER BY Nome;)
 'a qual apresenta asterístico apenas de um dos lados do
 'nome especificado pelo usuário. Por exemplo, ao pesquisar
 'o nome Paulo, a instrução SQL irá procurar os nomes Paulo*.
 'Lembre-se de que o asterístico funciona como um caractere
 '(ou conjunto de caracteres) coringa, mais ou menos como
 'acontece com o MS-DOS, quando executamos o comando
 'Dir P* o qual irá mostrar todos os programas que começam
 'com a letra P (no diretório ativo).
 
 
 Dim Dado As String, Busca As String, Mensagem As String
 Dim Consulta_SQL As String, Resultado As String
 Dim Contador As Integer, Total_encontrado As Integer
 
     
 Contador = 0
 Total_encontrado = 0
                
 Dado = InputBox("Informe a Primeira Parte do Nome ")
 
 'Verifique a instrução SQL montada a seguir:
 Consulta_SQL = "SELECT * FROM Agenda_Eletronica WHERE Nome LIKE '" + Dado + "*' ORDER BY Nome;"
 'Ela pode ser "traduzida" como da seguinte forma:
 'SELECIONE todos os campos (*) DA Agenda_Eletronica ONDE Nome IGUAL
 'nome especificado (seguido de asterístico).
    
 
 Set Pesquisa = Agenda.OpenRecordset(Consulta_SQL, dbOpenSnapshot)
 
 Mensagem = Consulta_SQL + (Chr(13) & Chr(10)) + _
 "Apresentou total de registros = " + Str(Pesquisa.RecordCount)
 
 Resultado = MsgBox(Mensagem, vbInformation)
         
 Total_encontrado = Pesquisa.RecordCount
     
 
 'A seguir, os registros encontrados são
 'apresentados, um a um.
 Do
     If Pesquisa.AbsolutePosition > -1 Then
     
         Atualiza_Campos_Pesquisa_SQL
         
         Mensagem = Str(Contador + 1) + "º Registro encontrado: " + _
         Pesquisa("nome")
         
         Resultado = MsgBox(Mensagem, vbInformation)
         
         Contador = Contador + 1
     
         Pesquisa.MoveNext
     
     Else
         Resultado = MsgBox("Nome não encontrado.", vbInformation)
     End If
 
 Loop Until Contador >= Total_encontrado
 
End Sub

Private Sub Refresh_Btn_Click()
    
    'Desabilita o botão Refresh (Atualizar)
    Refresh_Btn.Enabled = False
        
    'Verifica o texto em cada caixa de texto.
    'Caseo seja igual a zero "", é acrescentado
    'um espaço em branco, simplesmente para que
    'nós possamos atribuir esses conteúdos aos
    'campos do banco de dados. No caso do
    'conteúdo de uma caixa de texto ser igual
    'a zero "" e nós tentarmos atribuí-lo a
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
    
    'No caso data de nascimento, como ela é
    'um campo específico, não basta simplesmente
    'colocarmos um espaço em branco para que
    'a atribuição seja possível. Nesse caso, temos
    'de colocar a data correta, motivo pelo qual
    'solicitamos isso ao usuário no caso do campo
    'estar vazio.
    
    If Text8.Text = "" Then
       Texto = "Você deve digitar a data " & Chr$(13) & _
                "de nascimento para prosseguir:"
       Text8.Text = InputBox(Texto, "Data de Nascimento")
    End If
    
    If Text9.Text = "" Then
       Text9.Text = " "
    End If
    
    'Faz a atribuição do conteúdo das
    'caixas de texto aos respectivos
    'campos no banco de dados.
    Tabela_Agenda("Nome") = Text1.Text
    Tabela_Agenda("Telefone") = Text2.Text
    Tabela_Agenda("Endereço") = Text3.Text
    Tabela_Agenda("Fax") = Text4.Text
    Tabela_Agenda("Cidade") = Text5.Text
    Tabela_Agenda("EstadoOuProvíncia") = Text6.Text
    Tabela_Agenda("EndereçoNoCorreioEletrônico") = Text7.Text
    Tabela_Agenda("Data de Nascimento") = Text8.Text
    Tabela_Agenda("Cep") = Text9.Text
           
    'Finalmente, atualiza o banco
    'de dados utilizando Update
    Tabela_Agenda.Update
       
    
End Sub
