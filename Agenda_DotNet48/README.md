# Visão Geral do Projeto

O sistema **Agenda_DotNet48** é uma aplicação ASP.NET WebForms (.NET Framework 4.8) para gerenciamento de uma agenda eletrônica, permitindo cadastro, pesquisa, atualização e exclusão de contatos. O projeto é uma evolução de um sistema legado em VB6, modernizando a interface e arquitetura, mas mantendo a simplicidade e foco em operações CRUD sobre um banco de dados Access.

---

## Estrutura de Pastas e Projetos

```
/Agenda_DotNet48           # Projeto principal Web (ASP.NET WebForms)
  /App_Data                # Base de dados Access (.accdb)
  /App_Start               # Configurações de rotas e bundles
  /Content                 # CSS e arquivos estáticos
  /Scripts                 # JS (jQuery, Bootstrap, Modernizr, WebForms)
  /bin                     # Binários compilados
  /Properties              # Configurações do projeto
  /obj, /packages, etc.    # Pastas auxiliares do .NET
  Agenda.aspx              # Página principal da agenda
  AgendaRepository.cs      # Repositório de dados (OleDb)
  AgendaService.cs         # Lógica de negócio
  AgendaModel.cs           # Modelo de dados
  Global.asax              # Inicialização global
  Web.config               # Configuração da aplicação

/Agenda_DotNet48_Tests     # Projeto de testes unitários (MSTest)
  AgendaServiceTests.cs    # Testes de unidade do serviço
  MSTestSettings.cs        # Configuração de paralelismo dos testes

/AgendaVB6                 # Código legado em VB6 (referência histórica)
```

**Responsabilidades:**
- `Agenda_DotNet48`: WebApp principal, lógica de negócio, acesso a dados e interface.
- `Agenda_DotNet48_Tests`: Testes unitários automatizados.
- `AgendaVB6`: Código fonte original em VB6 (não utilizado em produção, apenas referência).

---

## Responsabilidades das Classes e Métodos

### Projeto: `Agenda_DotNet48`

#### **AgendaModel**
- **Responsabilidade:** Representa a entidade de contato da agenda, contendo propriedades como Nome, Telefone, Endereço, etc.
- **Propriedades:**  
  - `Nome`, `Telefone`, `Endereco`, `Fax`, `Cidade`, `Estado`, `Email`, `DataNasc`, `CEP`: armazenam os dados do contato.

#### **IAgendaRepository**
- **Responsabilidade:** Define o contrato para operações de acesso a dados da agenda.
- **Métodos:**
  - `List<AgendaModel> GetAll()`: Retorna todos os contatos.
  - `void Add(AgendaModel model)`: Adiciona um novo contato.
  - `void Update(string nomeAntigo, AgendaModel model)`: Atualiza um contato existente.
  - `void Delete(string nome)`: Remove um contato pelo nome.
  - `List<AgendaModel> SearchByName(string nome)`: Busca contatos por nome exato.
  - `List<AgendaModel> SearchByNameStart(string nomeInicio)`: Busca contatos por início do nome.

#### **AgendaRepository**
- **Responsabilidade:** Implementa o acesso a dados usando OleDb e banco Access.
- **Métodos:**  
  - Implementa todos os métodos definidos em `IAgendaRepository`, realizando operações SQL no banco de dados.
  - `private AgendaModel DataRowToModel(DataRow row)`: Converte uma linha do DataTable em um objeto `AgendaModel`.

#### **IAgendaService**
- **Responsabilidade:** Define o contrato para a lógica de negócio da agenda.
- **Métodos:**
  - `List<AgendaModel> ObterTodos()`: Retorna todos os contatos.
  - `void Incluir(AgendaModel model)`: Valida e adiciona um novo contato.
  - `void Atualizar(string nomeAntigo, AgendaModel model)`: Valida e atualiza um contato.
  - `void Excluir(string nome)`: Valida e remove um contato.
  - `List<AgendaModel> PesquisarPorNome(string nome)`: Valida e busca contatos por nome exato.
  - `List<AgendaModel> PesquisarPorInicioNome(string nomeInicio)`: Valida e busca contatos por início do nome.

#### **AgendaService**
- **Responsabilidade:** Implementa a lógica de negócio, validações e delega operações ao repositório.
- **Métodos:**
  - `ObterTodos()`: Chama o repositório para obter todos os contatos.
  - `Incluir(AgendaModel model)`: Valida o modelo (ex: nome obrigatório) e chama o repositório para adicionar.
  - `Atualizar(string nomeAntigo, AgendaModel model)`: Valida e atualiza um contato existente.
  - `Excluir(string nome)`: Valida e remove um contato.
  - `PesquisarPorNome(string nome)`: Valida e busca contatos por nome exato.
  - `PesquisarPorInicioNome(string nomeInicio)`: Valida e busca contatos por início do nome.

#### **Agenda.aspx.cs** (Code-behind da página principal)
- **Responsabilidade:** Controla a interação do usuário com a interface WebForms, manipula eventos dos botões e exibe dados.
- **Principais métodos/eventos:**
  - `OnInit`: Inicializa o serviço de agenda.
  - `Page_Load`: Carrega os dados na primeira carga da página.
  - `LoadAgenda()`: Carrega todos os contatos para a ViewState.
  - `ShowRecord(int index)`: Exibe um contato na interface.
  - `LimpaCampos()`: Limpa os campos do formulário.
  - Eventos de botões (`btnPrimeiro_Click`, `btnAnterior_Click`, etc.): Navegação, inclusão, atualização, exclusão e pesquisa de contatos, sempre interagindo com o serviço.

#### **Global.asax.cs**
- **Responsabilidade:** Inicializa configurações globais da aplicação, como rotas e bundles, no evento `Application_Start`.

#### **RouteConfig**
- **Responsabilidade:** Configura URLs amigáveis para a aplicação.

#### **BundleConfig**
- **Responsabilidade:** Configura o agrupamento e minificação de scripts e estilos.

---

### Projeto: `Agenda_DotNet48_Tests`

#### **AgendaServiceTests**
- **Responsabilidade:** Testa a lógica de negócio do serviço de agenda, usando mocks para o repositório.
- **Métodos de teste:**
  - `Incluir_DadosValidos_DeveChamarRepositorio`: Garante que a inclusão chama o repositório.
  - `Incluir_NomeVazio_DeveLancarExcecao`: Garante que não é possível incluir contato sem nome.
  - `Incluir_RepositorioFalha_DevePropagarExcecao`: Garante que exceções do repositório são propagadas.
  - `PesquisarPorNome_Existente_DeveRetornarLista`: Garante que a busca retorna resultados esperados.
  - `PesquisarPorNome_NomeVazio_DeveLancarExcecao`: Garante que a busca com nome vazio lança exceção.

#### **MSTestSettings**
- **Responsabilidade:** Configura a execução paralela dos testes.

---

## Tecnologias Utilizadas

- **ASP.NET WebForms** (.NET Framework 4.8)
- **ADO.NET OleDb** (acesso a banco Access)
- **Bootstrap 5.2.3** (UI responsiva)
- **jQuery 3.7.0** (interatividade)
- **Modernizr 2.8.3** (detecção de recursos)
- **Newtonsoft.Json 13.0.3** (serialização JSON)
- **MSTest** (testes unitários)
- **Moq** (mock de dependências nos testes)
- **Microsoft.AspNet.FriendlyUrls** (URLs amigáveis)
- **WebGrease, Antlr** (otimização e parsing)
- **Microsoft.CodeDom.Providers.DotNetCompilerPlatform** (compilação dinâmica)
- **Microsoft.NET.Test.Sdk** (infraestrutura de testes)

---

## Configuração do Ambiente

### Requisitos

- **Windows** com IIS ou IIS Express
- **Visual Studio 2019/2022** (com suporte a .NET Framework 4.8)
- **Microsoft Access Database Engine** (para conexão OleDb com .accdb)
- **.NET Framework 4.8 Developer Pack**

### Passos para rodar localmente

1. **Clone o repositório:**
   ```sh
   git clone <url-do-repositorio>
   ```

2. **Abra a solução no Visual Studio:**
   - Arquivo: `Agenda_DotNet48/Agenda_DotNet48.sln`

3. **Restaure os pacotes NuGet:**
   - Visual Studio: Clique com o direito na solução > "Restaurar Pacotes NuGet"
   - Ou via terminal:
     ```sh
     nuget restore Agenda_DotNet48.sln
     ```

4. **Configure o banco de dados:**
   - O arquivo `App_Data/Agenda1.accdb` já está incluído.
   - Certifique-se de que o usuário do IIS/IIS Express tenha permissão de leitura/escrita na pasta `App_Data`.

5. **Configuração IIS/IIS Express:**
   - O projeto está pronto para rodar com IIS Express (default do Visual Studio).
   - Para IIS local, crie um site apontando para a pasta `Agenda_DotNet48` e configure o Application Pool para .NET 4.8.

6. **Executando:**
   - Pressione `F5` no Visual Studio ou use o menu "Iniciar Depuração".

---

## Executando os Testes

### Visual Studio

- Abra o Test Explorer (`Ctrl+E, T`).
- Clique em "Run All" para executar todos os testes.

### Linha de comando

```sh
cd Agenda_DotNet48_Tests
dotnet test
```

- Os testes utilizam MSTest e Moq para simular dependências.
- Exemplo de teste (AgendaServiceTests.cs):

```csharp
[TestMethod]
public void Incluir_DadosValidos_DeveChamarRepositorio()
{
    var mockRepo = new Mock<IAgendaRepository>();
    var service = new AgendaService(mockRepo.Object);
    var model = new AgendaModel { Nome = "Teste" };

    service.Incluir(model);

    mockRepo.Verify(r => r.Add(model), Times.Once);
}
```

---

## Boas Práticas e Convenções

- **Naming conventions:** 
  - Classes: `PascalCase` (ex: `AgendaService`)
  - Métodos: `PascalCase` (ex: `Incluir`)
  - Variáveis/Parâmetros: `camelCase`
  - Interfaces: prefixo `I` (ex: `IAgendaRepository`)
- **Organização:**
  - Separação clara entre camada de apresentação (WebForms), serviço (lógica de negócio) e repositório (acesso a dados).
  - Modelos de dados simples (POCO).
  - Testes unitários isolam dependências com Moq.
- **Testes:**
  - Estruture os testes por classe de serviço.
  - Use `[ExpectedException]` para validar erros esperados.
- **Dicas de manutenção:**
  - Centralize a string de conexão em um único local.
  - Evite lógica de negócio no code-behind das páginas.
  - Prefira injeção de dependência (mesmo que manual) para facilitar testes.

---

## Possíveis Melhorias

- **Separação de camadas:** Adotar um padrão mais claro de camadas (ex: Application, Domain, Infrastructure).
- **Injeção de Dependência:** Utilizar um container de DI (ex: Unity, Autofac).
- **Validação:** Implementar validação de dados com DataAnnotations.
- **Acesso a dados:** Migrar para Entity Framework ou Dapper para maior produtividade e manutenção.
- **Migração para .NET Core/ASP.NET Core:** Modernizar o projeto para multiplataforma e melhor performance.
- **Testes de integração:** Adicionar testes que validem o fluxo completo com banco de dados real (usando banco em memória ou mocks avançados).
- **CI/CD:** Automatizar build e testes com GitHub Actions ou Azure DevOps.

---

## Histórico de Versões

| Versão | Data       | Descrição                      |
|--------|------------|-------------------------------|
| 1.0    | 2024-06-XX | Primeira versão migrada do VB6 |
| 1.1    | 2024-06-XX | Adição de testes unitários     |

--- 