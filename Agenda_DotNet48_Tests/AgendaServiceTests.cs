using System;
using System.Collections.Generic;
using Agenda_DotNet48;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;

namespace Agenda_DotNet48_Tests
{
    [TestClass]
    public class AgendaServiceTests
    {
        // Teste de sucesso para inclusão
        [TestMethod]
        public void Incluir_DadosValidos_DeveChamarRepositorio()
        {
            // Arrange
            var mockRepo = new Mock<IAgendaRepository>();
            var service = new AgendaService(mockRepo.Object);
            var model = new AgendaModel { Nome = "Teste" };

            // Act
            service.Incluir(model);

            // Assert
            mockRepo.Verify(r => r.Add(model), Times.Once);
            // Este teste garante que, ao incluir dados válidos, o repositório é chamado corretamente.
        }

        // Teste de falha para inclusão (nome vazio)
        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void Incluir_NomeVazio_DeveLancarExcecao()
        {
            // Arrange
            var mockRepo = new Mock<IAgendaRepository>();
            var service = new AgendaService(mockRepo.Object);
            var model = new AgendaModel { Nome = "" };

            // Act
            service.Incluir(model);
            // Este teste garante que, ao tentar incluir com nome vazio, uma exceção é lançada.
        }

        // Teste de exceção do repositório
        [TestMethod]
        [ExpectedException(typeof(Exception))]
        public void Incluir_RepositorioFalha_DevePropagarExcecao()
        {
            // Arrange
            var mockRepo = new Mock<IAgendaRepository>();
            mockRepo.Setup(r => r.Add(It.IsAny<AgendaModel>())).Throws(new Exception("Erro DB"));
            var service = new AgendaService(mockRepo.Object);
            var model = new AgendaModel { Nome = "Teste" };

            // Act
            service.Incluir(model);
            // Este teste garante que, se o repositório lançar exceção, ela é propagada pelo serviço.
        }

        // Teste de pesquisa por nome com sucesso
        [TestMethod]
        public void PesquisarPorNome_Existente_DeveRetornarLista()
        {
            // Arrange
            var mockRepo = new Mock<IAgendaRepository>();
            var esperado = new List<AgendaModel> { new AgendaModel { Nome = "Teste" } };
            mockRepo.Setup(r => r.SearchByName("Teste")).Returns(esperado);
            var service = new AgendaService(mockRepo.Object);

            // Act
            var resultado = service.PesquisarPorNome("Teste");

            // Assert
            Assert.AreEqual(1, resultado.Count);
            // Este teste garante que a pesquisa por nome retorna a lista esperada.
        }

        // Teste de pesquisa por nome com nome vazio
        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void PesquisarPorNome_NomeVazio_DeveLancarExcecao()
        {
            // Arrange
            var mockRepo = new Mock<IAgendaRepository>();
            var service = new AgendaService(mockRepo.Object);

            // Act
            service.PesquisarPorNome("");
            // Este teste garante que pesquisar por nome vazio lança exceção.
        }
    }
} 