using System;
using System.Collections.Generic;

namespace Agenda_DotNet48
{
    public class AgendaService : IAgendaService
    {
        private readonly IAgendaRepository _repo;
        public AgendaService(IAgendaRepository repo)
        {
            _repo = repo;
        }

        public List<AgendaModel> ObterTodos()
        {
            return _repo.GetAll();
        }

        public void Incluir(AgendaModel model)
        {
            if (string.IsNullOrWhiteSpace(model.Nome))
                throw new ArgumentException("Nome é obrigatório.");
            _repo.Add(model);
        }

        public void Atualizar(string nomeAntigo, AgendaModel model)
        {
            if (string.IsNullOrWhiteSpace(model.Nome))
                throw new ArgumentException("Nome é obrigatório.");
            _repo.Update(nomeAntigo, model);
        }

        public void Excluir(string nome)
        {
            if (string.IsNullOrWhiteSpace(nome))
                throw new ArgumentException("Nome é obrigatório para exclusão.");
            _repo.Delete(nome);
        }

        public List<AgendaModel> PesquisarPorNome(string nome)
        {
            if (string.IsNullOrWhiteSpace(nome))
                throw new ArgumentException("Nome é obrigatório para pesquisa.");
            return _repo.SearchByName(nome);
        }

        public List<AgendaModel> PesquisarPorInicioNome(string nomeInicio)
        {
            if (string.IsNullOrWhiteSpace(nomeInicio))
                throw new ArgumentException("Início do nome é obrigatório para pesquisa.");
            return _repo.SearchByNameStart(nomeInicio);
        }
    }
} 