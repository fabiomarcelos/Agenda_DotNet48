using System.Collections.Generic;

namespace Agenda_DotNet48
{
    public interface IAgendaService
    {
        List<AgendaModel> ObterTodos();
        void Incluir(AgendaModel model);
        void Atualizar(string nomeAntigo, AgendaModel model);
        void Excluir(string nome);
        List<AgendaModel> PesquisarPorNome(string nome);
        List<AgendaModel> PesquisarPorInicioNome(string nomeInicio);
    }
} 