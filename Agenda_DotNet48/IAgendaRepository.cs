using System.Collections.Generic;

namespace Agenda_DotNet48
{
    public interface IAgendaRepository
    {
        List<AgendaModel> GetAll();
        void Add(AgendaModel model);
        void Update(string nomeAntigo, AgendaModel model);
        void Delete(string nome);
        List<AgendaModel> SearchByName(string nome);
        List<AgendaModel> SearchByNameStart(string nomeInicio);
    }
} 