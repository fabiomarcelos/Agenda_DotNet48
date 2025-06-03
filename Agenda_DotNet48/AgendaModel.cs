using System;

namespace Agenda_DotNet48
{
    [Serializable]
    public class AgendaModel
    {
        public string Nome { get; set; }
        public string Telefone { get; set; }
        public string Endereco { get; set; }
        public string Fax { get; set; }
        public string Cidade { get; set; }
        public string Estado { get; set; }
        public string Email { get; set; }
        public string DataNasc { get; set; }
        public string CEP { get; set; }
    }
} 