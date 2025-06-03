using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Web;

namespace Agenda_DotNet48
{
    public class AgendaRepository : IAgendaRepository
    {
        private readonly string _connectionString;
        public AgendaRepository(string connectionString)
        {
            _connectionString = connectionString;
        }

        public List<AgendaModel> GetAll()
        {
            var list = new List<AgendaModel>();
            using (var conn = new OleDbConnection(_connectionString))
            {
                var da = new OleDbDataAdapter("SELECT * FROM Agenda_Eletronica ORDER BY Nome", conn);
                var dt = new DataTable();
                da.Fill(dt);
                foreach (DataRow row in dt.Rows)
                {
                    list.Add(DataRowToModel(row));
                }
            }
            return list;
        }

        public void Add(AgendaModel model)
        {
            using (var conn = new OleDbConnection(_connectionString))
            {
                conn.Open();
                var cmd = new OleDbCommand("INSERT INTO Agenda_Eletronica ([Nome], [Telefone], [Endereco], [Fax], [Cidade], [EstadoOuProvincia], [EnderecoNoCorreioEletronico], [Data de Nascimento], [CEP]) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)", conn);
                cmd.Parameters.AddWithValue("@Nome", model.Nome);
                cmd.Parameters.AddWithValue("@Telefone", model.Telefone);
                cmd.Parameters.AddWithValue("@Endereco", model.Endereco);
                cmd.Parameters.AddWithValue("@Fax", model.Fax);
                cmd.Parameters.AddWithValue("@Cidade", model.Cidade);
                cmd.Parameters.AddWithValue("@Estado", model.Estado);
                cmd.Parameters.AddWithValue("@Email", model.Email);
                cmd.Parameters.AddWithValue("@DataNasc", model.DataNasc);
                cmd.Parameters.AddWithValue("@CEP", model.CEP);
                cmd.ExecuteNonQuery();
            }
        }

        public void Update(string nomeAntigo, AgendaModel model)
        {
            using (var conn = new OleDbConnection(_connectionString))
            {
                conn.Open();
                var cmd = new OleDbCommand("UPDATE Agenda_Eletronica SET [Nome]=?, [Telefone]=?, [Endereco]=?, [Fax]=?, [Cidade]=?, [EstadoOuProvincia]=?, [EnderecoNoCorreioEletronico]=?, [Data de Nascimento]=?, [CEP]=? WHERE [Nome]=?", conn);
                cmd.Parameters.AddWithValue("@Nome", model.Nome);
                cmd.Parameters.AddWithValue("@Telefone", model.Telefone);
                cmd.Parameters.AddWithValue("@Endereco", model.Endereco);
                cmd.Parameters.AddWithValue("@Fax", model.Fax);
                cmd.Parameters.AddWithValue("@Cidade", model.Cidade);
                cmd.Parameters.AddWithValue("@Estado", model.Estado);
                cmd.Parameters.AddWithValue("@Email", model.Email);
                cmd.Parameters.AddWithValue("@DataNasc", model.DataNasc);
                cmd.Parameters.AddWithValue("@CEP", model.CEP);
                cmd.Parameters.AddWithValue("@NomeAntigo", nomeAntigo);
                cmd.ExecuteNonQuery();
            }
        }

        public void Delete(string nome)
        {
            using (var conn = new OleDbConnection(_connectionString))
            {
                conn.Open();
                var cmd = new OleDbCommand("DELETE FROM Agenda_Eletronica WHERE [Nome]=?", conn);
                cmd.Parameters.AddWithValue("@Nome", nome);
                cmd.ExecuteNonQuery();
            }
        }

        public List<AgendaModel> SearchByName(string nome)
        {
            var list = new List<AgendaModel>();
            using (var conn = new OleDbConnection(_connectionString))
            {
                var da = new OleDbDataAdapter("SELECT * FROM Agenda_Eletronica WHERE Nome = ?", conn);
                da.SelectCommand.Parameters.AddWithValue("@Nome", nome);
                var dt = new DataTable();
                da.Fill(dt);
                foreach (DataRow row in dt.Rows)
                {
                    list.Add(DataRowToModel(row));
                }
            }
            return list;
        }

        public List<AgendaModel> SearchByNameStart(string nomeInicio)
        {
            var list = new List<AgendaModel>();
            using (var conn = new OleDbConnection(_connectionString))
            {
                var da = new OleDbDataAdapter("SELECT * FROM Agenda_Eletronica WHERE Nome LIKE ? ORDER BY Nome", conn);
                da.SelectCommand.Parameters.AddWithValue("@Nome", nomeInicio + "%");
                var dt = new DataTable();
                da.Fill(dt);
                foreach (DataRow row in dt.Rows)
                {
                    list.Add(DataRowToModel(row));
                }
            }
            return list;
        }

        private AgendaModel DataRowToModel(DataRow row)
        {
            return new AgendaModel
            {
                Nome = row["Nome"].ToString(),
                Telefone = row["Telefone"].ToString(),
                Endereco = row["Endereco"].ToString(),
                Fax = row["Fax"].ToString(),
                Cidade = row["Cidade"].ToString(),
                Estado = row["EstadoOuProvincia"].ToString(),
                Email = row["EnderecoNoCorreioEletronico"].ToString(),
                DataNasc = row["Data de Nascimento"].ToString(),
                CEP = row["CEP"].ToString()
            };
        }
    }
} 