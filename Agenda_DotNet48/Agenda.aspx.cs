using System;
using System.Data;
using System.Data.OleDb;
using System.Web.UI;
using System.Linq;
using System.Collections.Generic;

namespace Agenda_DotNet48
{
    public partial class Agenda : Page
    {
        private IAgendaService _service;
        private string ConnectionString => $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Server.MapPath("~/App_Data/Agenda1.accdb");
        private List<AgendaModel> AgendaList
        {
            get { return (List<AgendaModel>)ViewState["AgendaList"]; }
            set { ViewState["AgendaList"] = value; }
        }
        private int CurrentIndex
        {
            get { return ViewState["CurrentIndex"] != null ? (int)ViewState["CurrentIndex"] : 0; }
            set { ViewState["CurrentIndex"] = value; }
        }

        protected override void OnInit(EventArgs e)
        {
            base.OnInit(e);
            _service = new AgendaService(new AgendaRepository(ConnectionString));
        }

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                LoadAgenda();
                ShowRecord(0);
            }
        }

        private void LoadAgenda()
        {
            AgendaList = _service.ObterTodos();
        }

        private void ShowRecord(int index)
        {
            if (AgendaList == null || AgendaList.Count == 0)
            {
                LimpaCampos();
                lblMessage.Text = "Nenhum registro encontrado.";
                return;
            }
            if (index < 0) index = 0;
            if (index >= AgendaList.Count) index = AgendaList.Count - 1;
            var model = AgendaList[index];
            txtNome.Text = model.Nome;
            txtTelefone.Text = model.Telefone;
            txtEndereco.Text = model.Endereco;
            txtFax.Text = model.Fax;
            txtCidade.Text = model.Cidade;
            txtEstado.Text = model.Estado;
            txtEmail.Text = model.Email;
            txtDataNasc.Text = model.DataNasc;
            txtCEP.Text = model.CEP;
            CurrentIndex = index;
            lblMessage.Text = string.Empty;
        }

        private void LimpaCampos()
        {
            txtNome.Text = "";
            txtTelefone.Text = "";
            txtEndereco.Text = "";
            txtFax.Text = "";
            txtCidade.Text = "";
            txtEstado.Text = "";
            txtEmail.Text = "";
            txtDataNasc.Text = "";
            txtCEP.Text = "";
        }

        protected void btnPrimeiro_Click(object sender, EventArgs e) => ShowRecord(0);
        protected void btnAnterior_Click(object sender, EventArgs e) => ShowRecord(CurrentIndex - 1);
        protected void btnProximo_Click(object sender, EventArgs e) => ShowRecord(CurrentIndex + 1);
        protected void btnUltimo_Click(object sender, EventArgs e) => ShowRecord(AgendaList.Count - 1);

        protected void btnIncluir_Click(object sender, EventArgs e)
        {
            LimpaCampos();
            lblMessage.Text = "Preencha os campos e clique em Atualizar para salvar.";
            ViewState["Inclusao"] = true;
        }

        protected void btnAtualizar_Click(object sender, EventArgs e)
        {
            bool inclusao = ViewState["Inclusao"] != null && (bool)ViewState["Inclusao"];
            var model = new AgendaModel
            {
                Nome = txtNome.Text.Trim(),
                Telefone = txtTelefone.Text.Trim(),
                Endereco = txtEndereco.Text.Trim(),
                Fax = txtFax.Text.Trim(),
                Cidade = txtCidade.Text.Trim(),
                Estado = txtEstado.Text.Trim(),
                Email = txtEmail.Text.Trim(),
                DataNasc = txtDataNasc.Text.Trim(),
                CEP = txtCEP.Text.Trim()
            };
            try
            {
                if (inclusao)
                {
                    _service.Incluir(model);
                    lblMessage.Text = "Registro incluído com sucesso.";
                }
                else
                {
                    var nomeAntigo = AgendaList[CurrentIndex].Nome;
                    _service.Atualizar(nomeAntigo, model);
                    lblMessage.Text = "Registro atualizado com sucesso.";
                }
                ViewState["Inclusao"] = false;
            }
            catch (Exception ex)
            {
                lblMessage.Text = ex.Message;
            }
            LoadAgenda();
            ShowRecord(CurrentIndex);
        }

        protected void btnExcluir_Click(object sender, EventArgs e)
        {
            if (AgendaList == null || AgendaList.Count == 0) return;
            try
            {
                var nome = AgendaList[CurrentIndex].Nome;
                _service.Excluir(nome);
                lblMessage.Text = "Registro excluído.";
            }
            catch (Exception ex)
            {
                lblMessage.Text = ex.Message;
            }
            LoadAgenda();
            ShowRecord(0);
        }

        protected void btnPesquisa_Click(object sender, EventArgs e)
        {
            string nome = txtNome.Text.Trim();
            try
            {
                var result = _service.PesquisarPorNome(nome);
                if (result.Count > 0)
                {
                    AgendaList = result;
                    ShowRecord(0);
                    lblMessage.Text = $"{result.Count} registro(s) encontrado(s).";
                }
                else
                {
                    lblMessage.Text = "Nome não encontrado.";
                }
            }
            catch (Exception ex)
            {
                lblMessage.Text = ex.Message;
            }
        }

        protected void btnNomeCompleto_Click(object sender, EventArgs e) => btnPesquisa_Click(sender, e);

        protected void btnNomeIniciando_Click(object sender, EventArgs e)
        {
            string nome = txtNome.Text.Trim();
            try
            {
                var result = _service.PesquisarPorInicioNome(nome);
                if (result.Count > 0)
                {
                    AgendaList = result;
                    ShowRecord(0);
                    lblMessage.Text = $"{result.Count} registro(s) encontrado(s).";
                }
                else
                {
                    lblMessage.Text = "Nenhum registro encontrado.";
                }
            }
            catch (Exception ex)
            {
                lblMessage.Text = ex.Message;
            }
        }

        protected void btnLetrasNome_Click(object sender, EventArgs e)
        {
            string letras = txtNome.Text.Trim();
            if (string.IsNullOrEmpty(letras))
            {
                lblMessage.Text = "Informe parte do nome para pesquisar.";
                return;
            }
            using (var conn = new OleDbConnection(ConnectionString))
            {
                var da = new OleDbDataAdapter($"SELECT * FROM Agenda_Eletronica WHERE Nome LIKE ? ORDER BY Nome", conn);
                da.SelectCommand.Parameters.AddWithValue("@Nome", "%" + letras + "%");
                var dt = new DataTable();
                da.Fill(dt);
                if (dt.Rows.Count > 0)
                {
                    AgendaList = dt.AsEnumerable().Select(row => new AgendaModel
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
                    }).ToList();
                    ShowRecord(0);
                    lblMessage.Text = $"{AgendaList.Count} registro(s) encontrado(s).";
                }
                else
                {
                    lblMessage.Text = "Nenhum registro encontrado.";
                }
            }
        }

        protected void btnImprimir_Click(object sender, EventArgs e)
        {
            // Em WebForms, a impressão é feita pelo navegador. Pode-se gerar uma página de impressão.
            lblMessage.Text = "Use o comando de impressão do navegador (Ctrl+P).";
        }

        protected void btnFim_Click(object sender, EventArgs e)
        {
            Response.Redirect("Default.aspx");
        }
    }
}