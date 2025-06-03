<%@ Page Title="Agenda Eletrônica" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Agenda.aspx.cs" Inherits="Agenda_DotNet48.Agenda" %>
<asp:Content ID="Content1" ContentPlaceHolderID="MainContent" runat="server">
    <h2>Agenda Eletrônica</h2>
    <asp:Label ID="lblMessage" runat="server" ForeColor="Red" />
    <div class="form-row">
        <span class="form-label">Nome:</span>
        <asp:TextBox ID="txtNome" runat="server" Width="300px" />
    </div>
    <div class="form-row">
        <span class="form-label">Telefone:</span>
        <asp:TextBox ID="txtTelefone" runat="server" Width="200px" />
    </div>
    <div class="form-row">
        <span class="form-label">Endereço:</span>
        <asp:TextBox ID="txtEndereco" runat="server" Width="350px" />
    </div>
    <div class="form-row">
        <span class="form-label">Fax:</span>
        <asp:TextBox ID="txtFax" runat="server" Width="200px" />
    </div>
    <div class="form-row">
        <span class="form-label">Cidade:</span>
        <asp:TextBox ID="txtCidade" runat="server" Width="200px" />
    </div>
    <div class="form-row">
        <span class="form-label">Estado:</span>
        <asp:TextBox ID="txtEstado" runat="server" Width="100px" />
    </div>
    <div class="form-row">
        <span class="form-label">E-mail:</span>
        <asp:TextBox ID="txtEmail" runat="server" Width="250px" />
    </div>
    <div class="form-row">
        <span class="form-label">Data Nasc.:</span>
        <asp:TextBox ID="txtDataNasc" runat="server" Width="120px" />
    </div>
    <div class="form-row">
        <span class="form-label">CEP:</span>
        <asp:TextBox ID="txtCEP" runat="server" Width="100px" />
    </div>
    <div class="actions">
        <asp:Button ID="btnPrimeiro" runat="server" Text="Primeiro" OnClick="btnPrimeiro_Click" />
        <asp:Button ID="btnAnterior" runat="server" Text="Anterior" OnClick="btnAnterior_Click" />
        <asp:Button ID="btnProximo" runat="server" Text="Próximo" OnClick="btnProximo_Click" />
        <asp:Button ID="btnUltimo" runat="server" Text="Último" OnClick="btnUltimo_Click" />
        <asp:Button ID="btnIncluir" runat="server" Text="Inclusão" OnClick="btnIncluir_Click" />
        <asp:Button ID="btnExcluir" runat="server" Text="Exclusão" OnClick="btnExcluir_Click" />
        <asp:Button ID="btnAtualizar" runat="server" Text="Atualizar" OnClick="btnAtualizar_Click" />
        <asp:Button ID="btnPesquisa" runat="server" Text="Pesquisa" OnClick="btnPesquisa_Click" />
        <asp:Button ID="btnNomeCompleto" runat="server" Text="Nome Completo" OnClick="btnNomeCompleto_Click" />
        <asp:Button ID="btnNomeIniciando" runat="server" Text="Nomes Iniciando Com" OnClick="btnNomeIniciando_Click" />
        <asp:Button ID="btnLetrasNome" runat="server" Text="Letras do Nome" OnClick="btnLetrasNome_Click" />
        <asp:Button ID="btnImprimir" runat="server" Text="Imprimir" OnClick="btnImprimir_Click" />
        <asp:Button ID="btnFim" runat="server" Text="Fim" OnClick="btnFim_Click" />
    </div>
</asp:Content> 