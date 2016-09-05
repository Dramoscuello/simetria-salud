<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="login.aspx.vb" Inherits="LoginSalud.login" MasterPageFile="~/Master.Master"%>

<asp:Content ID="c1" ContentPlaceHolderID="Chead" runat="server"></asp:Content>
<asp:Content ID="cb" ContentPlaceHolderID="Cbody" runat="server">
    <div class="row">
        <div class="col-xs-12">
            <h1>Bienvenido <asp:Label ID="Label1" runat="server" Text="Label"></asp:Label></h1>
            <asp:Button ID="Button1" runat="server" Text="Cerrar sesión" CssClass="btn btn-danger"/>
        </div>
    </div>
</asp:Content>