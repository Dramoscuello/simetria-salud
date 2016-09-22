<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Master.Master" CodeBehind="ValidacionRips.aspx.vb" Inherits="LoginSalud.ValidacionRips" %>

<asp:Content ID="Content1" ContentPlaceHolderID="Cbody" runat="server">
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/1.12.4/jquery.min.js"></script>
    <link href="../Content/bootstrap/css/bootstrap.css" rel="stylesheet" />
    <link href="../Content/bootstrap/css/fileinput.min.css" rel="stylesheet" />
    <link href="../Content/bootstrap/css/bootstrap.min.css" rel="stylesheet" />
    <link href="../Content/bootstrap/css/fileinput.css" rel="stylesheet" />
    <link href="../Content/bootstrap/css/StyleSheet1.css" rel="stylesheet" />
    <script src="../Content/bootstrap/js/fileinput.min.js"></script>
    <script src="http://ajax.googleapis.com/ajax/libs/jquery/1.11.0/jquery.min.js"></script>
    <script src="../Scripts/sweetalert.min.js"></script>
    <link href="../Content/sweetalert.css" rel="stylesheet"/>
    
    <div class="Contenido">
        
        <asp:Label ID="Label1" runat="server" Text="Porcentaje de validacion  "></asp:Label>
        <asp:DropDownList ID="DropDownListPorcentaje" runat="server">
            <asp:ListItem Value="0">0</asp:ListItem>
            <asp:ListItem Value="10">10%</asp:ListItem>
            <asp:ListItem Value="20">20%</asp:ListItem>
            <asp:ListItem Value="30">30%</asp:ListItem>
            <asp:ListItem Value="40">40%</asp:ListItem>
            <asp:ListItem Value="50">50%</asp:ListItem>
            <asp:ListItem Value="60">60%</asp:ListItem>
            <asp:ListItem Value="70">70%</asp:ListItem>
            <asp:ListItem Value="80">80%</asp:ListItem>
            <asp:ListItem Value="90">90%</asp:ListItem>
            <asp:ListItem Value="100">100%</asp:ListItem>
        </asp:DropDownList>
        <br />
        <br />
        <div class="form-group">
            <asp:FileUpload ID="FileUploadImportar" runat="server" class="file" Multiple="Multiple" />
        </div>
         <asp:Button ID="ButtonValidar" runat="server" Text="Validar" ToolTip="Iniciar validacion" CssClass="btn btn-success" OnClientClick="Saludar()"/>
        <asp:Button ID="ButtonInforme" runat="server" Text="Descagar informe" CssClass="btn btn-success" />

        <script>
            function Saludar() {
                swal({ title: "Validando...", text: "Espere por favor...", imageUrl: "../IMG/idem-loading.gif", showConfirmButton: false });
            }
        </script>
    </div>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="Cfoot" runat="server">
</asp:Content>
