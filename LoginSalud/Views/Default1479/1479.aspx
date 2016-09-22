<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Master.Master" CodeBehind="1479.aspx.vb" Inherits="LoginSalud._1479" %>
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
        
    
        <div class="form-group">
            <asp:FileUpload ID="FileUploadImportar" runat="server" class="file" Multiple="Multiple" />
        </div>
         <asp:Button ID="ButtonValidar" runat="server" Text="Validar" ToolTip="Iniciar validacion" CssClass="btn btn-success" OnClientClick="Saludar()"/>
        <asp:Button ID="ButtonInforme" runat="server" Text="Descagar informe" CssClass="btn btn-success" />

        <script>
            function Saludar() {
                swal({ title: "Validando 1479 ...", text: "Espere por favor...", imageUrl: "../IMG/idem-loading.gif", showConfirmButton: false });
            }
        </script>
    </div>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="Cfoot" runat="server">
</asp:Content>