<%--<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Default.aspx.vb" Inherits="LoginSalud._Default" %>--%>

<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Default.aspx.vb" Inherits="LoginSalud._Default" MasterPageFile="~/Site1.Master"%>


<asp:Content ID="cb" ContentPlaceHolderID="Cbody" runat="server">
      
    <link rel="shortcut icon" type="image/x-icon" href="img/icon.ico" />
    <link href="Content/bootstrap/css/form.css" rel="stylesheet" />
   <div class="card card-container">
  
        <h4 class="panel-title cta-text cta-panel" style="box-sizing: border-box; font-family:Neuropolitical;font-size: 18.8px;  font-weight: normal; line-height: normal; color: rgb(0, 96, 172); margin-top: 10px;">
          
            <%--<h4 class="panel-title cta-text cta-panel" style="box-sizing: border-box; font-family:Neuropolitical; font-weight: normal; line-height: normal; color: rgb(0, 96, 172); margin-top: 10px; margin-bottom: 0px; font-size: 18.8px; text-align: center; text-decoration: none; font-style: normal; font-variant: normal; font-stretch: normal; text-shadow: rgb(247, 250, 247) 0px 1px 0px, rgb(201, 201, 201) 0px 2px 0px, rgb(247, 250, 247) 0px 3px 0px, rgb(247, 250, 247) 0px 4px 0px, rgb(247, 250, 247) 0px 5px 0px, rgba(0, 0, 0, 0.0980392) 0px 6px 1px, rgba(0, 0, 0, 0.0980392) 0px 0px 5px, rgba(0, 0, 0, 0.298039) 0px 1px 3px, rgba(0, 0, 0, 0.2) 0px 3px 5px, rgba(0, 0, 0, 0.247059) 0px 5px 10px, rgba(0, 0, 0, 0.2) 0px 10px 10px, rgba(0, 0, 0, 0.14902) 0px 20px 20px; letter-spacing: normal; orphans: 2; text-indent: 0px; text-transform: none; white-space: normal; widows: 1; word-spacing: 0px; -webkit-text-stroke-width: 0px;">--%>
           <strong class="">Acceso de Usuarios</strong>
          
            <hr />

           &nbsp;</h4>
         <img id="profile-img" class="profile-img-card"   src="IMG/Logo.png" />
        <p id="profile-name" class="profile-name-card"></p>
        <div class="form-signin">
            <asp:Label ID="Label1" runat="server" Text="Usuario:" CssClass="control-label" Font-Bold="True"></asp:Label>
            <asp:TextBox ID="inputEmail" runat="server" CssClass="form-control"></asp:TextBox>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ErrorMessage="¡Campo requerido!" ControlToValidate="inputEmail" CssClass="text-danger"></asp:RequiredFieldValidator>
            <br/>
            <asp:Label ID="Label2" runat="server" Text="Contraseña:" CssClass="control-label" Font-Bold="True"></asp:Label>
            <asp:TextBox ID="inputPassword" runat="server" CssClass="form-control"   ></asp:TextBox>
            <asp:RequiredFieldValidator  ID="RequiredFieldValidator2" runat="server" ErrorMessage="¡Campo requerido!" ControlToValidate="inputPassword" CssClass="text-danger"></asp:RequiredFieldValidator>
            <br />
            <!--<div id="remember" class="checkbox">
                <asp:CheckBox ID="CheckBox1" runat="server" /><asp:Label ID="Labeltt" runat="server" Text="Label"></asp:Label>
            </div>-->
    
            <asp:Button ID="login" runat="server" Text="Ingresar" CssClass="btn btn-lg btn-primary btn-block btn-signin"/>
        </div>
        <div>
            <asp:Label ID="Label3" runat="server">
                <div class="alert alert-dismissible alert-danger mtop">
                    <button type="button" class="close" data-dismiss="alert">&times;</button>
                    <strong>Usuario y/o contraseña incorrecto</strong>
                </div>
            </asp:Label>
        </div>
    </div>

</asp:Content>