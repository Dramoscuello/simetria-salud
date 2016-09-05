Public Class Formulario_web1
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Response.Cache.SetCacheability(HttpCacheability.ServerAndNoCache)
        Response.Cache.SetAllowResponseInBrowserHistory(False)
        Response.Cache.SetNoStore()
        If Session("usuario") IsNot Nothing Then
            'Label1.Text = Session("usuario")
            'Label2.Text = Session("tusuario")
        Else
            Response.Redirect("~/Ingreso")
        End If

    End Sub

End Class