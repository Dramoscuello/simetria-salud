Public Class Master
    Inherits System.Web.UI.MasterPage



    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Label1.Text = Session("usuario")
        If Session("tusuario") = 2 Then
            hide1.Visible = False
            hide2.Visible = False
            hide3.Visible = False

        ElseIf Session("tusuario") = 1 Then
            hide1.Visible = True
            hide2.Visible = True
            hide3.Visible = True
        End If
    End Sub

    Protected Sub Logout(sender As Object, e As EventArgs)
        Session.Clear()
        Response.Redirect("Ingreso")
    End Sub
End Class