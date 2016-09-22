Imports System.Web.Routing
Public Class Global_asax
    Inherits HttpApplication

    Sub Application_Start(sender As Object, e As EventArgs)
        ' Fires when the application is started
        ValidationSettings.UnobtrusiveValidationMode = System.Web.UI.UnobtrusiveValidationMode.None
        RegisterRoutes(RouteTable.Routes)
    End Sub

    Private Shared Sub RegisterRoutes(routes As RouteCollection)
        routes.MapPageRoute("Ingreso", "Ingreso", "~/Default.aspx")
        routes.MapPageRoute("Validacion1479", "Validacion1479", "~/Views/Default1479/1479.aspx")
        routes.MapPageRoute("Principal", "Principal", "~/Views/Prueba.aspx")
        routes.MapPageRoute("ValidacionRips", "ValidacionRips", "~/Views/Rips/ValidacionRips.aspx")
    End Sub
End Class