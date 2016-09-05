Imports MySql.Data.MySqlClient
Public Class Conexion

    Public Function CrearConexion() As MySqlConnection
        Try
            Dim dato_Con As String = "server=localhost;user id=root;password='0101';port=3306;database=dblevalidamos;CHARSET=utf8; Allow User Variables=True;"
            Dim MySqlConexion As New MySqlConnection(dato_Con)
            Return MySqlConexion
        Catch ex As Exception
            MsgBox("Error de conexion", MsgBoxStyle.Information, "Validacion")
            Return Nothing
        End Try
    End Function




End Class
