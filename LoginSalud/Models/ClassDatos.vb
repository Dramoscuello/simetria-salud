Imports System.Data
Imports MySql.Data.MySqlClient
Public Class ClassDatos

    Dim conect As New ClassConexion
    Dim oComando As MySqlCommand

    Dim conexion As String = conect.CrearConexion.ConnectionString


    Public Function Datos_Llamar() As DataSet
        Try
            Dim myData As New DataSet
            Dim myAdapter As New MySqlDataAdapter
            Dim Conectar_ As New MySqlConnection(conexion)
            Conectar_.Open()
            Dim cmd As New MySqlCommand
            cmd.Connection = Conectar_
            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandText = "llamarDatos"
            myAdapter.SelectCommand = cmd
            myAdapter.Fill(myData)
            Return myData
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        End Try
    End Function
End Class
