Imports System.Data
Imports MySql.Data.MySqlClient
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
Public Class Pasar_ErroresRips

    Dim conect As New ClassConexion
    Dim oComando As MySqlCommand

    Dim conexion As String = conect.CrearConexion.ConnectionString
    Dim tbl As New DataTable

    Public Function Obtener_Errores_AF(ByVal id As String) As DataTable
        tbl = New DataTable
        Using conn As New MySqlConnection(conexion)
            conn.Open()
            Dim query As String = "SELECT error_af_.* FROM error_af_ where Usuario='" & id & "';"
            Dim cmd As New MySqlDataAdapter(query, conn)
            cmd.Fill(tbl)
        End Using
        Return tbl
    End Function
    Public Function Obtener_Errores_CA(ByVal id As String) As DataTable
        tbl = New DataTable
        'Dim lista As New List(Of CERIPS.Errr_AC_CE)()
        Using conn As New MySqlConnection(conexion)
            conn.Open()
            Dim query As String = "SELECT error_ac_.* FROM error_ac_ where Usuario='" & "02" & "';"
            Dim cmd As New MySqlDataAdapter(query, conn)
            cmd.Fill(tbl)
        End Using
        Return tbl
    End Function
    Public Function Obtener_Errores_AH(ByVal id As String) As DataTable
        tbl = New DataTable
        'Dim lista As New List(Of CERIPS.ERROR_AH_CE)()
        Using conn As New MySqlConnection(conexion)
            conn.Open()
            Dim query As String = "SELECT error_ah_.* FROM error_ah_ where Usuario='" & id & "';"
            Dim cmd As New MySqlDataAdapter(query, conn)
            cmd.Fill(tbl)
        End Using
        Return tbl
    End Function
    Public Function Obtener_Errores_AM(ByVal id As String) As DataTable
        tbl = New DataTable
        'Dim lista As New List(Of CERIPS.ERROR_AM_CE)()
        Using conn As New MySqlConnection(conexion)
            conn.Open()
            Dim query As String = "SELECT error_am_.* FROM error_am_ where Usuario='" & id & "';"
            Dim cmd As New MySqlDataAdapter(query, conn)
            cmd.Fill(tbl)
        End Using
        Return tbl
    End Function
    Public Function Obtener_Errores_AN(ByVal id As String) As DataTable
        tbl = New DataTable
        'Dim lista As New List(Of CERIPS.ERROR_AN_CE)()
        Using conn As New MySqlConnection(conexion)
            conn.Open()
            Dim query As String = "SELECT error_an_.* FROM error_an_ where Usuario='" & id & "';"
            Dim cmd As New MySqlDataAdapter(query, conn)
            cmd.Fill(tbl)
        End Using
        Return tbl
    End Function
    Public Function Obtener_Errores_AP(ByVal id As String) As DataTable
        tbl = New DataTable
        'Dim lista As New List(Of CERIPS.ERROR_AP_CE)()
        Using conn As New MySqlConnection(conexion)
            conn.Open()
            Dim query As String = "SELECT error_ap_.* FROM error_ap_ where Usuario='" & id & "';"
            Dim cmd As New MySqlDataAdapter(query, conn)
            cmd.Fill(tbl)
        End Using
        Return tbl
    End Function
    Public Function Obtener_Errores_AT(ByVal id As String) As DataTable
        tbl = New DataTable
        'Dim lista As New List(Of CERIPS.ERROR_AT_CE)()
        Using conn As New MySqlConnection(conexion)
            conn.Open()
            Dim query As String = "SELECT error_at_.* FROM error_at_ where Usuario='" & id & "';"
            Dim cmd As New MySqlDataAdapter(query, conn)
            cmd.Fill(tbl)
        End Using
        Return tbl
    End Function
    Public Function Obtener_Errores_AU(ByVal id As String) As DataTable
        tbl = New DataTable
        'Dim lista As New List(Of CERIPS.ERROR_AU_CE)()
        Using conn As New MySqlConnection(conexion)
            conn.Open()
            Dim query As String = "SELECT error_au_.* FROM error_au_ where Usuario='" & id & "';"
            Dim cmd As New MySqlDataAdapter(query, conn)
            cmd.Fill(tbl)
        End Using
        Return tbl
    End Function
    Public Function Obtener_Errores_US(ByVal id As String) As DataTable
        tbl = New DataTable
        'Dim lista As New List(Of CERIPS.ERROR_US_CE)()
        Using conn As New MySqlConnection(conexion)
            conn.Open()
            Dim query As String = "SELECT error_us_.* FROM error_us_ where Usuario='" & id & "';"
            Dim cmd As New MySqlDataAdapter(query, conn)
            cmd.Fill(tbl)
        End Using
        Return tbl
    End Function
    Public Function Obtener_Total_Facturado(ByVal id As String) As DataTable
        tbl = New DataTable
        Using conn As New MySqlConnection(conexion)
            conn.Open()
            Dim query As String = "SELECT TIPO_IDENT, IDENTIFICACION, NOMBRE, EDAD, UNIDAD_EDAD, SEXO, PRESTADOR, REGIMEN, ENTIDAD, FECHA_FACT, FECHA_AFI, `Nº FACTURA`, TOTAL_AF, COPAGO, COMISION, `10%` DCTO, DIF, A_PAGAR , DX, DESCRIPCION FROM TOTAL_FACTURADO WHERE USUARIO='" & id & "';"
            Dim cmd As New MySqlDataAdapter(query, conn)
            cmd.Fill(tbl)
        End Using
        Return tbl
    End Function

End Class
