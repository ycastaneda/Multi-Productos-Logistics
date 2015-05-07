Imports System.Data.SqlClient
Imports System.Net

Public Class DatosGenerales
    Dim conexion As New ClaseConexion
    Dim cn As SqlConnection
    Dim ds As DataSet

    Public Function buscarRango(empresa As String, fechaInicial As String, fechaFinal As String) As DataSet
        Dim ds As New DataSet
        Dim dr As SqlDataReader
        Try
            Dim conexion As New ClaseConexion
            Dim comando As New SqlCommand
            Dim con As SqlConnection
            Dim tabla As New DataTable
            con = conexion.Conectar()
            Using con
                comando.CommandText = "select * from....."
                comando.Connection = con
                comando.Parameters.Add("@empresa", SqlDbType.VarChar).Value = empresa
                comando.Parameters.Add("@fechaInicial", SqlDbType.Date).Value = fechaInicial
                comando.Parameters.Add("@fechaFinal", SqlDbType.Date).Value = fechaFinal
                dr = comando.ExecuteReader()
                tabla.Load(dr)
                ds.Tables.Add(tabla)
            End Using
            Return ds
        Catch ex As Exception
            Return ds
        End Try
    End Function

    Public Function obtenerProductos(empresa As String) As DataSet
        Dim ds As New DataSet
        Dim dr As SqlDataReader
        Try
            Dim conexion As New ClaseConexion
            Dim comando As New SqlCommand
            Dim tabla As New DataTable
            Dim con As SqlConnection
            con = conexion.Conectar()
            Using con
                comando.CommandText = "select cod_empresa, cod_producto, unidad_medida, descripcion, tipo_producto, stock_minimo, precio_costo_unit, precio_venta_unit, existencia from purif_producto where cod_empresa=@empresa order by descripcion asc"
                comando.Connection = con
                comando.Parameters.Add("@empresa", SqlDbType.VarChar).Value = empresa
                dr = comando.ExecuteReader
                tabla.Load(dr)
                ds.Tables.Add(tabla)
            End Using
            Return ds
        Catch ex As Exception
            Return ds
        End Try
    End Function

    Public Function insertarDatos(grupo As String, empleado As String, usuario As String, otro As String) As Boolean
        Try
            Dim conexion As New ClaseConexion
            Dim comando As New SqlCommand
            Dim cn As SqlConnection
            cn = conexion.Conectar()
            Dim i As Integer = 0
            Using cn
                comando.CommandText = "insert into ...."
                comando.Connection = cn
                comando.Parameters.Add("@grupo", SqlDbType.VarChar).Value = grupo
                comando.Parameters.Add("@empleado", SqlDbType.VarChar).Value = empleado
                comando.Parameters.Add("@usuario", SqlDbType.VarChar).Value = usuario
                comando.Parameters.Add("@otro", SqlDbType.VarChar).Value = otro
                comando.ExecuteScalar()
            End Using
            Return True
        Catch ex As Exception
            'Console.WriteLine(ex.Message)
            Return False
        End Try
    End Function

    Public Function obtenerRoles() As DataSet
        'funcion general para retornar un dataset con los roles del sistema
        Dim ds As New DataSet
        Try
            Dim conexion As New ClaseConexion
            Dim adapter As New SqlDataAdapter
            adapter = conexion.ConsultarSetDatos("select .... from gen_rol")
            adapter.Fill(ds)
            Return ds
        Catch ex As NullReferenceException
            Return ds
        End Try
    End Function

    Public Function guardarEncabezado(comando As SqlCommand, empresa As String, numPartida As String, fecha As String, tipoPartida As String, descripcionPartida As Object) As Boolean
        ' Se recibe un SqlCommand con una transaccion implicita, utilizado para facturas y procedimientos que ameriten uso de transacciones
        Try
            comando.CommandText = "insert into ....."
            comando.Parameters.Add("@empresa", SqlDbType.Char).Value = empresa
            comando.Parameters.Add("@numPartida", SqlDbType.Int).Value = Integer.Parse(numPartida)
            comando.Parameters.Add("@fecha", SqlDbType.Date).Value = fecha
            comando.Parameters.Add("@tipoPartida", SqlDbType.Char).Value = tipoPartida
            comando.Parameters.Add("@descripcion", SqlDbType.VarChar).Value = descripcionPartida
            comando.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            'Console.WriteLine(ex.Message)
            Return False
        End Try
    End Function

    Public Function guardarDetalle(comando As SqlCommand, empresa As String, numPartida As String, numLinea As String, cuenta As String, debe As String, haber As String, descripcion As Object) As Boolean
        Try
            comando.CommandText = "insert into blabla_detalle() values()"
            comando.Parameters.Add("@empresa", SqlDbType.VarChar).Value = empresa
            comando.Parameters.Add("@partida", SqlDbType.Int).Value = Integer.Parse(numPartida)
            comando.Parameters.Add("@linea", SqlDbType.SmallInt).Value = Integer.Parse(numLinea)
            comando.Parameters.Add("@cuenta", SqlDbType.VarChar).Value = cuenta
            comando.Parameters.Add("@debe", SqlDbType.Decimal).Value = Decimal.Parse(debe)
            comando.Parameters.Add("@haber", SqlDbType.Decimal).Value = Decimal.Parse(haber)
            comando.Parameters.Add("@descripcion", SqlDbType.VarChar).Value = descripcion
            comando.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            'Console.WriteLine(ex.Message)
            Return False
        End Try
    End Function


End Class
