Imports System.Data
Imports System.Data.SqlClient
Public Class ClaseConexion
    Dim cn As SqlConnection
    Public Shared cadenaconexion As String = "Server=YONATHAN\SQLEXPRESS;Database=multiproductoslogistics;Trusted_Connection=True;"
    'Public Shared cadenaconexion As String = "Server=mi_ip;uid=mi_userdb;pwd=mi_clavedb;database=mibasededatos"
    
    Public Function Conectar() As SqlConnection

        Try
            ' CADENA PARA CONEXION 
            cn = New SqlConnection(cadenaconexion)
            cn.Open()
            Return cn
        Catch ex As SqlException
            'MessageBox.Show("Error de Conexion")
            Return cn
        Catch ex As InvalidCastException
            'MessageBox.Show("Error de Conexion")
            Return cn
        Catch ex As IndexOutOfRangeException
            Return cn
        End Try

    End Function
    Public Sub desconectar(con As SqlConnection)
        Dim cn As SqlConnection = Nothing
        Try
            con.Close()
        Catch ex As SqlException

        Catch ex As InvalidCastException

        Catch ex As IndexOutOfRangeException

        End Try
    End Sub

    Public Function Consulta(cadena)
        'funcion para realizar una consulta que devuelve un solo valor, la funcion retorna una variable con el valor de la consulta
        Try
            Dim cm As SqlCommand
            Dim x As String
            Dim cn As SqlConnection
            cm = New SqlCommand(cadena)
            'La consulta de la línea anterior debe devolver únicamente un registro
            cn = Conectar()
            cm.Connection = cn
            'x = cm.ExecuteNonQuery()
            x = cm.ExecuteScalar().ToString()
            Return x
        Catch ex As SqlException
            MsgBox("Error en Base de Datos, codigo: " + ex.ErrorCode.ToString())
            'MsgBox("Error en Base de Datos, codigo: " + ex.Message.ToString())
            Return False
        Catch ex As InvalidCastException
            Return False
        Catch ex As IndexOutOfRangeException
            Return False
        End Try

    End Function

    Public Function ConsultarSetDatos(cadena As String)
        'funcion para realiza consultas de multiples campos a seleccionar de una base de datos retorna un dataadapter
        Dim da As SqlDataAdapter
        Dim con As SqlConnection
        con = Conectar()
        Dim cmd As SqlCommand = New SqlCommand(cadena, con)
        Try

            da = New SqlDataAdapter(cmd)
            con.Close()
            Return da
        Catch ex As SqlException
            MsgBox("Error en Base de Datos, codigo: " + ex.ErrorCode.ToString())
            Return False
        Catch ex As InvalidCastException
            Return False
        Catch ex As IndexOutOfRangeException
            Return False
            'Catch ex As Exception
            '    MsgBox(ex.Message)
            '    Return False
        Finally
            con.Close()
        End Try

    End Function

    Private Function insertar(ByVal cadenaSql As String)
        Dim con As SqlConnection = Nothing
        'funcion para insertar datos a la base de datos, se recibe la cadena sql y se procesa. retorna valor booleano
        Try
            ''inicializamos el comando sql pasando la cadena que recibimos del formulario y la conexion a la base

            con = Conectar()
            Dim cmd As SqlCommand = New SqlCommand(cadenaSql, con)
            cmd.ExecuteNonQuery()
            Return True
            'Catch ex As Exception
            '    'MsgBox(ex.Message)
            '    Return False
        Catch ex As SqlException
            MsgBox("Error en Base de Datos, codigo: " + ex.ErrorCode.ToString())
            Return False
        Catch ex As InvalidCastException
            Return False
        Catch ex As IndexOutOfRangeException
            Return False
        Finally
            con.Close()
        End Try
    End Function

    

End Class
