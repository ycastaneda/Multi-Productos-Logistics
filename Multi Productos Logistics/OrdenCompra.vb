Imports Datos
Imports System.Data.SqlClient

Public Class OrdenCompra

    Private Sub OrdenCompra_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        obtenerProductos()
    End Sub

    Private Sub obtenerProductos()
        Try
            Dim datos As New DatosGenerales
            Dim ds As New DataSet
            ds = datos.obtenerProductos("")
            If ds.Tables(0).Rows.Count > 0 Then
                For i As Integer = 0 To ds.Tables(0).Rows.Count - 1
                    col_producto.DataSource = ds.Tables(0)
                    col_producto.DisplayMember = ds.Tables(0).Columns(3).ToString()
                    col_producto.ValueMember = ds.Tables(0).Columns(1).ToString()
                Next
                ds.Dispose()
            End If
        Catch ex As Exception

        End Try

    End Sub

End Class