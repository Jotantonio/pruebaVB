Imports System.Data.OleDb
Imports System.Globalization


Public Class frmActualizarDato

#Region "Propiedades"

    ''Variables
    Private mCodigoArticulo As Integer
    Dim miClsFuncionalidades As New clsFuncionalidades

#End Region

    ' Constructor que recibe el ID del registro a editar
    Public Sub New(ByVal pCodArticulo As Integer)
        InitializeComponent()
        mCodigoArticulo = pCodArticulo
    End Sub

    Private Sub cmdCancelar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancelar.Click
        Me.Close()
    End Sub

    Private Sub cmdAceptar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAceptar.Click

        Try
            Dim query As String

            DatabaseConnection.Instance.Open()

            If lblActual.Text = "Stock Actual" Then
                query = "UPDATE Articulos SET ExistActuales = ? WHERE CodigoArticulo = ?"
            ElseIf lblActual.Text = "Precio Compra Actual" Then
                query = "UPDATE Articulos SET PrecioCompra = ? WHERE CodigoArticulo = ?"
            ElseIf lblActual.Text = "Precio Venta Actual" Then
                query = "UPDATE Articulos SET PrecioVenta = ? WHERE CodigoArticulo = ?"
            Else
                MessageBox.Show("No se actualizará nada al no haber un dato a actualizar", "Actualizar artículo", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End If

            Using command As New OleDb.OleDbCommand(query, DatabaseConnection.Instance.Connection)
                If lblActual.Text = "Stock Actual" Then
                    command.Parameters.AddWithValue("?", Convert.ToInt32(txtNuevo.Text.Replace(".", ",").Replace(",", "")))
                Else
                    command.Parameters.AddWithValue("?", Convert.ToDouble(txtNuevo.Text.Replace(".", ","), CultureInfo.CurrentCulture))
                End If

                command.Parameters.AddWithValue("?", Convert.ToDecimal(mCodigoArticulo))

                command.ExecuteNonQuery()
            End Using

            Me.Close()

        Catch ex As Exception
            ' Manejo de errores
            miClsFuncionalidades.InsertarenLog(Me.Name, "cargarComboMarcas", ex.Message.ToString)
            MessageBox.Show("Error al actualizar los datos: " & ex.Message)
        Finally
            If DatabaseConnection.Instance.Connection.State = ConnectionState.Open Then
                DatabaseConnection.Instance.Close()
            End If
        End Try
    End Sub


End Class