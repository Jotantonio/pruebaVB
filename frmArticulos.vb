Imports System.Data.OleDb
Imports GestionPeluquería.clsFuncionalidades

Public Class frmArticulos

#Region "variables"

    Dim mBaja As Boolean = False
    Dim miClsFuncionalidades As New clsFuncionalidades

#End Region

#Region "eventos"

    Private Sub frmArticulos_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        ' Siempre maximizar el formulario cuando se active
        Me.WindowState = FormWindowState.Maximized
        Me.PerformLayout()
    End Sub

    Private Sub frmArticulos_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            lblNumRegistrosArticulos.Text = ""

            'Se carga con el índice 0 que son todos los artículos
            CargarDatosFamilia(0)

            cboMarca.DataSource = Nothing
            cboMarca.Items.Clear()

            cboFamilia.DataSource = Nothing
            cboFamilia.Items.Clear()

            cboGrupo.DataSource = Nothing
            cboGrupo.Items.Clear()

            cboSubgrupo.DataSource = Nothing
            cboSubgrupo.Items.Clear()

            cargarComboMarcas()

        Catch ex As Exception
            miClsFuncionalidades.InsertarenLog(Me.Name, "form_load", ex.Message.ToString)
            MessageBox.Show("Error en la carga de datos: " & ex.Message)
        End Try


    End Sub

#Region "botones"

    Private Sub btnSalir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSalir.Click
        Me.Close()
    End Sub

    ''' <summary>
    ''' abre el formulario de añadir artículo
    ''' </summary>
    Private Sub btnAniadir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAniadir.Click
        Try
            'Creamos una nueva instancia del formulario pasándole codigo 0 al ser nuevo
            Dim mfrmAniadirArticulo As New frmAniadirArticulo(0)

            mfrmAniadirArticulo.BringToFront()
            mfrmAniadirArticulo.ShowDialog()

            If mfrmAniadirArticulo.SalirSinModificar = False Then
                cargarTodosCombos()
            End If
        Catch ex As Exception
            miClsFuncionalidades.InsertarenLog(Me.Name, "btn_aniadir_click", ex.Message.ToString)
            MessageBox.Show("Error en la carga de datos: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' método para borrar un artículo
    ''' </summary>
    Private Sub btnBorrar_Click(sender As System.Object, e As System.EventArgs) Handles btnBorrar.Click

        Dim mCodArticulo As String = ""
        Dim clave As Integer
        Try
            mCodArticulo = dgArticulos.CurrentRow.Cells("CodigoArticulo").Value.ToString()

            If mCodArticulo <> "" Then
                If (MessageBox.Show("El artículo con código " & mCodArticulo & ", se va a dar de baja. Desea continua?", "Dar de baja un artículo", MessageBoxButtons.YesNo)) = Windows.Forms.DialogResult.Yes Then

                    DatabaseConnection.Instance.Open()

                    ' Crear el DataAdapter
                    Dim query As String = "DELETE FROM ARTICULOS Where CodigoArticulo = ?;"

                    ' Crear un OleDbCommand
                    Using command As New OleDbCommand(query, DatabaseConnection.Instance.Connection)
                        ' Agregar el parámetro ID
                        command.Parameters.AddWithValue("?", mCodArticulo)

                        ' Ejecutar el comando
                        Dim filasAfectadas As Integer = command.ExecuteNonQuery()

                        If filasAfectadas > 0 Then
                            MessageBox.Show("Registro eliminado correctamente.")
                        Else
                            MessageBox.Show("No se encontró el registro.")
                        End If
                    End Using
                    DatabaseConnection.Instance.Close()

                    If cboSubgrupo.SelectedIndex < 1 Then
                        If cboGrupo.SelectedIndex < 1 Then
                            If cboFamilia.SelectedIndex < 1 Then
                                If cboMarca.SelectedIndex < 1 Then
                                    clave = 0
                                Else
                                    clave = cboMarca.SelectedValue
                                End If
                                CargarDatosMarca(clave)
                            Else
                                clave = cboFamilia.SelectedValue
                                CargarDatosFamilia(clave)
                            End If
                        Else
                            clave = cboGrupo.SelectedValue
                            CargarDatosGrupo(clave)
                        End If
                    Else
                        clave = cboSubgrupo.SelectedValue
                        CargarDatosSubgrupo(clave)
                    End If


                End If
            End If
        Catch ex As Exception
            ' Manejo de errores
            miClsFuncionalidades.InsertarenLog(Me.Name, "btnBorrar_click", ex.Message.ToString)
            MessageBox.Show("Error al dar de baja el artículo: " & ex.Message)
        Finally
            If DatabaseConnection.Instance.Connection.State = ConnectionState.Open Then
                DatabaseConnection.Instance.Close()
            End If
        End Try
    End Sub

    ''' <summary>
    ''' metodo para modificar un artículo seleccionado en el grid
    ''' </summary>
    Private Sub btnEditar_Click(sender As System.Object, e As System.EventArgs) Handles btnEditar.Click
        Try
            Dim mCodArticulo As String
            mCodArticulo = dgArticulos.CurrentRow.Cells("CodigoArticulo").Value.ToString()
            Dim mfrmAniadirArticulo As New frmAniadirArticulo(mCodArticulo)
            mfrmAniadirArticulo.BringToFront()
            mfrmAniadirArticulo.ShowDialog()

            If mfrmAniadirArticulo.SalirSinModificar = False Then
                cargarTodosCombos()
            End If
        Catch ex As Exception
            ' Manejo de errores
            miClsFuncionalidades.InsertarenLog(Me.Name, "btnEditar_click", ex.Message.ToString)
            MessageBox.Show("Error al editar el artículo: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' edición del stock del registro seleccionado
    ''' </summary>
    Private Sub cmdStock_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStock.Click

        Dim mCodArticulo As String
        Dim mNombreArticulo As String
        Dim mStockActual As Integer

        mNombreArticulo = dgArticulos.CurrentRow.Cells("Nombre").Value.ToString()
        mCodArticulo = dgArticulos.CurrentRow.Cells("CodigoArticulo").Value.ToString()

        Try
            DatabaseConnection.Instance.Open()

            ' Crear el DataAdapter
            Dim query As String = "SELECT ARTICULOS.ExistActuales FROM ARTICULOS Where CodigoArticulo = ?;"

            ' Crear un OleDbCommand
            Using command As New OleDbCommand(query, DatabaseConnection.Instance.Connection)
                ' Agregar el parámetro ID
                command.Parameters.AddWithValue("?", mCodArticulo)

                ' Ejecutar la consulta y leer el resultado
                Using reader As OleDbDataReader = command.ExecuteReader()
                    If reader.Read() Then
                        ' Recuperar el valor del campo
                        mStockActual = reader("ExistActuales").ToString()
                    End If
                End Using
            End Using
            DatabaseConnection.Instance.Close()

            Dim mfrmActualizarDato = New frmActualizarDato(mCodArticulo)
            mfrmActualizarDato.Text = mNombreArticulo
            mfrmActualizarDato.lblActual.Text = "Stock Actual"
            mfrmActualizarDato.lblNuevo.Text = "Stock Nuevo"
            mfrmActualizarDato.txtActual.Text = mStockActual
            mfrmActualizarDato.txtNuevo.Text = mStockActual
            mfrmActualizarDato.BringToFront()
            mfrmActualizarDato.ShowDialog()

            cargarTodosCombos()

        Catch ex As Exception
            ' Manejo de errores
            miClsFuncionalidades.InsertarenLog(Me.Name, "btnStock_click", ex.Message.ToString)
            MessageBox.Show("Error al actualizar el stock: " & ex.Message)
        Finally
            If DatabaseConnection.Instance.Connection.State = ConnectionState.Open Then
                DatabaseConnection.Instance.Close()
            End If
        End Try

    End Sub

    ''' <summary>
    ''' edición del precio de compra del articulo seleccionado
    ''' </summary>
    Private Sub cmdPrecioCompra_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPrecioCompra.Click

        Dim mCodArticulo As String
        Dim mNombreArticulo As String
        Dim mValorActual As Double

        Try
            mNombreArticulo = dgArticulos.CurrentRow.Cells("Nombre").Value.ToString()
            mCodArticulo = dgArticulos.CurrentRow.Cells("CodigoArticulo").Value.ToString()

            DatabaseConnection.Instance.Open()

            ' Crear el DataAdapter
            Dim query As String = "SELECT ARTICULOS.PrecioCompra FROM ARTICULOS Where CodigoArticulo = ?;"

            ' Crear un OleDbCommand
            Using command As New OleDbCommand(query, DatabaseConnection.Instance.Connection)
                ' Agregar el parámetro ID
                command.Parameters.AddWithValue("?", mCodArticulo)

                ' Ejecutar la consulta y leer el resultado
                Using reader As OleDbDataReader = command.ExecuteReader()
                    If reader.Read() Then
                        ' Recuperar el valor del campo
                        mValorActual = reader("PrecioCompra").ToString()
                    End If
                End Using
            End Using
            DatabaseConnection.Instance.Close()

            Dim mfrmActualizarDato = New frmActualizarDato(mCodArticulo)
            mfrmActualizarDato.Text = mNombreArticulo
            mfrmActualizarDato.lblActual.Text = "Precio Compra Actual"
            mfrmActualizarDato.lblNuevo.Text = "Precio Compra Nuevo"
            mfrmActualizarDato.txtActual.Text = mValorActual
            mfrmActualizarDato.txtNuevo.Text = mValorActual
            mfrmActualizarDato.BringToFront()
            mfrmActualizarDato.ShowDialog()

            cargarTodosCombos()

        Catch ex As Exception
            ' Manejo de errores
            miClsFuncionalidades.InsertarenLog(Me.Name, "cmdPrecioCompra", ex.Message.ToString)
            MessageBox.Show("Error al actualizar el precio de compra: " & ex.Message)
        Finally
            If DatabaseConnection.Instance.Connection.State = ConnectionState.Open Then
                DatabaseConnection.Instance.Close()
            End If
        End Try

    End Sub

    ''' <summary>
    ''' edición del precio de venta del artículo seleccionado
    ''' </summary>
    Private Sub cmdPVP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPVP.Click

        Dim mCodArticulo As String
        Dim mNombreArticulo As String
        Dim mValorActual As Double

        Try
            mNombreArticulo = dgArticulos.CurrentRow.Cells("Nombre").Value.ToString()
            mCodArticulo = dgArticulos.CurrentRow.Cells("CodigoArticulo").Value.ToString()


            DatabaseConnection.Instance.Open()

            ' Crear el DataAdapter
            Dim query As String = "SELECT ARTICULOS.PrecioVenta FROM ARTICULOS Where CodigoArticulo = ?;"

            ' Crear un OleDbCommand
            Using command As New OleDbCommand(query, DatabaseConnection.Instance.Connection)
                ' Agregar el parámetro ID
                command.Parameters.AddWithValue("?", mCodArticulo)

                ' Ejecutar la consulta y leer el resultado
                Using reader As OleDbDataReader = command.ExecuteReader()
                    If reader.Read() Then
                        ' Recuperar el valor del campo
                        mValorActual = reader("PrecioVenta").ToString()
                    End If
                End Using
            End Using
            DatabaseConnection.Instance.Close()

            Dim mfrmActualizarDato = New frmActualizarDato(mCodArticulo)
            mfrmActualizarDato.Text = mNombreArticulo
            mfrmActualizarDato.lblActual.Text = "Precio Venta Actual"
            mfrmActualizarDato.lblNuevo.Text = "Precio Venta Nuevo"
            mfrmActualizarDato.txtActual.Text = mValorActual
            mfrmActualizarDato.txtNuevo.Text = mValorActual
            mfrmActualizarDato.BringToFront()
            mfrmActualizarDato.ShowDialog()

            cargarTodosCombos()

        Catch ex As Exception
            ' Manejo de errores
            miClsFuncionalidades.InsertarenLog(Me.Name, "cmdPVP_click", ex.Message.ToString)
            MessageBox.Show("Error al actualizar el precio de venta: " & ex.Message)
        Finally
            If DatabaseConnection.Instance.Connection.State = ConnectionState.Open Then
                DatabaseConnection.Instance.Close()
            End If
        End Try

    End Sub

    ''' <summary>
    ''' se mueve al anterior registro del grid
    ''' </summary>
    Private Sub btnAnterior_Click(sender As System.Object, e As System.EventArgs) Handles btnAnterior.Click
        If dgArticulos.CurrentRow.Index > 0 And dgArticulos.Rows.Count > 0 Then
            dgArticulos.CurrentCell = dgArticulos.Rows(dgArticulos.CurrentRow.Index - 1).Cells(0)
        End If
    End Sub

    ''' <summary>
    ''' se mueve al siguiente registro del grid
    ''' </summary>
    Private Sub btnSiguiente_Click(sender As System.Object, e As System.EventArgs) Handles btnSiguiente.Click
        If dgArticulos.CurrentRow.Index < dgArticulos.Rows.Count - 1 And dgArticulos.Rows.Count > 0 Then
            dgArticulos.CurrentCell = dgArticulos.Rows(dgArticulos.CurrentRow.Index + 1).Cells(0)
        End If
    End Sub

    ''' <summary>
    ''' se mueve al primer registro del grid
    ''' </summary>
    Private Sub btnPrimero_Click(sender As System.Object, e As System.EventArgs) Handles btnPrimero.Click
        If dgArticulos.Rows.Count > 0 Then
            dgArticulos.CurrentCell = dgArticulos.Rows(0).Cells(0)
        End If
    End Sub

    ''' <summary>
    ''' se mueve al ultimo registro del grid
    ''' </summary>
    Private Sub btnUltimo_Click(sender As System.Object, e As System.EventArgs) Handles btnUltimo.Click
        If dgArticulos.Rows.Count > 0 Then
            dgArticulos.CurrentCell = dgArticulos.Rows(dgArticulos.Rows.Count - 1).Cells(0)
        End If
    End Sub

    ''' <summary>
    ''' botón para ver las altas o las bajas, cargando los registros de alta o baja filtrado por los combos
    ''' </summary>
    Private Sub btnVerAlta_Click(sender As System.Object, e As System.EventArgs) Handles btnVerAlta.Click

        Dim clave As Integer
        Dim mMarca As Integer
        Try
            If btnVerAlta.Text = "Alta" Then
                btnVerAlta.Image = GestionPeluquería.My.Resources.Resources._09VerBajaTexto
                btnAlta.Image = GestionPeluquería.My.Resources.Resources._08BajaTexto
                btnVerAlta.Text = "Baja"
            Else
                btnVerAlta.Image = GestionPeluquería.My.Resources.Resources._09VerAltaTexto
                btnAlta.Image = GestionPeluquería.My.Resources.Resources._08AltaTexto
                btnVerAlta.Text = "Alta"
            End If

            If cboMarca.SelectedIndex > 0 Then
                mMarca = cboMarca.SelectedValue
            Else
                mMarca = 0
            End If

            If cboSubgrupo.SelectedIndex < 1 Then
                If cboGrupo.SelectedIndex < 1 Then
                    If cboFamilia.SelectedIndex < 1 Then
                        If cboMarca.SelectedIndex < 1 Then
                            clave = 0
                        Else
                            clave = cboMarca.SelectedValue
                        End If
                        CargarDatosMarca(clave)
                    Else
                        clave = cboFamilia.SelectedValue
                        CargarDatosFamilia(clave)
                    End If
                Else
                    clave = cboGrupo.SelectedValue
                    CargarDatosGrupo(clave)
                End If
            Else
                clave = cboSubgrupo.SelectedValue
                CargarDatosSubgrupo(clave)
            End If
        Catch ex As Exception
            ' Manejo de errores
            miClsFuncionalidades.InsertarenLog(Me.Name, "btnVerAlta_click", ex.Message.ToString)
            MessageBox.Show("Error no controlado al obtener información de altas/bajas de artículos: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' botón para dar de alta o de baja un producto
    ''' </summary>
    Private Sub btnAlta_Click(sender As System.Object, e As System.EventArgs) Handles btnAlta.Click
        Try
            Dim mCodigoArticulo As String
            Dim query As String
            Dim clave As Integer
            Dim mMarca As Integer

            mCodigoArticulo = dgArticulos.CurrentRow.Cells("CodigoArticulo").Value.ToString()

            'Si el botón es ver bajas, es que estamos en las altas, se quiere dar de baja una artículo
            If btnVerAlta.Text = "Baja" Then
                If (MessageBox.Show("¿Seguro de que quieres dar de baja el artículo?", "Alta/Baja artículos", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes) Then
                    query = "UPDATE Articulos SET Baja = Yes WHERE CodigoArticulo = ?"
                    DatabaseConnection.Instance.Open()

                    Using command As New OleDb.OleDbCommand(query, DatabaseConnection.Instance.Connection)

                        command.Parameters.AddWithValue("?", Convert.ToDecimal(mCodigoArticulo))
                        command.ExecuteNonQuery()
                    End Using
                End If
            Else
                If (MessageBox.Show("¿Seguro de que quieres dar de alta el artículo?", "Alta/Baja artículos", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes) Then
                    query = "UPDATE Articulos SET Baja = No WHERE CodigoArticulo = ?"
                    DatabaseConnection.Instance.Open()

                    Using command As New OleDb.OleDbCommand(query, DatabaseConnection.Instance.Connection)
                        command.Parameters.AddWithValue("?", Convert.ToDecimal(mCodigoArticulo))
                        command.ExecuteNonQuery()
                    End Using
                End If
            End If

            If cboMarca.SelectedIndex > 0 Then
                mMarca = cboMarca.SelectedValue
            Else
                mMarca = 0
            End If

            If cboSubgrupo.SelectedIndex < 1 Then
                If cboGrupo.SelectedIndex < 1 Then
                    If cboFamilia.SelectedIndex < 1 Then
                        If cboMarca.SelectedIndex < 1 Then
                            clave = 0
                        Else
                            clave = cboMarca.SelectedValue
                        End If
                        CargarDatosMarca(clave)
                    Else
                        clave = cboFamilia.SelectedValue
                        CargarDatosFamilia(clave)
                    End If
                Else
                    clave = cboGrupo.SelectedValue
                    CargarDatosGrupo(clave)
                End If
            Else
                clave = cboSubgrupo.SelectedValue
                CargarDatosSubgrupo(clave)
            End If

        Catch ex As Exception
            ' Manejo de errores
            miClsFuncionalidades.InsertarenLog(Me.Name, "btnVerAlta_click", ex.Message.ToString)
            MessageBox.Show("Error al actualizar alta/baja: " & ex.Message)
        Finally
            If DatabaseConnection.Instance.Connection.State = ConnectionState.Open Then
                DatabaseConnection.Instance.Close()
            End If
        End Try
    End Sub

    ''' <summary>
    ''' muestra la ayuda
    ''' </summary>
    Private Sub btnAyuda_Click(sender As System.Object, e As System.EventArgs) Handles btnAyuda.Click
        MessageBox.Show("En breve...", "Ayuda", MessageBoxButtons.OK)
    End Sub

    Private Sub btnImprimir_Click(sender As System.Object, e As System.EventArgs) Handles btnImprimir.Click
        MessageBox.Show("En breve...", "Imprimir", MessageBoxButtons.OK)
    End Sub

#End Region

#Region "textbox"

    ''' <summary>
    ''' busca por el nombre del texto incluido en el textbox de nombre
    ''' </summary>
    Private Sub txtNombre_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtNombre.TextChanged
        Try
            CargarArticulosPorNombre(txtNombre.Text)
        Catch ex As Exception
            ' Manejo de errores
            miClsFuncionalidades.InsertarenLog(Me.Name, "txtNombre_TextChanged", ex.Message.ToString)
            MessageBox.Show("Error al consultar articulos por el nombre: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' busca por el codigo de barras del texto incluido en el textbox de codigo de barras
    ''' </summary>
    Private Sub txtCodBarras_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtCodBarras.TextChanged
        Try
            CargarArticulosPorCodigoBarras(txtCodBarras.Text)
        Catch ex As Exception
            ' Manejo de errores
            miClsFuncionalidades.InsertarenLog(Me.Name, "txtCodBarras_TextChanged", ex.Message.ToString)
            MessageBox.Show("Error al consultar articulos por el codigo de barras: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' busca por el codigo de barras incluido en el textbox del codigo de barras
    ''' </summary>
    Private Sub txtCodBarras_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtCodBarras.KeyPress
        Try
            ' Verifica si la tecla presionada no es un dígito ni una tecla de control (como Backspace)
            If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
                e.Handled = True ' Ignora la tecla
            End If
        Catch ex As Exception
            ' Manejo de errores
            miClsFuncionalidades.InsertarenLog(Me.Name, "txtCodBarras_KeyPress", ex.Message.ToString)
            MessageBox.Show("Error al controlar el texto numérico introducido en la caja texto de codigo de barras: " & ex.Message)
        End Try
    End Sub

#End Region

#Region "combos"

    Private Sub cboMarca_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cboMarca.SelectedIndexChanged
        Try
            If cboMarca.SelectedIndex > 0 Then
                cboFamilia.DataSource = Nothing
                cboFamilia.Items.Clear()

                CargarComboFamilias(cboMarca.SelectedValue)
                CargarDatosMarca(cboMarca.SelectedValue)
            Else
                cboFamilia.DataSource = Nothing
                cboFamilia.Items.Clear()
                CargarDatosMarca(0)
            End If
        Catch ex As Exception
            ' Manejo de errores
            miClsFuncionalidades.InsertarenLog(Me.Name, "cboMarca_SelectedIndexChanged", ex.Message.ToString)
            MessageBox.Show("Error al cambiar la marca en la lista desplegable: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Cambia la familia del combo de familias y carga la rejilla con la familia seleccionada
    ''' </summary>
    Private Sub cboFamilia_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cboFamilia.SelectedIndexChanged
        Try

            Dim clave As Integer

            cboGrupo.DataSource = Nothing
            cboGrupo.Items.Clear()

            cboSubgrupo.DataSource = Nothing
            cboSubgrupo.Items.Clear()

            If cboFamilia.SelectedIndex < 1 Then
                If cboMarca.SelectedIndex < 1 Then
                    clave = 0
                Else
                    clave = cboMarca.SelectedValue
                End If
                CargarDatosMarca(clave)
                CargarComboFamilias(clave)
            Else
                clave = cboFamilia.SelectedValue
                CargarDatosFamilia(clave)
                CargarComboGrupos(clave)
            End If

        Catch ex As Exception
            miClsFuncionalidades.InsertarenLog(Me.Name, "cboFamilia_SelectedIndexChanged", ex.Message.ToString)
            MessageBox.Show("Error al cargar datos al cambiar la familia en la lista desplegable: " & ex.Message)
        End Try

    End Sub

    ''' <summary>
    ''' cambia el grupo del combo de grupos y carga la rejilla con el grupo seleccionado
    ''' </summary>
    Private Sub cboGrupo_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cboGrupo.SelectedIndexChanged
        Try
            Dim clave As Integer

            cboSubgrupo.DataSource = Nothing
            cboSubgrupo.Items.Clear()

            If cboGrupo.SelectedIndex = 0 Then
                clave = 0
            Else
                clave = cboGrupo.SelectedValue
            End If

            CargarComboSubgrupos(clave)
            CargarDatosGrupo(clave)
        Catch ex As Exception
            miClsFuncionalidades.InsertarenLog(Me.Name, "cboGrupo_SelectedIndexChanged", ex.Message.ToString)
            MessageBox.Show("Error al cargar datos al cambiar el grupo en la lista desplegable: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' cambia el subgrupo del combo de subgrupos y carga la rejilla con el subgrupo seleccionado
    ''' </summary>
    Private Sub cboSubgrupo_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cboSubgrupo.SelectedIndexChanged
        Try
            Dim clave As Integer

            If cboSubgrupo.SelectedIndex = 0 Then
                clave = 0
            Else
                clave = cboSubgrupo.SelectedValue
            End If

            CargarDatosSubgrupo(clave)
        Catch ex As Exception
            miClsFuncionalidades.InsertarenLog(Me.Name, "cboSubgrupo_SelectedIndexChanged", ex.Message.ToString)
            MessageBox.Show("Error al cargar datos: " & ex.Message)
        End Try
    End Sub

#End Region

    Private Sub EditarFila(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgArticulos.CellDoubleClick
        btnEditar_Click(Nothing, EventArgs.Empty)
    End Sub

#End Region

#Region "metodos privados"

    ''' <summary>
    ''' carga el combo de las marcas dadas de alta
    ''' </summary>
    Private Sub cargarComboMarcas()

        Dim query As String
        Dim items As New List(Of clsComboboxItem)
        Try
            'Abrimos conexión
            DatabaseConnection.Instance.Open()

            'montamos la query
            query = "SELECT Marcas.CodigoMarca as Id, Marcas.NombreMarca as Nombre FROM MARCAS WHERE MARCAS.CODIGOMARCA > 1"

            Using command As New OleDbCommand(query, DatabaseConnection.Instance.Connection)

                ' Ejecutar la consulta
                Using reader As OleDbDataReader = command.ExecuteReader()

                    'Incluir item vacio
                    Dim itemVacio As New clsComboboxItem With {
                            .Id = 0,
                            .Descripcion = ""}
                    items.Add(itemVacio)

                    ' Leer los datos y agregarlos al ComboBox
                    While reader.Read()
                        Dim item As New clsComboboxItem With {
                            .Id = Convert.ToInt32(reader("Id")),
                            .Descripcion = reader("Nombre").ToString()}
                        items.Add(item)

                        ' Agregar elementos al ComboBox (Texto visible y valor subyacente)
                        'cboFamilia.Items.Add(reader(1))
                    End While

                    cboMarca.DataSource = items
                    cboMarca.DisplayMember = "Descripcion"
                    cboMarca.ValueMember = "Id"

                    cboFamilia.DataSource = Nothing
                    cboFamilia.Items.Clear()

                End Using
            End Using

        Catch ex As Exception
            miClsFuncionalidades.InsertarenLog(Me.Name, "cargarComboMarcas", ex.Message.ToString)
            MessageBox.Show("Error al cargar datos: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' carga el combo de familias con la clave del combo del tipo que se pasa
    ''' </summary>
    ''' <param name="pCodigoMarca">tipo: producto o servicio</param>
    ''' <remarks></remarks>
    Private Sub CargarComboFamilias(ByVal pCodigoMarca As Integer)

        Dim query As String
        Dim items As New List(Of clsComboboxItem)
        Dim itemVacioIncluido As Boolean = False
        Dim numRegistros As Integer = 0

        Try
            'Productos
            query = "SELECT Familias.CodigoFamilia as Id, FAMILIAS.Nombre as Familia FROM FAMILIAS WHERE FAMILIAS.Servicio=No"

            If pCodigoMarca <> 0 Then
                query = query & " And Familias.CodigoMarca =" & pCodigoMarca
            End If

            ' Configurar el ComboBox
            cboGrupo.DataSource = Nothing
            cboGrupo.Items.Clear() ' Limpiar items existentes

            DatabaseConnection.Instance.Open()

            Using command As New OleDbCommand(query, DatabaseConnection.Instance.Connection)

                ' Ejecutar la consulta
                Using reader As OleDbDataReader = command.ExecuteReader()

                    ' Leer los datos y agregarlos al ComboBox
                    While reader.Read()

                        If itemVacioIncluido = False Then
                            Dim itemVacio As New clsComboboxItem With {
                            .Id = 1,
                            .Descripcion = ""}
                            items.Add(itemVacio)
                            itemVacioIncluido = True
                        End If

                        Dim item As New clsComboboxItem With {
                            .Id = Convert.ToInt32(reader("Id")),
                            .Descripcion = reader("Familia").ToString()}
                        items.Add(item)

                        numRegistros += 1
                        ' Agregar elementos al ComboBox (Texto visible y valor subyacente)
                        'cboFamilia.Items.Add(reader(1))
                    End While

                    If numRegistros > 0 Then
                        cboFamilia.DataSource = items
                        cboFamilia.DisplayMember = "Descripcion"
                        cboFamilia.ValueMember = "Id"
                    End If

                End Using
            End Using

        Catch ex As Exception
            miClsFuncionalidades.InsertarenLog(Me.Name, "CargarComboFamilias", ex.Message.ToString)
            MessageBox.Show("Error al cargar datos: " & ex.Message)
        Finally
            If DatabaseConnection.Instance.Connection.State = ConnectionState.Open Then
                DatabaseConnection.Instance.Close()
            End If
        End Try

    End Sub

    ''' <summary>
    ''' Carga el combo de grupos con la clave del combo de la familia que se pasa
    ''' </summary>
    ''' <param name="indice">codigo familia</param>
    ''' <remarks></remarks>
    Private Sub CargarComboGrupos(ByVal indice As Integer)
        Try
            Dim query As String
            Dim items As New List(Of clsComboboxItem)
            Dim itemVacioIncluido As Boolean = False
            Dim numRegistros As Integer = 0

            If indice = 0 Then
                'Productos
                query = "SELECT DISTINCT GRUPOS.CODIGOGRUPO AS Id, GRUPOS.NOMBRE As Grupo" & _
                        " FROM GRUPOS INNER JOIN ARTICULOS ON GRUPOS.CodigoGrupo = ARTICULOS.CodigoGrupo"
            Else
                query = "SELECT DISTINCT GRUPOS.CODIGOGRUPO AS Id, GRUPOS.NOMBRE As Grupo" & _
                        " FROM GRUPOS INNER JOIN ARTICULOS ON GRUPOS.CodigoGrupo = ARTICULOS.CodigoGrupo" & _
                        " WHERE ARTICULOS.CODIGOFAMILIA =" & indice
            End If

            ' Configurar el ComboBox
            cboGrupo.Items.Clear() ' Limpiar items existentes

            DatabaseConnection.Instance.Open()

            Using command As New OleDbCommand(query, DatabaseConnection.Instance.Connection)
                ' Ejecutar la consulta
                Using reader As OleDbDataReader = command.ExecuteReader()

                    ' Leer los datos y agregarlos al ComboBox
                    While reader.Read()

                        If itemVacioIncluido = False Then
                            Dim itemVacio As New clsComboboxItem With {
                            .Id = 1,
                            .Descripcion = ""}
                            items.Add(itemVacio)
                            itemVacioIncluido = True
                        End If

                        Dim item As New clsComboboxItem With {
                            .Id = Convert.ToInt32(reader("Id")),
                            .Descripcion = reader("Grupo").ToString()}
                        items.Add(item)

                        numRegistros += 1
                        ' Agregar elementos al ComboBox (Texto visible y valor subyacente)
                        'cboGrupo.Items.Add(reader(1))
                    End While

                    If numRegistros > 0 Then
                        cboGrupo.DataSource = items
                        cboGrupo.DisplayMember = "Descripcion"
                        cboGrupo.ValueMember = "Id"
                    End If

                End Using
            End Using

        Catch ex As Exception
            miClsFuncionalidades.InsertarenLog(Me.Name, "CargarComboGrupos", ex.Message.ToString)
            MessageBox.Show("Error al cargar datos: " & ex.Message)
        Finally
            If DatabaseConnection.Instance.Connection.State = ConnectionState.Open Then
                DatabaseConnection.Instance.Close()
            End If
        End Try

    End Sub

    ''' <summary>
    ''' Carga el combo de subgrupos con la clave del combo de grupos que se pasa
    ''' </summary>
    ''' <param name="indice">codigo grupo</param>
    ''' <remarks></remarks>
    Private Sub CargarComboSubgrupos(ByVal indice As Integer)

        Dim query As String
        Dim items As New List(Of clsComboboxItem)
        Dim itemVacioIncluido As Boolean = False
        Dim numRegistros As Integer = 0

        Try
            If indice = 0 Then
                'Productos
                query = "SELECT DISTINCT SUBGRUPOS.CODIGOSUBGRUPO as Id, SUBGRUPOS.NOMBRE as Subgrupo" & _
                        " FROM SUBGRUPOS INNER JOIN ARTICULOS ON SUBGRUPOS.CodigoSubgrupo = ARTICULOS.CodigoSubgrupo"
            Else
                query = "SELECT DISTINCT SUBGRUPOS.CODIGOSUBGRUPO as Id, SUBGRUPOS.NOMBRE as Subgrupo" & _
                        " FROM SUBGRUPOS INNER JOIN ARTICULOS ON SUBGRUPOS.CodigoSubgrupo = ARTICULOS.CodigoSubgrupo" & _
                        " WHERE ARTICULOS.CODIGOSUBGRUPO =" & indice
            End If

            ' Configurar el ComboBox
            cboSubgrupo.Items.Clear() ' Limpiar items existentes

            DatabaseConnection.Instance.Open()

            Using command As New OleDbCommand(query, DatabaseConnection.Instance.Connection)

                ' Ejecutar la consulta
                Using reader As OleDbDataReader = command.ExecuteReader()

                    ' Leer los datos y agregarlos al ComboBox
                    While reader.Read()

                        If itemVacioIncluido = False Then
                            Dim itemVacio As New clsComboboxItem With {
                            .Id = 1,
                            .Descripcion = ""}
                            items.Add(itemVacio)
                            itemVacioIncluido = True
                        End If

                        Dim item As New clsComboboxItem With {
                            .Id = Convert.ToInt32(reader("Id")),
                            .Descripcion = reader("Subgrupo").ToString()}
                        items.Add(item)

                        numRegistros += 1
                        ' Agregar elementos al ComboBox (Texto visible y valor subyacente)
                        'cboSubgrupo.Items.Add(reader(1))
                    End While

                    If numRegistros > 0 Then
                        cboSubgrupo.DataSource = items
                        cboSubgrupo.DisplayMember = "Descripcion"
                        cboSubgrupo.ValueMember = "Id"
                    End If
                End Using
            End Using
        Catch ex As Exception
            miClsFuncionalidades.InsertarenLog(Me.Name, "CargarComboSubgrupos", ex.Message.ToString)
            MessageBox.Show("Error al cargar datos: " & ex.Message)
        Finally
            If DatabaseConnection.Instance.Connection.State = ConnectionState.Open Then
                DatabaseConnection.Instance.Close()
            End If
        End Try

    End Sub

    ''' <summary>
    ''' Carga articulos con la clave del combo de tipos que se pasa
    ''' </summary>
    ''' <param name="mMarca">Tipo: producto o Servicio</param>
    ''' <remarks></remarks>
    Private Sub CargarDatosMarca(ByVal mMarca As String)
        Try
            Dim query As String = ""
            Dim flagAlta As String = ""

            'Si el botón es ver bajas, es que estamos en las altas, se muestran las altas
            If btnVerAlta.Text = "Baja" Then
                flagAlta = " AND ARTICULOS.BAJA = No"
            Else
                flagAlta = " AND ARTICULOS.BAJA = Yes"
            End If

            ' Abrir la conexión global
            DatabaseConnection.Instance.Open()
            
            'productos
            query = "SELECT ARTICULOS.CodigoArticulo, ARTICULOS.Referencia, ARTICULOS.CodigoBarras, ARTICULOS.Nombre, ARTICULOS.ExistMin," & _
                            " ARTICULOS.ExistMax, ARTICULOS.ExistActuales, ARTICULOS.PrecioCompra, ARTICULOS.PrecioVenta, FAMILIAS.Nombre as Familia," & _
                            " GRUPOS.Nombre as Grupo" & _
                            " FROM (FAMILIAS INNER JOIN ARTICULOS ON FAMILIAS.CodigoFamilia = ARTICULOS.CodigoFamilia)" & _
                            " INNER JOIN GRUPOS ON (GRUPOS.CodigoGrupo = ARTICULOS.CodigoGrupo)" & _
                            " WHERE FAMILIAS.SERVICIO = No" & flagAlta

            
            If mMarca <> "0" Then
                query = query & " And ARTICULOS.CodigoMarca = " & mMarca
            End If
            ' Crear el DataAdapter
            Dim adapter As New OleDbDataAdapter(query, DatabaseConnection.Instance.Connection)

            ' Crear el DataSet
            Dim dataSet As New DataSet()

            ' Llenar el DataSet
            adapter.Fill(dataSet, "ARTICULOS")

            ' Enlazar el DataSet a un control, por ejemplo, un DataGridView
            dgArticulos.DataSource = dataSet.Tables("ARTICULOS")

            'Se cierra la conexión
            DatabaseConnection.Instance.Close()
            lblNumRegistrosArticulos.Text = dgArticulos.RowCount & " registros"
        Catch ex As Exception
            ' Manejo de errores
            miClsFuncionalidades.InsertarenLog(Me.Name, "CargarDatosMarca", ex.Message.ToString)
            MessageBox.Show("Error al cargar datos: " & ex.Message)
        Finally
            If DatabaseConnection.Instance.Connection.State = ConnectionState.Open Then
                DatabaseConnection.Instance.Close()
            End If
        End Try

    End Sub

    ''' <summary>
    ''' Carga articulos con la clave del combo de familias que se pasa
    ''' </summary>
    ''' <param name="mClave">codigo familia</param>
    ''' <remarks></remarks>
    Private Sub CargarDatosFamilia(ByVal mClave As Integer)
        Try
            Dim query As String = ""
            Dim flagAlta As String = ""

            'Si el botón es ver bajas, es que estamos en las altas, se muestran las altas
            If btnVerAlta.Text = "Baja" Then
                flagAlta = " AND ARTICULOS.BAJA = No"
            Else
                flagAlta = " AND ARTICULOS.BAJA = Yes"
            End If

            ' Abrir la conexión global
            DatabaseConnection.Instance.Open()
            If mClave = 0 Then
                ' todo
                query = "SELECT ARTICULOS.CodigoArticulo, ARTICULOS.Referencia, ARTICULOS.CodigoBarras, ARTICULOS.Nombre, ARTICULOS.ExistMin," & _
                                                  " ARTICULOS.ExistMax, ARTICULOS.ExistActuales, ARTICULOS.PrecioCompra, ARTICULOS.PrecioVenta, FAMILIAS.Nombre as Familia," & _
                                                  " GRUPOS.Nombre as Grupo, IIF(Familias.Servicio = Yes, 'SI', 'NO') as Servicio" & _
                                                  " FROM (FAMILIAS INNER JOIN ARTICULOS ON FAMILIAS.CodigoFamilia = ARTICULOS.CodigoFamilia)" & _
                                                  " INNER JOIN GRUPOS ON (GRUPOS.CodigoGrupo = ARTICULOS.CodigoGrupo)" & _
                                                  " WHERE 1 = 1" & flagAlta
            Else
                'servicios
                query = "SELECT ARTICULOS.CodigoArticulo, ARTICULOS.Referencia, ARTICULOS.CodigoBarras, ARTICULOS.Nombre, ARTICULOS.ExistMin," & _
                                                  " ARTICULOS.ExistMax, ARTICULOS.ExistActuales, ARTICULOS.PrecioCompra, ARTICULOS.PrecioVenta, FAMILIAS.Nombre as Familia," & _
                                                  " GRUPOS.Nombre as Grupo, IIF(Familias.Servicio = Yes, 'SI', 'NO') as Servicio" & _
                                                  " FROM (FAMILIAS INNER JOIN ARTICULOS ON FAMILIAS.CodigoFamilia = ARTICULOS.CodigoFamilia)" & _
                                                  " INNER JOIN GRUPOS ON (GRUPOS.CodigoGrupo = ARTICULOS.CodigoGrupo)" & _
                                                  " WHERE ARTICULOS.CodigoFamilia = " & mClave & flagAlta
            End If
            ' Crear el DataAdapter
            Dim adapter As New OleDbDataAdapter(query, DatabaseConnection.Instance.Connection)

            ' Crear el DataSet
            Dim dataSet As New DataSet()

            ' Llenar el DataSet
            adapter.Fill(dataSet, "ARTICULOS")

            ' Enlazar el DataSet a un control, por ejemplo, un DataGridView
            dgArticulos.DataSource = dataSet.Tables("ARTICULOS")

            'Se cierra la conexión
            DatabaseConnection.Instance.Close()
            lblNumRegistrosArticulos.Text = dgArticulos.RowCount & " registros"
        Catch ex As Exception
            ' Manejo de errores
            miClsFuncionalidades.InsertarenLog(Me.Name, "CargarDatosFamilia", ex.Message.ToString)
            MessageBox.Show("Error al cargar datos: " & ex.Message)
        Finally
            If DatabaseConnection.Instance.Connection.State = ConnectionState.Open Then
                DatabaseConnection.Instance.Close()
            End If
        End Try

    End Sub

    ''' <summary>
    ''' Carga articulos con la clave del combo de grupos que se pasa
    ''' </summary>
    ''' <param name="mClave">codigo grupo</param>
    ''' <remarks></remarks>
    Private Sub CargarDatosGrupo(ByVal mClave As Integer)
        Try
            Dim query As String = ""
            Dim flagAlta As String = ""

            'Si el botón es ver bajas, es que estamos en las altas, se muestran las altas
            If btnVerAlta.Text = "Baja" Then
                flagAlta = " AND ARTICULOS.BAJA = No"
            Else
                flagAlta = " AND ARTICULOS.BAJA = Yes"
            End If

            ' Abrir la conexión global
            DatabaseConnection.Instance.Open()
            If mClave = 0 Then
                'Si no hay seleccionado grupo se pregunta por la familia
                If cboFamilia.SelectedIndex > 0 Then
                    'Si hay familia seleccionada se filtra por esa familia
                    query = "SELECT ARTICULOS.CodigoArticulo, ARTICULOS.Referencia, ARTICULOS.CodigoBarras, ARTICULOS.Nombre, ARTICULOS.ExistMin," & _
                                                      " ARTICULOS.ExistMax, ARTICULOS.ExistActuales, ARTICULOS.PrecioCompra, ARTICULOS.PrecioVenta, FAMILIAS.Nombre as Familia," & _
                                                      " GRUPOS.Nombre as Grupo, IIF(Familias.Servicio = Yes, 'SI', 'NO') as Servicio" & _
                                                      " FROM (FAMILIAS INNER JOIN ARTICULOS ON FAMILIAS.CodigoFamilia = ARTICULOS.CodigoFamilia)" & _
                                                      " INNER JOIN GRUPOS ON (GRUPOS.CodigoGrupo = ARTICULOS.CodigoGrupo)" & _
                                                      " WHERE ARTICULOS.CodigoFamilia = " & cboFamilia.SelectedValue & flagAlta
                Else
                    ' todo
                    query = "SELECT ARTICULOS.CodigoArticulo, ARTICULOS.Referencia, ARTICULOS.CodigoBarras, ARTICULOS.Nombre, ARTICULOS.ExistMin," & _
                                                      " ARTICULOS.ExistMax, ARTICULOS.ExistActuales, ARTICULOS.PrecioCompra, ARTICULOS.PrecioVenta, FAMILIAS.Nombre as Familia," & _
                                                      " GRUPOS.Nombre as Grupo, IIF(Familias.Servicio = Yes, 'SI', 'NO') as Servicio" & _
                                                      " FROM (FAMILIAS INNER JOIN ARTICULOS ON FAMILIAS.CodigoFamilia = ARTICULOS.CodigoFamilia)" & _
                                                      " INNER JOIN GRUPOS ON (GRUPOS.CodigoGrupo = ARTICULOS.CodigoGrupo)" & _
                                                      " WHERE 1 = 1" & flagAlta
                End If
            Else
                'servicios
                query = "SELECT ARTICULOS.CodigoArticulo, ARTICULOS.Referencia, ARTICULOS.CodigoBarras, ARTICULOS.Nombre, ARTICULOS.ExistMin," & _
                                                  " ARTICULOS.ExistMax, ARTICULOS.ExistActuales, ARTICULOS.PrecioCompra, ARTICULOS.PrecioVenta, FAMILIAS.Nombre as Familia," & _
                                                  " GRUPOS.Nombre as Grupo, IIF(Familias.Servicio = Yes, 'SI', 'NO') as Servicio" & _
                                                  " FROM (FAMILIAS INNER JOIN ARTICULOS ON FAMILIAS.CodigoFamilia = ARTICULOS.CodigoFamilia)" & _
                                                  " INNER JOIN GRUPOS ON (GRUPOS.CodigoGrupo = ARTICULOS.CodigoGrupo)" & _
                                                  " WHERE GRUPOS.CodigoGrupo = " & mClave & flagAlta
            End If
            ' Crear el DataAdapter
            Dim adapter As New OleDbDataAdapter(query, DatabaseConnection.Instance.Connection)

            ' Crear el DataSet
            Dim dataSet As New DataSet()

            ' Llenar el DataSet
            adapter.Fill(dataSet, "ARTICULOS")

            ' Enlazar el DataSet a un control, por ejemplo, un DataGridView
            dgArticulos.DataSource = dataSet.Tables("ARTICULOS")

            'Se cierra la conexión
            DatabaseConnection.Instance.Close()
            lblNumRegistrosArticulos.Text = dgArticulos.RowCount & " registros"
        Catch ex As Exception
            ' Manejo de errores            
            miClsFuncionalidades.InsertarenLog(Me.Name, "CargarDatosGrupo", ex.Message.ToString)
            MessageBox.Show("Error al cargar datos: " & ex.Message)
        Finally
            If DatabaseConnection.Instance.Connection.State = ConnectionState.Open Then
                DatabaseConnection.Instance.Close()
            End If
        End Try

    End Sub

    ''' <summary>
    ''' Carga artículos con la clave del combo de subgrupos que se pasa
    ''' </summary>
    ''' <param name="mClave">codigo subgrupo</param>
    ''' <remarks></remarks>
    Private Sub CargarDatosSubgrupo(ByVal mClave As Integer)
        Try
            Dim query As String = ""
            Dim flagAlta As String = ""

            'Si el botón es ver bajas, es que estamos en las altas, se muestran las altas
            If btnVerAlta.Text = "Baja" Then
                flagAlta = " AND ARTICULOS.BAJA = No"
            Else
                flagAlta = " AND ARTICULOS.BAJA = Yes"
            End If

            ' Abrir la conexión global
            DatabaseConnection.Instance.Open()
            If mClave = 0 Then
                ' todo
                query = "SELECT ARTICULOS.CodigoArticulo, ARTICULOS.Referencia, ARTICULOS.CodigoBarras, ARTICULOS.Nombre, ARTICULOS.ExistMin," & _
                                                  " ARTICULOS.ExistMax, ARTICULOS.ExistActuales, ARTICULOS.PrecioCompra, ARTICULOS.PrecioVenta, FAMILIAS.Nombre as Familia," & _
                                                  " GRUPOS.Nombre as Grupo, IIF(Familias.Servicio = Yes, 'SI', 'NO') as Servicio" & _
                                                  " FROM (FAMILIAS INNER JOIN ARTICULOS ON FAMILIAS.CodigoFamilia = ARTICULOS.CodigoFamilia)" & _
                                                  " INNER JOIN GRUPOS ON (GRUPOS.CodigoGrupo = ARTICULOS.CodigoGrupo)" & _
                                                  " WHERE 1 = 1" & flagAlta
            Else
                'servicios
                query = "SELECT ARTICULOS.CodigoArticulo, ARTICULOS.Referencia, ARTICULOS.CodigoBarras, ARTICULOS.Nombre, ARTICULOS.ExistMin," & _
                                                  " ARTICULOS.ExistMax, ARTICULOS.ExistActuales, ARTICULOS.PrecioCompra, ARTICULOS.PrecioVenta, FAMILIAS.Nombre as Familia," & _
                                                  " GRUPOS.Nombre as Grupo, IIF(Familias.Servicio = Yes, 'SI', 'NO') as Servicio" & _
                                                  " FROM (FAMILIAS INNER JOIN ARTICULOS ON FAMILIAS.CodigoFamilia = ARTICULOS.CodigoFamilia)" & _
                                                  " INNER JOIN GRUPOS ON (GRUPOS.CodigoGrupo = ARTICULOS.CodigoGrupo)" & _
                                                  " WHERE SUBGRUPOS.CodigoSubgrupo = " & mClave & flagAlta
            End If
            ' Crear el DataAdapter
            Dim adapter As New OleDbDataAdapter(query, DatabaseConnection.Instance.Connection)

            ' Crear el DataSet
            Dim dataSet As New DataSet()

            ' Llenar el DataSet
            adapter.Fill(dataSet, "ARTICULOS")

            ' Enlazar el DataSet a un control, por ejemplo, un DataGridView
            dgArticulos.DataSource = dataSet.Tables("ARTICULOS")

            'Se cierra la conexión
            DatabaseConnection.Instance.Close()
            lblNumRegistrosArticulos.Text = dgArticulos.RowCount & " registros"
        Catch ex As Exception
            ' Manejo de errores
            miClsFuncionalidades.InsertarenLog(Me.Name, "CargarDatosSubgrupo", ex.Message.ToString)
            MessageBox.Show("Error al cargar datos: " & ex.Message)
        Finally
            If DatabaseConnection.Instance.Connection.State = ConnectionState.Open Then
                DatabaseConnection.Instance.Close()
            End If
        End Try

    End Sub

    ''' <summary>
    ''' Carga artículos con el literal del nombre tecleado
    ''' </summary>
    ''' <param name="mNombre">nombre del artículo</param>
    Private Sub CargarArticulosPorNombre(ByVal mNombre As String)
        Try
            Dim query As String = ""
            Dim whereClave As String = ""
            Dim flagAlta As String = ""
            Dim whereNombre As String = ""


            'Si el botón es ver bajas, es que estamos en las altas, se muestran las altas
            If btnVerAlta.Text = "Baja" Then
                flagAlta = " AND ARTICULOS.BAJA = No"
            Else
                flagAlta = " AND ARTICULOS.BAJA = Yes"
            End If
            
            If cboSubgrupo.SelectedIndex < 1 Then
                If cboGrupo.SelectedIndex < 1 Then
                    If cboFamilia.SelectedIndex < 1 Then
                        If cboMarca.SelectedIndex < 1 Then
                            whereClave = ""
                        Else
                            whereClave = " AND ARTICULOS.CODIGOMARCA = " & cboMarca.SelectedValue
                        End If
                    Else
                        whereClave = " AND FAMILIAS.CODIGOFAMILIA = " & cboFamilia.SelectedValue
                    End If
                Else
                    whereClave = " AND FAMILIAS.CODIGOGRUPO = " & cboGrupo.SelectedValue
                End If
            Else
                whereClave = " AND FAMLIAS.CODIGOSUBGRUPO = " & cboSubgrupo.SelectedValue
            End If
            'AND ARTICULOS.Nombre LIKE '*Pu*'
            whereNombre = " AND ARTICULOS.Nombre LIKE '%" & mNombre & "%'"

            ' Abrir la conexión global
            DatabaseConnection.Instance.Open()

            ' todo
            query = "SELECT ARTICULOS.CodigoArticulo, ARTICULOS.Referencia, ARTICULOS.CodigoBarras, ARTICULOS.Nombre, ARTICULOS.ExistMin," & _
                    " ARTICULOS.ExistMax, ARTICULOS.ExistActuales, ARTICULOS.PrecioCompra, ARTICULOS.PrecioVenta, FAMILIAS.Nombre as Familia," & _
                    " GRUPOS.Nombre as Grupo, SUBGRUPOS.Nombre as Subgrupo, IIF(Familias.Servicio = Yes, 'SI', 'NO') as Servicio" & _
                    " FROM SUBGRUPOS RIGHT JOIN (GRUPOS RIGHT JOIN (FAMILIAS RIGHT JOIN ARTICULOS ON FAMILIAS.CodigoFamilia = ARTICULOS.CodigoFamilia)" & _
                    " ON GRUPOS.CodigoGrupo = ARTICULOS.CodigoGrupo) ON SUBGRUPOS.CodigoSubGrupo = ARTICULOS.CodigoSubGrupo" & _
                    " WHERE 1 = 1" & whereClave & flagAlta & whereNombre

            ' Crear el DataAdapter
            Dim adapter As New OleDbDataAdapter(query, DatabaseConnection.Instance.Connection)

            ' Crear el DataSet
            Dim dataSet As New DataSet()

            ' Llenar el DataSet
            adapter.Fill(dataSet, "ARTICULOS")

            ' Enlazar el DataSet a un control, por ejemplo, un DataGridView
            dgArticulos.DataSource = dataSet.Tables("ARTICULOS")

            'Se cierra la conexión
            DatabaseConnection.Instance.Close()
            lblNumRegistrosArticulos.Text = dgArticulos.RowCount & " registros"
        Catch ex As Exception
            ' Manejo de errores
            miClsFuncionalidades.InsertarenLog(Me.Name, "CargarArticulosPorNombre", ex.Message.ToString)
            MessageBox.Show("Error al cargar datos: " & ex.Message)
        Finally
            If DatabaseConnection.Instance.Connection.State = ConnectionState.Open Then
                DatabaseConnection.Instance.Close()
            End If
        End Try

    End Sub

    ''' <summary>
    ''' Carga artículos con el literal del codigo de barras tecleado
    ''' </summary>
    ''' <param name="mCodigo">codigo de barras del artículo</param>
    Private Sub CargarArticulosPorCodigoBarras(ByVal mCodigo As String)
        Try
            Dim query As String = ""
            Dim whereClave As String = ""
            Dim flagAlta As String = ""
            Dim whereNombre As String = ""

            'Si el botón es ver bajas, es que estamos en las altas, se muestran las altas
            If btnVerAlta.Text = "Baja" Then
                flagAlta = " AND ARTICULOS.BAJA = No"
            Else
                flagAlta = " AND ARTICULOS.BAJA = Yes"
            End If

            If cboSubgrupo.SelectedIndex < 1 Then
                If cboGrupo.SelectedIndex < 1 Then
                    If cboFamilia.SelectedIndex < 1 Then
                        If cboMarca.SelectedIndex < 1 Then
                            whereClave = ""
                        Else
                            whereClave = " AND ARTICULOS.CODIGOMARCA = " & cboMarca.SelectedValue
                        End If
                    Else
                        whereClave = " AND FAMILIAS.CODIGOFAMILIA = " & cboFamilia.SelectedValue
                    End If
                Else
                    whereClave = " AND FAMILIAS.CODIGOGRUPO = " & cboGrupo.SelectedValue
                End If
            Else
                whereClave = " AND FAMLIAS.CODIGOSUBGRUPO = " & cboSubgrupo.SelectedValue
            End If
            'AND ARTICULOS.Nombre LIKE '*Pu*'
            whereNombre = " AND ARTICULOS.CodigoBarras LIKE '%" & mCodigo & "%'"
            ' Abrir la conexión global
            DatabaseConnection.Instance.Open()

            ' todo
            query = "SELECT ARTICULOS.CodigoArticulo, ARTICULOS.Referencia, ARTICULOS.CodigoBarras, ARTICULOS.Nombre, ARTICULOS.ExistMin," & _
                    " ARTICULOS.ExistMax, ARTICULOS.ExistActuales, ARTICULOS.PrecioCompra, ARTICULOS.PrecioVenta, FAMILIAS.Nombre as Familia," & _
                    " GRUPOS.Nombre as Grupo, SUBGRUPOS.Nombre as Subgrupo, IIF(Familias.Servicio = Yes, 'SI', 'NO') as Servicio" & _
                    " FROM SUBGRUPOS RIGHT JOIN (GRUPOS RIGHT JOIN (FAMILIAS RIGHT JOIN ARTICULOS ON FAMILIAS.CodigoFamilia = ARTICULOS.CodigoFamilia)" & _
                    " ON GRUPOS.CodigoGrupo = ARTICULOS.CodigoGrupo) ON SUBGRUPOS.CodigoSubGrupo = ARTICULOS.CodigoSubGrupo" & _
                    " WHERE 1 = 1" & whereClave & flagAlta & whereNombre

            ' Crear el DataAdapter
            Dim adapter As New OleDbDataAdapter(query, DatabaseConnection.Instance.Connection)

            ' Crear el DataSet
            Dim dataSet As New DataSet()

            ' Llenar el DataSet
            adapter.Fill(dataSet, "ARTICULOS")

            ' Enlazar el DataSet a un control, por ejemplo, un DataGridView
            dgArticulos.DataSource = dataSet.Tables("ARTICULOS")

            'Se cierra la conexión
            DatabaseConnection.Instance.Close()
            lblNumRegistrosArticulos.Text = dgArticulos.RowCount & " registros"
        Catch ex As Exception
            ' Manejo de errores
            miClsFuncionalidades.InsertarenLog(Me.Name, "CargarArticulosPorCodigoBarras", ex.Message.ToString)
            MessageBox.Show("Error al cargar datos: " & ex.Message)
        Finally
            If DatabaseConnection.Instance.Connection.State = ConnectionState.Open Then
                DatabaseConnection.Instance.Close()
            End If
        End Try

    End Sub

    Private Sub cargarTodosCombos()
        Try
            Dim mMarca As Integer
            Dim Clave As Integer

            If cboMarca.SelectedIndex > 0 Then
                mMarca = cboMarca.SelectedValue
            Else
                mMarca = 0
            End If

            If cboSubgrupo.SelectedIndex < 1 Then
                If cboGrupo.SelectedIndex < 1 Then
                    If cboFamilia.SelectedIndex < 1 Then
                        If cboMarca.SelectedIndex < 1 Then
                            Clave = 0
                        Else
                            Clave = cboMarca.SelectedValue
                        End If
                        CargarDatosMarca(Clave)
                    Else
                        Clave = cboFamilia.SelectedValue
                        CargarDatosFamilia(Clave)
                    End If
                Else
                    Clave = cboGrupo.SelectedValue
                    CargarDatosGrupo(Clave)
                End If
            Else
                Clave = cboSubgrupo.SelectedValue
                CargarDatosSubgrupo(Clave)
            End If
        Catch ex As Exception
            ' Manejo de errores
            miClsFuncionalidades.InsertarenLog(Me.Name, "cargarTodosCombos", ex.Message.ToString)
            MessageBox.Show("Error al cargar datos: " & ex.Message)
        End Try
    End Sub

#End Region


End Class