Imports System.Data.OleDb

Public Class frmAniadirArticulo

    Private mCodigoArticulo As Integer
    Private mServicio As Integer
    
    Dim mSalirSinModificar As Boolean = True
    Dim miClsFuncionalidades As New clsFuncionalidades

    ''' <summary>
    ''' Propiedad para saber si se sale del formulario sin actualizar
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property SalirSinModificar
        Get
            Return mSalirSinModificar
        End Get
        Set(value)
            mSalirSinModificar = value
        End Set
    End Property

#Region "constructor"

    Public Sub New(ByVal pCodigoArticulo As String)
        ' Este llamado es necesario para inicializar los componentes del formulario
        InitializeComponent()

        ' Guardar el parámetro en una variable
        mCodigoArticulo = pCodigoArticulo
    End Sub

#End Region

    Private Sub frmAniadirArticulo_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        'esta variable se utilizará para saber si se sale sin cambiar o no, para si se sale sin cambiar
        ' no llamar a refrescar en el formulario frmArticulos
        mSalirSinModificar = True

        Dim query As String = ""
        Dim where As String = ""
        Dim mCodigoMarca As Integer
        Dim mCodigoFamilia As Integer
        Dim mCodigoGrupo As Integer
        Dim mCodigoSubgrupo As Integer
        Dim mComentarios As String

        cmdNuevaFamilia.Enabled = False
        cmdNuevoGrupo.Enabled = False
        cmdNuevoSubgrupo.Enabled = False
        Try
            CargarComboIVACompra()
            CargarComboIVAVenta()
            If cboIVACompra.Items.Count > 1 Then
                cboIVACompra.SelectedIndex = 1
            End If

            If cboIVAVenta.Items.Count > 1 Then
                cboIVAVenta.SelectedIndex = 1
            End If

            If mCodigoArticulo <> 0 Then

                where = " WHERE ARTICULOS.CodigoArticulo =" & mCodigoArticulo
                Dim items As New List(Of clsComboboxItem)


                'Sin seleccion
                query = "SELECT ARTICULOS.Referencia, ARTICULOS.CodigoBarras, ARTICULOS.Nombre, " & _
                        " ARTICULOS.PrecioCompra, ARTICULOS.PCconIVA," & _
                        " ARTICULOS.ExistMin, ARTICULOS.ExistMax, ARTICULOS.ExistActuales, " & _
                        " ARTICULOS.PrecioVenta, ARTICULOS.PVP, ARTICULOS.CodigoMarca, ARTICULOS.CodigoFamilia, " & _
                        " ARTICULOS.CodigoGrupo, ARTICULOS.CodigoSubGrupo, ARTICULOS.Compuesto, " & _
                        " ARTICULOS.ArticuloUsoInterno, ARTICULOS.Comentarios, " & _
                        " Familias.Servicio FROM ARTICULOS" & _
                        " INNER JOIN FAMILIAS ON FAMILIAS.CODIGOFAMILIA = ARTICULOS.CODIGOFAMILIA " & where & ";"

                'Para la pestaña de promoción invisible por ahora, en caso de habilitarla habrá que incluirlo en la query
                '" , ARTICULOS.CodigoPromocion, ARTICULOS.Descuento," & _
                '" ARTICULOS.PrecioUnidad, ARTICULOS.IvaVentaPromocion, ARTICULOS.PrecioTotalpromocion, " & _


                DatabaseConnection.Instance.Open()

                Using command As New OleDbCommand(query, DatabaseConnection.Instance.Connection)

                    ' Ejecutar la consulta
                    Using reader As OleDbDataReader = command.ExecuteReader()

                        ' Leer los datos y agregarlos al ComboBox
                        If reader.Read() Then
                            'agregar los elementos a los controles
                            txtReferencia.Text = reader("Referencia").ToString
                            txtCodigoBarras.Text = reader("CodigoBarras").ToString
                            txtDescripcion.Text = reader("Nombre").ToString
                            txtPrecio.Text = reader("PrecioCompra").ToString
                            txtPrecioMasIVA.Text = reader("PCconIVA").ToString
                            txtExistenciasMaximas.Text = reader("ExistMax").ToString
                            txtExistenciasMinimas.Text = reader("ExistMin").ToString
                            txtExistenciasActuales.Text = reader("ExistActuales").ToString
                            txtPrecioVenta.Text = reader("PrecioVenta").ToString
                            txtPVP1.Text = reader("PVP").ToString

                            If reader("CodigoMarca").ToString = "" Then
                                mCodigoMarca = 0
                            Else
                                mCodigoMarca = reader("CodigoMarca")
                            End If
                            If reader("CodigoFamilia").ToString = "" Then
                                mCodigoFamilia = 0
                            Else
                                mCodigoFamilia = reader("CodigoFamilia")
                            End If
                            If reader("CodigoGrupo").ToString = "" Then
                                mCodigoGrupo = 0
                            Else
                                mCodigoGrupo = reader("CodigoGrupo")
                            End If
                            If reader("CodigoSubgrupo").ToString = "" Then
                                mCodigoSubgrupo = 0
                            Else
                                mCodigoSubgrupo = reader("CodigoSubgrupo")
                            End If

                            chkArticuloCombinado.Checked = IIf(reader("Compuesto").ToString = "Yes", True, False)
                            chkUsoInterno.Checked = IIf(reader("ArticuloUsoInterno").ToString = "Yes", True, False)

                            mComentarios = reader("Comentarios").ToString
                            txtComentarios.Text = mComentarios

                            mServicio = IIf(reader("Servicio") = False, 0, 1)

                            'para la pestaña de promoción (invisible por ahora) en caso de habilitar la pestaña
                            'habra que descomentar estas líneas y las de la select de arriba
                            'txtDescuento.Text = reader("DescuentoProveedor").ToString
                            'txtPVUnidad.Text = ""
                            'txtPrecioTotal.Text = ""
                            'CargarComboPromocion()

                            cargarComboMarcas(mCodigoMarca)
                            cargarComboFamilias(mCodigoFamilia, mCodigoMarca)
                            CargarComboGrupos(mCodigoFamilia, mCodigoGrupo)
                            CargarComboSubgrupos(mCodigoGrupo, mCodigoSubgrupo)

                        End If
                    End Using
                End Using

            Else
                'Si viene de nuevo se cargan todas las marcas
                cargarComboMarcas(0)
                CargarComboIVACompra()
                CargarComboIVAVenta()

                If cboIVACompra.Items.Count > 1 Then
                    cboIVACompra.SelectedIndex = 1
                End If

                If cboIVAVenta.Items.Count > 1 Then
                    cboIVAVenta.SelectedIndex = 1
                End If

            End If
        Catch ex As Exception
            miClsFuncionalidades.InsertarenLog(Me.Name, "form_load", ex.Message.ToString)
            MessageBox.Show("Error al cargar datos del artículo: " & ex.Message)
        Finally
            If DatabaseConnection.Instance.Connection.State = ConnectionState.Open Then
                DatabaseConnection.Instance.Close()
            End If
        End Try

    End Sub

#Region "Eventos"

#Region "Botones"

    Private Sub cmdSalir_Click(sender As System.Object, e As System.EventArgs) Handles cmdSalir.Click

        Me.Close()

    End Sub

    Private Sub cmdNuevaFamilia_Click(sender As System.Object, e As System.EventArgs) Handles cmdNuevaFamilia.Click
        Try
            Dim mCodigoFamilia As Integer = 0
            If cboFamilia.SelectedIndex > 0 Then
                mCodigoFamilia = cboFamilia.SelectedValue
            End If

            If cboMarca.SelectedIndex < -1 Then
                MessageBox.Show("La marca es obligatoria", "Nueva Familia", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                Dim mfrmFamilia As New frmFamilia(0, cboMarca.SelectedValue, mCodigoFamilia)
                mfrmFamilia.VengodeArticulos = True
                mfrmFamilia.BringToFront()
                mfrmFamilia.ShowDialog()
                cargarComboFamilias(0, cboMarca.SelectedValue)
            End If
        Catch ex As Exception
            miClsFuncionalidades.InsertarenLog(Me.Name, "cmdNuevaFamilia_Click", ex.Message.ToString)
            MessageBox.Show("Error al cargar datos del artículo: " & ex.Message)
        End Try

    End Sub

    Private Sub cmdNuevo_Click(sender As System.Object, e As System.EventArgs) Handles cmdNuevo.Click

        'Limpiamos los controles
        mCodigoArticulo = 0
        mServicio = 0
        
        txtReferencia.Text = ""
        txtCodigoBarras.Text = ""
        txtDescripcion.Text = ""

        txtPrecio.Text = ""
        txtPrecioMasIVA.Text = ""
        txtExistenciasMinimas.Text = ""
        txtExistenciasMaximas.Text = ""
        txtExistenciasActuales.Text = ""
        txtPrecioVenta.Text = ""
        txtPVP1.Text = ""

        chkArticuloCombinado.Checked = False
        chkUsoInterno.Checked = False
        txtComentarios.Text = ""

        ' limpiamos los combos dependientes
        cboFamilia.DataSource = Nothing
        cboGrupo.DataSource = Nothing
        cboSubgrupo.DataSource = Nothing
        cboProveedor.DataSource = Nothing
        cboIVACompra.DataSource = Nothing

        cboProveedor.Items.Clear()
        cboIVACompra.Items.Clear()
        cboFamilia.Items.Clear()
        cboGrupo.Items.Clear()
        cboSubgrupo.Items.Clear()

        cmdNuevaFamilia.Enabled = True
        cmdNuevoGrupo.Enabled = False
        cmdNuevoSubgrupo.Enabled = False

        cboMarca.SelectedIndex = 0

        'Esto es del tab promociones invisible por ahora
        txtDescuento.Text = ""
        txtPVUnidad.Text = ""
        txtPrecioTotal.Text = ""
        cboPromocion.DataSource = Nothing
        cboPromocion.Items.Clear()
        cboIVAVenta.DataSource = Nothing
        cboIVAVenta.Items.Clear()

    End Sub

    Private Sub cmdNuevoGrupo_Click(sender As System.Object, e As System.EventArgs) Handles cmdNuevoGrupo.Click
        Try
            Dim mFamilia As Integer = 0
            Dim mGrupo As Integer = 0
            Dim msubgrupo As Integer = 0

            If cboFamilia.SelectedIndex > 0 Then
                mFamilia = cboFamilia.SelectedValue
            End If
            If cboGrupo.SelectedIndex > 0 Then
                mGrupo = cboGrupo.SelectedValue
            End If

            If mFamilia <> 0 Then
                Dim mfrmGrupo As New frmGrupo(mFamilia, mGrupo, msubgrupo, True)
                mfrmGrupo.BringToFront()
                mfrmGrupo.ShowDialog()
                CargarComboGrupos(mFamilia, mGrupo)
            Else
                MessageBox.Show("No se puede dar de alta un grupo sin que pertenezca a una familia", "Nuevo grupo", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            miClsFuncionalidades.InsertarenLog(Me.Name, "cmdNuevoGrupo_Click", ex.Message.ToString)
            MessageBox.Show("Error al cargar datos del artículo: " & ex.Message)
        End Try

    End Sub

    Private Sub cmdNuevoSubgrupo_Click(sender As System.Object, e As System.EventArgs) Handles cmdNuevoSubgrupo.Click

        Dim mFamilia As Integer = 0
        Dim mGrupo As Integer = 0
        Dim mSubgrupo As Integer = 0

        If cboFamilia.SelectedIndex > 0 Then
            mFamilia = cboFamilia.SelectedValue
        End If
        If cboGrupo.SelectedIndex > 0 Then
            mGrupo = cboGrupo.SelectedValue
        End If
        If cboSubgrupo.SelectedIndex > 0 Then
            mSubgrupo = cboSubgrupo.SelectedValue
        End If

        If mgrupo <> 0 Then
            Dim mfrmGrupo As New frmGrupo(mfamilia, mgrupo, msubgrupo, False)
            mfrmGrupo.BringToFront()
            mfrmGrupo.ShowDialog()
            CargarComboSubgrupos(mGrupo, mSubgrupo)
        Else
            MessageBox.Show("No se puede dar de alta un subgrupo sin que pertenezca a un grupo", "Nuevo subgrupo", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

    End Sub

    Private Sub cmdGuardar_Click(sender As System.Object, e As System.EventArgs) Handles cmdGuardar.Click

        Dim mfamilia As Integer
        Dim mgrupo As Integer
        Dim msubgrupo As Integer
        Dim mvalido As Boolean = True
        Dim mMensaje As String = ""

        If cboFamilia.SelectedIndex < 1 Then
            mfamilia = 0
        Else
            mfamilia = cboFamilia.SelectedValue
        End If
        If cboGrupo.SelectedIndex < 1 Then
            mgrupo = 0
        Else
            mgrupo = cboGrupo.SelectedValue
        End If

        If cboSubgrupo.SelectedIndex < 1 Then
            msubgrupo = 0
        Else
            msubgrupo = cboSubgrupo.SelectedValue
        End If

        'validar campos obligatorios
        If mfamilia = 0 Or mgrupo = 0 Then
            mvalido = False
            mMensaje = "La familia y el grupo son obligatorios"
        End If
        If txtReferencia.Text = "" Then
            mvalido = False
            mMensaje = "No se puede dar de alta un artículo sin código de referencia"
        End If
        If txtCodigoBarras.Text = "" Then
            mvalido = False
            mMensaje = "No se puede dar de alta un artículo sin el código de barras"
        End If
        If txtDescripcion.Text = "" Then
            mvalido = False
            mMensaje = "La descripción del artículo es obligatoria"
        End If
        If cboMarca.SelectedIndex < 1 Then
            mvalido = False
            mMensaje = "La marca es obligatoria"
        End If

        'rellenar campos por defecto
        If txtExistenciasActuales.Text = "" Then
            txtExistenciasActuales.Text = "0"
        End If
        If txtExistenciasMaximas.Text = "" Then
            txtExistenciasMaximas.Text = "0"
        End If
        If txtExistenciasMinimas.Text = "" Then
            txtExistenciasMinimas.Text = "0"
        End If

        If mvalido = False Then
            MessageBox.Show(mMensaje, "Alta de artículo", MessageBoxButtons.OK)
        Else

            Dim miClsFuncionalidades As New clsFuncionalidades
            ' Recolectar datos
            Dim mNombre As String = txtDescripcion.Text
            Dim mDescripcion As String = txtDescripcion.Text
            Dim mUsoInterno As Boolean = chkUsoInterno.Checked
            Dim query As String

            If mCodigoArticulo = 0 Then
                'Consulta SQL de inserción
                query = "INSERT INTO ARTICULOS(Referencia, CodigoBarras, Nombre, CodigoProveedor, PrecioCompra," & _
                        " IVACOMPRA, PCconIVA, PrecioVenta, ExistMin, ExistMax, ExistActuales, IVAVENTA, PVP, Compuesto," & _
                        " ArticuloUsoInterno, CodigoMarca, CodigoFamilia, CodigoGrupo, CodigoSubGrupo, Comentarios, Baja" & _
                        " ) Values (@Referencia, @CodigoBarras, @Nombre, @CodigoProveedor, @PrecioCompra," & _
                        " @IVACOMPRA, @PCconIVA, @PrecioVenta, @ExistMin, @ExistMax, @ExistActuales, @IVAVENTA, @PVP, @Compuesto," & _
                        " @ArticuloUsoInterno, @CodigoMarca, @CodigoFamilia, @CodigoGrupo, @CodigoSubGrupo, @Comentarios, 0)"

                'para la pestaña invisible (por ahora) de promocion. Habría que incluirlo en la query de insert
                '" ,CodigoPromocion, Descuento, PrecioUnidad, IVAVentaPromocion, PrecioTotalPromocion)" & _

                'para la pestaña invisible (por ahora) de promocion. Habría que incluirlo en la query de insert para los values
                '" @CodigoPromocion, @Descuento, @PrecioUnidad, @IVAVentaPromocion, @PrecioTotalPromocion 0)"
            Else
                query = "UPDATE ARTICULOS SET Referencia = @Referencia, CodigoBarras = @CodigoBarras, Nombre = @Nombre," & _
                        " CodigoProveedor = @CodigoProveedor, PrecioCompra = @PrecioCompra," & _
                        " IVACOMPRA = @IVACOMPRA, PCconIVA = @PCconIVA, PrecioVenta = @PrecioVenta," & _
                        " ExistMin = @ExistMin, ExistMax = @ExistMax, ExistActuales = @ExistActuales," & _
                        " IVAVENTA = @IVAVENTA, PVP = @PVP, Compuesto = @Compuesto, ArticuloUsoInterno = @ArticuloUsoInterno," & _
                        " CodigoMarca = @CodigoMarca, CodigoFamilia = @CodigoFamilia, CodigoGrupo = @CodigoGrupo," & _
                        " CodigoSubGrupo = @CodigoSubGrupo, Comentarios = @Comentarios " & _
                        " WHERE ARTICULOS.CODIGOARTICULO = " & mCodigoArticulo

                'para la pestaña invisible (por ahora) de promocion. Habría que incluirlo en la query de update
                '" ,CodigoPromocion = @CodigoPromocion, Descuento = @Descuento, PrecioUnidad = @PrecioUnidad," & _
                '" IVAVentaPromocion = @IVAVentaPromocion, PrecioTotalPromocion = @PrecioTotalPromocion" & _
            End If

            Try
                DatabaseConnection.Instance.Open()

                Using command As New OleDbCommand(query, DatabaseConnection.Instance.Connection)

                    command.Parameters.AddWithValue("@Referencia", txtReferencia.Text)
                    command.Parameters.AddWithValue("@CodigoBarras", txtCodigoBarras.Text)
                    command.Parameters.AddWithValue("@Nombre", txtDescripcion.Text)
                    command.Parameters.AddWithValue("@CodigoProveedor", IIf(cboProveedor.SelectedIndex < 1, 0, cboProveedor.SelectedValue))
                    command.Parameters.AddWithValue("@PrecioCompra", IIf(txtPrecio.Text = "", 0, txtPrecio.Text))
                    command.Parameters.AddWithValue("@IVACOMPRA", IIf(cboIVACompra.SelectedValue Is Nothing, "0", cboIVACompra.SelectedValue))

                    command.Parameters.AddWithValue("@PCconIVA", IIf(txtPrecioMasIVA.Text = "", 0, txtPrecioMasIVA.Text))
                    command.Parameters.AddWithValue("@PrecioVenta", IIf(txtPrecioVenta.Text = "", 0, txtPrecioVenta.Text))
                    command.Parameters.AddWithValue("@ExistMin", IIf(txtExistenciasMinimas.Text = "", 0, txtExistenciasMinimas.Text))
                    command.Parameters.AddWithValue("@ExistMax", IIf(txtExistenciasMaximas.Text = "", 0, txtExistenciasMaximas.Text))
                    command.Parameters.AddWithValue("@ExistActuales", IIf(txtExistenciasActuales.Text = "", 0, txtExistenciasActuales.Text))
                    command.Parameters.AddWithValue("@IVAVENTA", IIf(cboIVAVenta.SelectedValue Is Nothing, 0, cboIVAVenta.SelectedValue))
                    command.Parameters.AddWithValue("@PVP", IIf(txtPVP1.Text = "", 0, txtPVP1.Text))
                    command.Parameters.AddWithValue("@Compuesto", False)
                    command.Parameters.AddWithValue("@ArticuloUsoInterno", chkUsoInterno.Checked)
                    command.Parameters.AddWithValue("@CodigoMarca", cboMarca.SelectedValue)
                    command.Parameters.AddWithValue("@CodigoFamilia", mfamilia)
                    command.Parameters.AddWithValue("@CodigoGrupo", mgrupo)
                    command.Parameters.AddWithValue("@CodigoSubGrupo", IIf(msubgrupo = 0, 204, msubgrupo))
                    command.Parameters.AddWithValue("@Comentarios", txtComentarios.Text)

                    'para la pestaña de promoción (invisible por ahora) en caso de habilitarse habrá
                    ' que descomentar estos parametros así como los parametros en las querys de insert y update de arriba
                    'command.Parameters.AddWithValue("@CodigoPromocion", IIf(cboPromocion.SelectedIndex < 1, 0, cboPromocion.SelectedValue))
                    'command.Parameters.AddWithValue("@Descuento", txtDescuento.Text)
                    'command.Parameters.AddWithValue("@PrecioUnidad", txtPVUnidad.Text)
                    'command.Parameters.AddWithValue("@IVAVentaPromocion", cboIVAVentaPromocion.SelectedValue)
                    'command.Parameters.AddWithValue("@PrecioTotalPromocion", txtPrecioTotal.Text)


                    Dim rowsAffected As Integer = command.ExecuteNonQuery()
                    If rowsAffected > 0 Then
                        MessageBox.Show("Datos guardados correctamente.")
                    Else
                        MessageBox.Show("No se ha insertado ninguna fila.")
                    End If

                End Using

            Catch ex As Exception
                miClsFuncionalidades.InsertarenLog(Me.Name, "cmdGuardar_Click", ex.Message.ToString)
                mSalirSinModificar = True
                MessageBox.Show("Error al guardar los datos: " & ex.Message)
            Finally
                If DatabaseConnection.Instance.Connection.State = ConnectionState.Open Then
                    DatabaseConnection.Instance.Close()
                End If
            End Try

            'Con esto nos aseguramos que salimos sin cambiar nada, para que no se refresque la
            'pantalla que llama (frmArticulos)
            mSalirSinModificar = False

        End If

    End Sub

#End Region

#Region "Combos"

    Private Sub cboMarca_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cboMarca.SelectedIndexChanged
        Try
            If cboMarca.SelectedIndex > 0 Then
                cboFamilia.DataSource = Nothing
                cboFamilia.Items.Clear()

                cargarComboFamilias(0, cboMarca.SelectedValue)
                cmdNuevaFamilia.Enabled = True

            Else
                cboFamilia.DataSource = Nothing
                cboFamilia.Items.Clear()
                cmdNuevaFamilia.Enabled = False
            End If
        Catch ex As Exception
            miClsFuncionalidades.InsertarenLog(Me.Name, "cboMarca_SelectedIndexChanged", ex.Message.ToString)
            MessageBox.Show("Error al cambiar la marca de la lista desplegable: " & ex.Message)
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

            If cboFamilia.SelectedIndex = 0 Then
                cmdNuevoGrupo.Enabled = False
                cmdNuevoSubgrupo.Enabled = False
            ElseIf cboFamilia.SelectedIndex = -1 Then
                cmdNuevoGrupo.Enabled = False
                cmdNuevoSubgrupo.Enabled = False
            Else
                clave = cboFamilia.SelectedValue
                CargarComboGrupos(clave, 0)
                cmdNuevoGrupo.Enabled = True
            End If

        Catch ex As Exception
            miClsFuncionalidades.InsertarenLog(Me.Name, "cboFamilia_SelectedIndexChanged", ex.Message.ToString)
            MessageBox.Show("Error al cargar datos: " & ex.Message)
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
                cmdNuevoSubgrupo.Enabled = False
            Else
                clave = cboGrupo.SelectedValue
                cmdNuevoSubgrupo.Enabled = True
            End If

            CargarComboSubgrupos(clave, 0)

        Catch ex As Exception
            miClsFuncionalidades.InsertarenLog(Me.Name, "cboGrupo_SelectedIndexChanged", ex.Message.ToString)
            MessageBox.Show("Error al cargar datos: " & ex.Message)
        End Try
    End Sub

#End Region

#Region "textbox"

    Private Sub txtCodigoBarras_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtCodigoBarras.KeyPress
        ' Verifica si la tecla presionada no es un dígito ni una tecla de control (como Backspace)
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True ' Ignora la tecla
        End If
    End Sub

    Private Sub txtExistenciasMinimas_keyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtExistenciasMinimas.KeyPress
        ' Verifica si la tecla presionada no es un dígito ni una tecla de control (como Backspace)
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True ' Ignora la tecla
        End If
    End Sub

    Private Sub txtExistenciasMaximas_keypress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtExistenciasMaximas.KeyPress
        ' Verifica si la tecla presionada no es un dígito ni una tecla de control (como Backspace)
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True ' Ignora la tecla
        End If
    End Sub

    Private Sub txtExistenciasActuales_keypress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtExistenciasActuales.KeyPress
        ' Verifica si la tecla presionada no es un dígito ni una tecla de control (como Backspace)
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True ' Ignora la tecla
        End If
    End Sub

    Private Sub txtPrecio_keypress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtPrecio.KeyPress
        ' Verifica si la tecla presionada no es un dígito ni una tecla de control (como Backspace)
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso e.KeyChar <> "," Then
            e.Handled = True ' Ignora la tecla
        End If
    End Sub

    Private Sub txtPrecioVenta_keypress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtPrecioVenta.KeyPress
        ' Verifica si la tecla presionada no es un dígito ni una tecla de control (como Backspace)
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso e.KeyChar <> "," Then
            e.Handled = True ' Ignora la tecla
        End If
    End Sub

    Private Sub txtPrecioMasIva_keypress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtPrecioMasIVA.KeyPress
        ' Verifica si la tecla presionada no es un dígito ni una tecla de control (como Backspace)
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso e.KeyChar <> "," Then
            e.Handled = True ' Ignora la tecla
        End If
    End Sub

    Private Sub txtPVP1_keypress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtPVP1.KeyPress
        ' Verifica si la tecla presionada no es un dígito ni una tecla de control (como Backspace)
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso e.KeyChar <> "," Then
            e.Handled = True ' Ignora la tecla
        End If
    End Sub

    Private Sub txtDescuento_keypress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtDescuento.KeyPress
        ' Verifica si la tecla presionada no es un dígito ni una tecla de control (como Backspace)
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso e.KeyChar <> "," Then
            e.Handled = True ' Ignora la tecla
        End If
    End Sub

    Private Sub txtPVUnidad_keypress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtPVUnidad.KeyPress
        ' Verifica si la tecla presionada no es un dígito ni una tecla de control (como Backspace)
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso e.KeyChar <> "," Then
            e.Handled = True ' Ignora la tecla
        End If
    End Sub

    Private Sub txtPrecioTotal_keypress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtPrecioTotal.KeyPress
        ' Verifica si la tecla presionada no es un dígito ni una tecla de control (como Backspace)
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso e.KeyChar <> "," Then
            e.Handled = True ' Ignora la tecla
        End If
    End Sub

    Private Sub txtPVP1_Leave(sender As System.Object, e As System.EventArgs) Handles txtPVP1.Leave
        Try
            Dim mPrecioSinIVA As Double
            Dim mPorcentajeIVA As Double
            If txtPVP1.Text = "0" Or txtPVP1.Text = "" Then
                txtPrecioVenta.Text = ""
            Else
                mPorcentajeIVA = CType(cboIVACompra.SelectedItem.ToString, Decimal)

                mPrecioSinIVA = txtPVP1.Text - (txtPVP1.Text * mPorcentajeIVA / 100)
                txtPrecioVenta.Text = Math.Round(mPrecioSinIVA, 2)

            End If
        Catch ex As Exception
            miClsFuncionalidades.InsertarenLog(Me.Name, "txtPVP1_Leave", ex.Message.ToString)
            MessageBox.Show("Error al calcluar el precio con el IVA: " & ex.Message)
        End Try
    End Sub

    Private Sub txtPrecioMasIVA_Leave(sender As System.Object, e As System.EventArgs) Handles txtPrecioMasIVA.Leave
        Try
            Dim mPrecioSinIVA As Double
            Dim mPorcentajeIVA As Double

            If txtPrecioMasIVA.Text = "0" Or txtPrecioMasIVA.Text = "" Then
                txtPrecio.Text = ""
            Else
                mPorcentajeIVA = CType(cboIVACompra.SelectedItem.ToString, Decimal)

                mPrecioSinIVA = txtPrecioMasIVA.Text - (txtPrecioMasIVA.Text * mPorcentajeIVA / 100)
                txtPrecio.Text = Math.Round(mPrecioSinIVA, 2)

            End If
        Catch ex As Exception
            miClsFuncionalidades.InsertarenLog(Me.Name, "txtPrecioMasIVA_Leave", ex.Message.ToString)
            MessageBox.Show("Error al calcluar el precio con el IVA: " & ex.Message)
        End Try
    End Sub

#End Region

#End Region

#Region "metodos privados"

    Private Sub cargarComboFamilias(ByVal pCodigoFamilia As Integer, ByVal pCodigoMarca As Integer)

        Dim query As String
        Dim items As New List(Of clsComboboxItem)
        Try

            ' limpiamos los combos dependientes
            cboFamilia.DataSource = Nothing
            cboGrupo.DataSource = Nothing
            cboSubgrupo.DataSource = Nothing

            cboFamilia.Items.Clear()
            cboGrupo.Items.Clear()
            cboSubgrupo.Items.Clear()

            cmdNuevoGrupo.Enabled = False
            cmdNuevoSubgrupo.Enabled = False

            'Abrimos conexión
            DatabaseConnection.Instance.Open()

            'montamos la query
            query = "SELECT Familias.CodigoFamilia as Id, FAMILIAS.Nombre as Familia FROM FAMILIAS Where 1 = 1 "

            If mServicio = 0 Then
                query = query & " And Familias.Servicio = No"
            Else
                query = query & " And Familias.Servicio = Yes"
            End If
            If pCodigoMarca <> 0 Then
                query = query & " And Familias.CodigoMarca =" & pCodigoMarca
            End If

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
                            .Descripcion = reader("Familia").ToString()}
                        items.Add(item)

                        ' Agregar elementos al ComboBox (Texto visible y valor subyacente)
                        'cboFamilia.Items.Add(reader(1))
                    End While

                    cboFamilia.DataSource = items
                    cboFamilia.DisplayMember = "Descripcion"
                    cboFamilia.ValueMember = "Id"

                End Using
            End Using

            If pCodigoFamilia <> 0 Then
                cboFamilia.SelectedValue = pCodigoFamilia
                cmdNuevoGrupo.Enabled = True
            End If

        Catch ex As Exception
            miClsFuncionalidades.InsertarenLog(Me.Name, "cargarComboFamilias", ex.Message.ToString)
            MessageBox.Show("Error al cargar datos: " & ex.Message)
        Finally
            If DatabaseConnection.Instance.Connection.State = ConnectionState.Open Then
                DatabaseConnection.Instance.Close()
            End If
        End Try
    End Sub

    Private Sub CargarComboGrupos(ByVal pCodigoFamilia As Integer, ByVal pCodigoGrupo As Integer)

        Try
            Dim query As String
            Dim items As New List(Of clsComboboxItem)

            If pCodigoFamilia <> 0 Then
                query = "SELECT DISTINCT GRUPOS.CODIGOGRUPO AS Id, GRUPOS.NOMBRE As Grupo FROM GRUPOS" & _
                        " WHERE GRUPOS.CODIGOFAMILIA = " & pCodigoFamilia

                cboGrupo.DataSource = Nothing
                cboSubgrupo.DataSource = Nothing
                cboGrupo.Items.Clear() 'Limpiar items existentes
                cboSubgrupo.Items.Clear() ' Limpiar items existentes

                cmdNuevoSubgrupo.Enabled = False

                DatabaseConnection.Instance.Open()

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
                                .Descripcion = reader("Grupo").ToString()}
                            items.Add(item)

                            ' Agregar elementos al ComboBox (Texto visible y valor subyacente)
                            'cboGrupo.Items.Add(reader(1))
                        End While

                        cboGrupo.DataSource = items
                        cboGrupo.DisplayMember = "Descripcion"
                        cboGrupo.ValueMember = "Id"

                    End Using
                End Using

                If pCodigoGrupo <> 0 Then
                    cboGrupo.SelectedValue = pCodigoGrupo
                    cmdNuevoSubgrupo.Enabled = True
                End If

            End If
        Catch ex As Exception
            miClsFuncionalidades.InsertarenLog(Me.Name, "CargarComboGrupos", ex.Message.ToString)
            MessageBox.Show("Error al cargar datos: " & ex.Message)
        Finally
            If DatabaseConnection.Instance.Connection.State = ConnectionState.Open Then
                DatabaseConnection.Instance.Close()
            End If
        End Try

    End Sub

    Private Sub CargarComboSubgrupos(ByVal pCodigoGrupo As Integer, ByVal pCodigoSubgrupo As Integer)
        Try
            Dim query As String
            Dim items As New List(Of clsComboboxItem)

            If pCodigoGrupo <> 0 Then
                query = "SELECT DISTINCT SUBGRUPOS.CODIGOSUBGRUPO AS Id, SUBGRUPOS.NOMBRE As subgrupo FROM SUBGRUPOS" & _
                        " WHERE SUBGRUPOS.CODIGOGRUPO = " & pCodigoGrupo

                ' Configurar el ComboBox

                cboSubgrupo.DataSource = Nothing
                cboSubgrupo.Items.Clear() ' Limpiar items existentes

                DatabaseConnection.Instance.Open()

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
                                .Descripcion = reader("Subgrupo").ToString()}
                            items.Add(item)

                            ' Agregar elementos al ComboBox (Texto visible y valor subyacente)
                            'cboGrupo.Items.Add(reader(1))
                        End While

                        cboSubgrupo.DataSource = items
                        cboSubgrupo.DisplayMember = "Descripcion"
                        cboSubgrupo.ValueMember = "Id"

                    End Using
                End Using

                If pCodigoSubgrupo <> 0 Then
                    cboSubgrupo.SelectedValue = pCodigoSubgrupo
                End If

            End If
        Catch ex As Exception
            miClsFuncionalidades.InsertarenLog(Me.Name, "CargarComboSubgrupos", ex.Message.ToString)
            MessageBox.Show("Error al cargar datos: " & ex.Message)
        Finally
            If DatabaseConnection.Instance.Connection.State = ConnectionState.Open Then
                DatabaseConnection.Instance.Close()
            End If
        End Try
    End Sub

    Private Sub cargarComboMarcas(ByVal pCodigoMarca As Integer)

        Dim query As String
        Dim items As New List(Of clsComboboxItem)
        Try
            'Abrimos conexión
            DatabaseConnection.Instance.Open()

            If pCodigoMarca = 0 Then
                'montamos la query quitando la marca 1 que es "sin marca" para las familias de servicios
                query = "SELECT Marcas.CodigoMarca as Id, Marcas.NombreMarca as Nombre FROM MARCAS where Marcas.CodigoMarca <> 1"
            Else
                query = "SELECT Marcas.CodigoMarca as Id, Marcas.NombreMarca as Nombre FROM MARCAS Where Marcas.CodigoMarca =" & pCodigoMarca
            End If

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

                End Using
            End Using

            If pCodigoMarca <> 0 Then
                cboMarca.SelectedValue = pCodigoMarca
            End If

        Catch ex As Exception
            miClsFuncionalidades.InsertarenLog(Me.Name, "cargarComboMarcas", ex.Message.ToString)
            MessageBox.Show("Error al cargar datos: " & ex.Message)
        Finally
            If DatabaseConnection.Instance.Connection.State = ConnectionState.Open Then
                DatabaseConnection.Instance.Close()
            End If
        End Try
    End Sub

    Private Sub CargarComboIVACompra()

        Dim query As String
        Dim items As New List(Of clsComboboxItem)
        Try
            'Abrimos conexión
            DatabaseConnection.Instance.Open()

            'montamos la query
            query = "SELECT configuracion.valor as Id, configuracion.valor as Nombre FROM Configuracion where campo = 'IVA'"

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

                    End While

                    cboIVACompra.DataSource = items
                    cboIVACompra.DisplayMember = "Descripcion"
                    cboIVACompra.ValueMember = "Id"
                End Using
            End Using

        Catch ex As Exception
            miClsFuncionalidades.InsertarenLog(Me.Name, "CargarComboIVACompra", ex.Message.ToString)
            MessageBox.Show("Error al cargar datos: " & ex.Message)
        Finally
            If DatabaseConnection.Instance.Connection.State = ConnectionState.Open Then
                DatabaseConnection.Instance.Close()
            End If
        End Try
    End Sub

    Private Sub CargarComboIVAVenta()
        Dim query As String
        Dim items As New List(Of clsComboboxItem)
        Try
            'Abrimos conexión
            DatabaseConnection.Instance.Open()

            'montamos la query
            query = "SELECT configuracion.valor as Id, configuracion.valor as Nombre FROM Configuracion where campo = 'IVA'"

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
                    End While

                    cboIVAVenta.DataSource = items
                    cboIVAVenta.DisplayMember = "Descripcion"
                    cboIVAVenta.ValueMember = "Id"
                End Using
            End Using

        Catch ex As Exception
            miClsFuncionalidades.InsertarenLog(Me.Name, "CargarComboIVAVenta", ex.Message.ToString)
            MessageBox.Show("Error al cargar datos: " & ex.Message)
        Finally
            If DatabaseConnection.Instance.Connection.State = ConnectionState.Open Then
                DatabaseConnection.Instance.Close()
            End If
        End Try
    End Sub

#End Region


End Class