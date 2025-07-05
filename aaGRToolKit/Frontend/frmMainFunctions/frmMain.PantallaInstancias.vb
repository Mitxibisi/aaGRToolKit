Imports System.IO
Imports System.Text
Imports System.Xml.Serialization
Imports System.Xml.Linq

Partial Public Class frmMain
    ' ------- Pantallas Configuraciones Instancias -------

    ' Botón para resetear configuraciones
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click, Button1.Click
        My.Settings.Reset() ' Restaura las configuraciones guardadas
        My.Settings.Save()  ' Guarda las configuraciones restauradas
        restore_configdata(True) ' Restaura los datos de configuración en la interfaz
        AñadirAttributos() ' Añade los atributos necesarios a la interfaz
        CargarDatosAnteriores(True)
    End Sub

    Private Sub GenerarCSV_Click(sender As Object, e As EventArgs) Handles btnGenerarCSV.Click

        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Filter = "Archivos MLSM y XLSX (*.mlsm;*.xlsx)|*.mlsm;*.xlsx"
        openFileDialog.Title = "Seleccionar un archivo"

        If openFileDialog.ShowDialog() = DialogResult.OK Then
            ImportExcel = openFileDialog.FileName
            aaExcelData.GenerarCsv(ImportExcel,
                                    txtIname.Text,
                                    txtItemplate.Text,
                                    txtTags.Text,
                                    txtDirecc.Text,
                                    True,
                                    CheckBox2.Checked,
                                    txtestados.Text,
                                    txtordenes.Text)
        End If
    End Sub

    ' Botón para cargar datos desde archivo MLSM o XLSX
    Private Sub btnLoadData_Click(sender As Object, e As EventArgs) Handles btnLoadData.Click
        Try
            Dim openFileDialog As New OpenFileDialog()
            openFileDialog.Filter = "Archivos MLSM y XLSX (*.mlsm;*.xlsx)|*.mlsm;*.xlsx"
            openFileDialog.Title = "Seleccionar un archivo"

            If openFileDialog.ShowDialog() = DialogResult.OK Then
                ImportExcel = openFileDialog.FileName ' Se guarda la ruta del archivo

                Dim LoneAttributesInteger As New List(Of String)()
                Dim ArrayAttributesInteger As New List(Of String)()

                ' Procesa los atributos individuales
                For Each attr In LoneAttributes
                    Dim txtloneattr As TextBox = EditAttributes.Controls.OfType(Of TextBox)().FirstOrDefault(Function(txt) txt.Name.ToString() = "txt_" & attr)
                    Dim value As String = txtloneattr.Text

                    If value <> "" Then
                        LoneAttributesInteger.Add(value.ToString()) ' Valor válido
                    Else
                        LoneAttributesInteger.Add("10000") ' Valor por defecto si no es válido
                    End If
                Next

                ' Procesa los atributos de tipo array
                For Each attr In ArrayAttributes
                    Dim txtloneattr As TextBox = EditAttributes.Controls.OfType(Of TextBox)().FirstOrDefault(Function(txt) txt.Name.ToString() = "txt_" & attr)
                    Dim value As String = txtloneattr.Text

                    If value <> "" Then
                        ArrayAttributesInteger.Add(value.ToString()) ' Valor válido
                    Else
                        ArrayAttributesInteger.Add("10000") ' Valor por defecto
                    End If
                Next

                ' Carga los datos desde Excel y añade a la interfaz
                Dim InstancesData = aaExcelData.CargarDatosMapeado(ImportExcel,
                                                                   txtIname.Text,
                                                                   txtItemplate.Text,
                                                                   LoneAttributesInteger,
                                                                   ArrayAttributesInteger,
                                                                   txtIEngUnits.Text,
                                                                   txtMapIndex.Text)
                añadir_Mapeado(InstancesData)
            End If
        Catch y As Exception
            LogBox.Items.Add(y) ' Muestra cualquier error
        End Try
    End Sub

    ' Actualiza ComboBox de atributos cuando se cambia el template
    Private Sub txtTemplateData_TextChanged(sender As Object, e As EventArgs) Handles txtTemplateData.TextChanged
        Dim template = txtTemplateData.Text
        Dim AttrList = aaGalaxyTools.getTemplateAttributes(template)

        ' Limpiar ComboBox antes de añadir nuevos datos
        ComboBox1.Items.Clear()

        ' Verificar que la plantilla es válida
        If AttrList IsNot Nothing AndAlso AttrList.Count > 0 AndAlso AttrList(0) <> "BadTemplate" Then
            For Each Attr In AttrList
                If Not ComboBox1.Items.Contains(Attr) Then
                    ComboBox1.Items.Add(Attr) ' Añade atributo si no existe aún
                End If
            Next
        End If
    End Sub


    ' ------- Pantallas Instancias -------
    Private Sub btnDeploy_Click(sender As Object, e As EventArgs) Handles btnDeploy.Click
        Dim Cascade As Boolean = False
        Dim i As Integer
        ' Contar el total de GroupBox en TabPage2 (TPInstances)
        Dim totalGroupBoxes As Integer = TPInstances.Controls.OfType(Of GroupBox)().Count()
        ' Recorrer cada GroupBox dentro de TPInstances
        For Each grupo As GroupBox In TPInstances.Controls.OfType(Of GroupBox)()
            i = i + 1 ' Incrementar contador i
            ' Actualizar etiqueta de estado con el número de instancia actual
            lblStatus.Text = "Deploy Instance Number " & i & "/" & totalGroupBoxes
            ' Buscar TextBox con Tag "Nombre Instancia" dentro del grupo actual
            Dim txtNombreInstancia As TextBox = grupo.Controls.OfType(Of TextBox)().FirstOrDefault(Function(txt) txt.Tag.ToString() = "Nombre Instancia")
            ' Buscar CheckBox con nombre "cbCascade" en TPInstances
            Dim CCascade As CheckBox = TPInstances.Controls.OfType(Of CheckBox)().FirstOrDefault(Function(txt) txt.Name.ToString() = "cbCascade")
            ' Asignar valor a Cascade según si el checkbox está chequeado
            Cascade = (CCascade IsNot Nothing AndAlso CCascade.Checked)

            ' Verificar que txtNombreInstancia exista antes de llamar a la función
            If txtNombreInstancia IsNot Nothing Then
                ' Llamar función para desplegar o retraer instancia
                aaGalaxyTools.deployUndeployInstance(txtNombreInstancia.Text, Cascade, 0)
            End If
        Next
        ' Limpiar etiqueta de estado
        lblStatus.Text = ""
        ' Mostrar mensaje indicando finalización
        MessageBox.Show("La operación de despliegue ha finalizado correctamente.", "Despliegue completado", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub btnUndeploy_Click(sender As Object, e As EventArgs) Handles btnUndeploy.Click
        Dim i As Integer
        ' Contar total de GroupBox en TPInstances
        Dim totalGroupBoxes As Integer = TPInstances.Controls.OfType(Of GroupBox)().Count()
        ' Recorrer cada GroupBox dentro de TPInstances
        For Each grupo As GroupBox In TPInstances.Controls.OfType(Of GroupBox)()
            i = i + 1 ' Incrementar contador i (pero se reinicia en cada iteración)
            ' Actualizar etiqueta de estado con número de instancia actual
            lblStatus.Text = "UnDeploy Instance Number " & i & "/" & totalGroupBoxes
            ' Buscar TextBox con Tag "Nombre Instancia" dentro del grupo actual
            Dim txtNombreInstancia As TextBox = grupo.Controls.OfType(Of TextBox)().FirstOrDefault(Function(txt) txt.Tag.ToString() = "Nombre Instancia")

            ' Verificar que txtNombreInstancia exista antes de llamar a la función
            If txtNombreInstancia IsNot Nothing Then
                ' Llamar función para retraer instancia (flag 1 indica "undeploy")
                aaGalaxyTools.deployUndeployInstance(txtNombreInstancia.Text, False, 1)
            End If
        Next
        ' Limpiar etiqueta de estado
        lblStatus.Text = ""
        ' Mostrar mensaje indicando finalización
        MessageBox.Show("La operación de retraer instancias ha finalizado correctamente.", "Retracción completada", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub btnCreateInstance_Click(sender As Object, e As EventArgs) Handles btnNewInstance.Click
        If aaGalaxyTools.loggedIn Then
            ' Listas para guardar textos de atributos
            Dim LoneAttributesText As New List(Of String)()
            Dim ArrayAttributesText As New List(Of String)()
            Dim EngUnitsArray As New List(Of String)()

            ' Contar total de GroupBox en TPInstances
            Dim totalGroupBoxes As Integer = TPInstances.Controls.OfType(Of GroupBox)().Count()
            ' Recorrer cada GroupBox dentro de TPInstances
            For Each grupo As GroupBox In TPInstances.Controls.OfType(Of GroupBox)()
                ' Limpiar listas antes de llenar con datos nuevos
                LoneAttributesText.Clear()
                ArrayAttributesText.Clear()
                Dim i As Integer
                i = i + 1 ' Incrementar contador i (pero se reinicia en cada iteración)
                ' Actualizar etiqueta de estado con número de instancia actual
                lblStatus.Text = "Creating Instance Number " & i & "/" & totalGroupBoxes
                ' Buscar TextBox con Tag "Plantilla" dentro del grupo actual
                Dim txtPlantilla As TextBox = grupo.Controls.OfType(Of TextBox)().FirstOrDefault(Function(txt) txt.Tag.ToString() = "Plantilla")
                ' Buscar TextBox con Tag "Nombre Instancia" dentro del grupo actual
                Dim txtNombreInstancia As TextBox = grupo.Controls.OfType(Of TextBox)().FirstOrDefault(Function(txt) txt.Tag.ToString() = "Nombre Instancia")
                ' Buscar TextBox con Tag "Area" en todo TPInstances (no en el grupo)
                Dim txtArea As TextBox = TPInstances.Controls.OfType(Of TextBox)().FirstOrDefault(Function(txt) txt.Tag.ToString() = "Area")
                ' Buscar TextBox con Tag "EngUnits" dentro del grupo actual
                Dim txtEngUnits As TextBox = grupo.Controls.OfType(Of TextBox)().FirstOrDefault(Function(txt) txt.Tag.ToString() = "EngUnits")

                If txtArea.Text = "" Then
                    MessageBox.Show("Por favor, rellene el área.", "Campo vacío", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit Sub
                End If

                ' Recorrer lista LoneAttributes para obtener texto de cada atributo individual
                For Each attr In LoneAttributes
                    ' Buscar TextBox con Tag igual al nombre del atributo
                    Dim txtloneattr As TextBox = grupo.Controls.OfType(Of TextBox)().FirstOrDefault(Function(txt) txt.Tag.ToString() = attr)
                    ' Agregar texto del TextBox a la lista (uso de operador "?." para evitar error si es Nothing)
                    LoneAttributesText.Add(txtloneattr?.Text)
                Next

                ' Recorrer lista ArrayAttributes para obtener texto de cada atributo tipo arreglo
                For Each attr In ArrayAttributes
                    ' Buscar TextBox con Tag igual al nombre del atributo
                    Dim txtloneattr As TextBox = grupo.Controls.OfType(Of TextBox)().FirstOrDefault(Function(txt) txt.Tag.ToString() = attr)
                    ' Agregar texto del TextBox a la lista
                    ArrayAttributesText.Add(txtloneattr?.Text)
                Next

                ' Obtener líneas de texto de txtEngUnits en forma de arreglo de strings
                Dim Lineas() As String = txtEngUnits.Lines

                ' Verificar que txtPlantilla y txtNombreInstancia existan antes de llamar a la función
                If txtPlantilla IsNot Nothing AndAlso txtNombreInstancia IsNot Nothing Then
                    ' Llamar función para crear instancia con todos los parámetros recolectados
                    aaGalaxyTools.CreateInstance(txtPlantilla.Text, txtNombreInstancia.Text,
                                         txtArea?.Text,
                                         LoneAttributesText,
                                         ArrayAttributesText,
                                         LoneAttributes,
                                         ArrayAttributes,
                                         Lineas)
                End If
            Next
            ' Limpiar etiqueta de estado
            lblStatus.Text = ""

            MessageBox.Show("Se han creado nuevas instancias.", "Operación exitosa", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            MessageBox.Show("Es necesario iniciar sesión para continuar.", "Inicio de sesión requerido", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub

    Private Sub btnCreateInstances_Click(sender As Object, e As EventArgs) Handles btnCreateInstance.Click
        ' Refrescar lista de instancias
        Refres_Instances()
    End Sub


    ' ------- Pantallas de Instancias Funciones -------
    Private Sub GenerarDatosExcel(Restablecer As Boolean)
        CambiarAttributos()
        Refres_Instances()

        If (LoneAttributes Is Nothing OrElse LoneAttributes.Length = 0) AndAlso
           (ArrayAttributes Is Nothing OrElse ArrayAttributes.Length = 0) AndAlso
           (ArrayAttributesIndex Is Nothing OrElse ArrayAttributesIndex.Length = 0) AndAlso
           (LoneAttributesIndex Is Nothing OrElse LoneAttributesIndex.Length = 0) Or Restablecer Then

            ' Obtener los datos separados por coma
            Dim AloneAttrIntermedio = My.Settings.AloneAttributes.Split(","c)
            Dim ArrayAttrIntermedio = My.Settings.ArrayAttributes.Split(","c)

            ' Extraer solo el nombre antes del "|"
            LoneAttributes = AloneAttrIntermedio.Select(Function(s) s.Split("|"c)(0)).ToArray()
            ArrayAttributes = ArrayAttrIntermedio.Select(Function(s) s.Split("|"c)(0)).ToArray()

            LoneAttributesIndex = AloneAttrIntermedio.Select(Function(s) s.Split("|"c)(1)).ToArray()
            ArrayAttributesIndex = ArrayAttrIntermedio.Select(Function(s) s.Split("|"c)(1)).ToArray()

            addExcelData(LoneAttributes, ArrayAttributes)

            Dim valoresAlone = AloneAttrIntermedio.Select(Function(s) s.Split("|"c)(1)).ToArray()
            Dim valoresArray = ArrayAttrIntermedio.Select(Function(s) s.Split("|"c)(1)).ToArray()

            Dim totalAlone As Integer = valoresAlone.Length
            Dim totalArray As Integer = valoresArray.Length

            Dim i As Integer = 0 ' índice para Alone
            Dim j As Integer = 0 ' índice para Array

            Dim total As Integer = totalAlone + totalArray
            Dim k As Integer = 0 ' índice general

            For Each ctrl As Control In EditAttributes.Controls
                If TypeOf ctrl Is TextBox Then
                    If k < totalAlone Then
                        ctrl.Text = valoresAlone(i)
                        i += 1
                    ElseIf k - totalAlone < totalArray Then
                        ctrl.Text = valoresArray(j)
                        j += 1
                    End If
                    k += 1
                End If
            Next

        Else
            ' Paso 1: Guardar los valores actuales usando el nombre de atributo
            Dim valoresAnteriores As New Dictionary(Of String, String)

            For Each ctrl As Control In EditAttributes.Controls
                If TypeOf ctrl Is TextBox Then
                    Dim txt = CType(ctrl, TextBox)
                    ' Quitar el prefijo "txt_" para obtener el nombre real del atributo
                    Dim nombreAtributo As String = txt.Name.Replace("txt_", "")
                    valoresAnteriores(nombreAtributo) = txt.Text
                End If
            Next

            ' Paso 2: Llamada que actualiza LoneAttributes y ArrayAttributes
            addExcelData(LoneAttributes, ArrayAttributes)

            ' Paso 3: Reasignar los valores por nombre
            For Each attr In LoneAttributes.Concat(ArrayAttributes)
                Dim nombreControl As String = "txt_" & attr
                Dim ctrl As Control = EditAttributes.Controls(nombreControl)

                If ctrl IsNot Nothing AndAlso TypeOf ctrl Is TextBox Then
                    Dim txt = CType(ctrl, TextBox)
                    If valoresAnteriores.ContainsKey(attr) Then
                        txt.Text = valoresAnteriores(attr)
                    End If
                End If
            Next

        End If
    End Sub

    Private Sub Refres_Instances(Optional ByVal instances As Integer = 0)
        Dim lastTextBoxMultiple As Boolean = False
        Dim numInstancias As Integer

        CambiarAttributos()

        TPInstances.SuspendLayout()

        ' Limpiar campos previos, pero NO el NumericUpDown
        For Each control As Control In TPInstances.Controls
            If Not control.Name = "NumericUpDown2" Then
                control.Dispose() ' Eliminar todos los controles excepto el NumericUpDown
            End If
        Next

        ' Verificar si el NumericUpDown ya existe
        Dim numInstancesSelector As NumericUpDown = TryCast(TPInstances.Controls("NumericUpDown2"), NumericUpDown)

        ' Si no existe, crear uno nuevo
        If numInstancesSelector Is Nothing Then
            numInstancesSelector = New NumericUpDown()
            numInstancesSelector.Name = "NumericUpDown2"
            numInstancesSelector.Location = New Point(170, 10)
            numInstancesSelector.Width = 50
            numInstancesSelector.Minimum = 1
            numInstancesSelector.Maximum = 100
            numInstancesSelector.Value = 1 ' Valor inicial si es necesario

        End If

        ' Crear botón superior
        Dim btnSetInstances As New Button()
        btnSetInstances.Name = "btnCreateInstance"
        btnSetInstances.Text = "Generar instancias"
        btnSetInstances.Location = New Point(10, 10)
        btnSetInstances.Width = 150

        ' Agregar evento al botón para regenerar instancias
        AddHandler btnSetInstances.Click, AddressOf btnCreateInstances_Click

        If instances > 0 Then
            ' Obtener número de instancias desde el NumericUpDown
            numInstancias = instances
        Else
            ' Obtener número de instancias desde el NumericUpDown
            numInstancias = numInstancesSelector.Value
        End If

        ' Crear Label
        Dim areaEtiqueta As New Label()
        areaEtiqueta.Text = "Area"
        areaEtiqueta.Location = New Point(10, 43)
        areaEtiqueta.AutoSize = True

        ' Crear TextBox
        Dim areaTextBox As New TextBox()
        areaTextBox.Width = 150
        areaTextBox.Location = New Point(50, 40)
        areaTextBox.Tag = "Area"

        TPInstances.Controls.Clear()

        TPInstances.Controls.Add(btnSetInstances)
        TPInstances.Controls.Add(numInstancesSelector)
        TPInstances.Controls.Add(areaEtiqueta)
        TPInstances.Controls.Add(areaTextBox)

        ' Definir nombres de los campos y sus Tags
        Dim nombresAtributes As List(Of String) = New List(Of String) From {"Plantilla", "Nombre Instancia"}

        Dim tagsCamposIntermedios As List(Of String) = New List(Of String) From {"Plantilla", "Nombre Instancia"}

        If LoneAttributes.Length > 0 AndAlso Not String.IsNullOrWhiteSpace(LoneAttributes(0)) Then
            nombresAtributes.AddRange(LoneAttributes)
            tagsCamposIntermedios.AddRange(LoneAttributes)
        End If

        If ArrayAttributes.Length > 0 AndAlso Not String.IsNullOrWhiteSpace(ArrayAttributes(0)) Then
            nombresAtributes.AddRange(ArrayAttributes)
            tagsCamposIntermedios.AddRange(ArrayAttributes)
        End If


        ' Convertir a array
        Dim nombresCampos As String() = nombresAtributes.ToArray()
        ' Convertir a array
        Dim tagsCampos As String() = tagsCamposIntermedios.ToArray()

        ' Espaciado inicial (debajo del botón y NumericUpDown)
        Dim yOffset As Integer = 65

        For i As Integer = 0 To numInstancias - 1
            ' Crear un GroupBox para cada instancia
            Dim grupo As New GroupBox()
            Dim Height As Integer = 84
            grupo.Text = "Instancia " & (i + 1)
            grupo.Width = TPInstances.Width - 20
            If LoneAttributes.Length > 0 AndAlso Not String.IsNullOrWhiteSpace(LoneAttributes(0)) Then
                Height += LoneAttributes.Length * 30
            End If

            If ArrayAttributes.Length > 0 AndAlso Not String.IsNullOrWhiteSpace(ArrayAttributes(0)) Then
                Height += (ArrayAttributes.Length + 1) * 60
            End If

            grupo.Height = Height + 20
            grupo.Location = New Point(10, yOffset)

            yOffset += grupo.Height + 10 ' Espaciado entre grupos

            ' Añadir los campos dentro del grupo
            Dim campoOffset As Integer = 30

            For j As Integer = 0 To nombresCampos.Count - 1

                ' Crear Label
                Dim nuevaEtiqueta As New Label()
                nuevaEtiqueta.Text = nombresCampos(j)
                nuevaEtiqueta.Location = New Point(10, campoOffset)
                nuevaEtiqueta.AutoSize = True

                ' Crear TextBox
                Dim nuevoTextBox As New TextBox()
                nuevoTextBox.Width = 150
                nuevoTextBox.Location = New Point(150, campoOffset)
                nuevoTextBox.Tag = tagsCampos(j) ' ASIGNAR EL TAG CORRECTAMENTE
                If Array.IndexOf(ArrayAttributes, tagsCampos(j)) >= 0 Then
                    nuevoTextBox.Height = 50
                    nuevoTextBox.Multiline = True
                    nuevoTextBox.ScrollBars = ScrollBars.Vertical
                    lastTextBoxMultiple = True
                End If

                ' Agregar controles al GroupBox
                grupo.Controls.Add(nuevaEtiqueta)
                grupo.Controls.Add(nuevoTextBox)

                ToolTip1.SetToolTip(nuevoTextBox, "Propiedad " & tagsCampos(j))

                ' Si el TextBox anterior fue multilínea, ajustar solo si es necesario
                If lastTextBoxMultiple Then
                    campoOffset += 10 + nuevoTextBox.Height ' Ajuste basado en la altura
                    lastTextBoxMultiple = False
                Else
                    ' Ajustar campoOffset basado en el tamaño real del TextBox
                    campoOffset += nuevoTextBox.Height + 10 ' Espaciado base entre campos
                End If
            Next

            ' Crear Label
            Dim nuevaEtiqueta1 As New Label()
            nuevaEtiqueta1.Text = "Unidades Mapeado"
            nuevaEtiqueta1.Location = New Point(10, campoOffset)
            nuevaEtiqueta1.AutoSize = True

            ' Crear TextBox
            Dim nuevoTextBox1 As New TextBox()
            nuevoTextBox1.Width = 150
            nuevoTextBox1.Location = New Point(150, campoOffset)
            nuevoTextBox1.Tag = "EngUnits"
            nuevoTextBox1.Height = 50
            nuevoTextBox1.Multiline = True
            nuevoTextBox1.ScrollBars = ScrollBars.Vertical

            grupo.Controls.Add(nuevaEtiqueta1)
            grupo.Controls.Add(nuevoTextBox1)

            ToolTip1.SetToolTip(nuevoTextBox1, "Propiedad Unidades")

            ' Agregar el GroupBox al TabPage
            TPInstances.Controls.Add(grupo)
        Next

        ' Crear botón inferior para procesar instancias
        Dim btnProcessInstances As New Button()
        btnProcessInstances.Name = "btnNewInstance"
        btnProcessInstances.Text = "Crear instancias"
        btnProcessInstances.Location = New Point(10, yOffset)
        btnProcessInstances.Width = 100

        ' Agregar evento al botón para procesar instancias
        AddHandler btnProcessInstances.Click, AddressOf btnCreateInstance_Click

        ' Agregar botón a TabPage
        TPInstances.Controls.Add(btnProcessInstances)

        ' Crear botón inferior para procesar instancias
        Dim btnDeployInstances As New Button()
        btnDeployInstances.Name = "btnDeploy"
        btnDeployInstances.Text = "Deploy"
        btnDeployInstances.Location = New Point(10, btnProcessInstances.Height + 6 + yOffset)
        btnDeployInstances.Width = 100

        ' Agregar evento al botón para procesar instancias
        AddHandler btnDeployInstances.Click, AddressOf btnDeploy_Click
        TPInstances.Controls.Add(btnDeployInstances)

        ' Crear botón inferior para procesar instancias
        Dim btnUnDeployInstances As New Button()
        btnUnDeployInstances.Name = "btnUndeploy"
        btnUnDeployInstances.Text = "UnDeploy"
        btnUnDeployInstances.Location = New Point(10, btnDeployInstances.Height + 35 + yOffset)
        btnUnDeployInstances.Width = 100

        Dim chcmark As New CheckBox()
        chcmark.Name = "cbCascade"
        chcmark.Text = "Modo Cascada"
        chcmark.Tag = "ActivarCascada"
        chcmark.Location = New Point(120, yOffset)

        AddHandler btnUnDeployInstances.Click, AddressOf btnUndeploy_Click

        TPInstances.Controls.Add(btnUnDeployInstances)
        TPInstances.Controls.Add(chcmark)

        TPInstances.ResumeLayout()
    End Sub

    Private Sub añadir_Mapeado(InstancesData)
        Dim i As Integer = 0

        Refres_Instances(InstancesData.Count)

        For Each grupo As GroupBox In TPInstances.Controls.OfType(Of GroupBox)()
            ' Verificar que el índice i no exceda el tamaño de InstancesData
            If i >= InstancesData.Count Then Exit For

            ' Buscar los TextBox dentro de cada GroupBox por Tag
            Dim txtNombreInstancia As TextBox = grupo.Controls.OfType(Of TextBox)().FirstOrDefault(Function(txt) txt.Tag.ToString() = "Nombre Instancia")
            Dim txtPlantilla As TextBox = grupo.Controls.OfType(Of TextBox)().FirstOrDefault(Function(txt) txt.Tag.ToString() = "Plantilla")
            Dim txtEngUnits As TextBox = grupo.Controls.OfType(Of TextBox)().FirstOrDefault(Function(txt) txt.Tag.ToString() = "EngUnits")

            If Not String.IsNullOrWhiteSpace(InstancesData(i).InstanceName) Then
                txtNombreInstancia.Text = InstancesData(i).InstanceName
            End If

            If Not String.IsNullOrWhiteSpace(InstancesData(i).InstanceTemplate) Then
                txtPlantilla.Text = InstancesData(i).InstanceTemplate
            End If

            If Not String.IsNullOrWhiteSpace(InstancesData(i).InstanceTemplate) Then
                Dim attrList = InstancesData(i).EngUnitsArray

                For Each lblattr In attrList
                    txtEngUnits.AppendText(lblattr & Environment.NewLine)
                Next

            End If

            For index As Integer = 0 To Math.Min(InstancesData(i).InstanceAloneAttr.Count, LoneAttributes.Count) - 1
                Dim att As String = LoneAttributes(index)
                Dim attr As String = InstancesData(i).InstanceAloneAttr(index)

                ' Buscar el TextBox por el Tag correspondiente
                Dim txtbox As TextBox = grupo.Controls.OfType(Of TextBox)().FirstOrDefault(Function(txt) txt.Tag.ToString() = att)
                If txtbox IsNot Nothing Then
                    If attr <> "Valor Nulo" Then
                        txtbox.Text = attr
                    Else
                        txtbox.Text = ""
                    End If
                End If
            Next

            For index As Integer = 0 To Math.Min(InstancesData(i).InstanceArrayAttr.Count, ArrayAttributes.Count) - 1
                Dim att As String = ArrayAttributes(index)
                Dim attrList As List(Of String) = InstancesData(i).InstanceArrayAttr(index)

                ' Buscar el TextBox correspondiente por Tag
                Dim txtbox As TextBox = grupo.Controls.OfType(Of TextBox)().FirstOrDefault(Function(txt) txt.Tag.ToString() = att)
                If txtbox IsNot Nothing Then
                    txtbox.Clear() ' Limpiar el contenido anterior
                    For Each lblattr In attrList
                        If lblattr <> "Valor Nulo" Then
                            txtbox.AppendText(lblattr & Environment.NewLine)
                        Else
                            txtbox.AppendText("")
                        End If
                    Next
                End If
            Next
            i = i + 1
        Next
    End Sub

    Private Sub addExcelData(Mylist1 As String(), Mylist2 As String())
        ' Limpiar los controles previos en el Panel
        EditAttributes.Controls.Clear()

        ' Variables para la posición inicial
        Dim startY As Integer = 10
        Dim spacing As Integer = 30 ' Espacio entre cada control

        ' Añadir datos de la primera lista
        startY = AddControlsFromList(Mylist1, startY, spacing)

        ' Añadir datos de la segunda lista
        AddControlsFromList(Mylist2, startY, spacing)

        ' Habilitar el scroll si es necesario
        EditAttributes.AutoScroll = True
    End Sub

    Function AddControlsFromList(MyList As String(), startY As Integer, spacing As Integer)

        EditAttributes.SuspendLayout()

        For i As Integer = 0 To MyList.Length - 1
            ' Crear el Label
            Dim lbl As New Label()
            lbl.Text = MyList(i)
            lbl.Location = New Point(10, startY)
            lbl.AutoSize = True

            ' Crear el TextBox
            Dim txt As New TextBox()
            txt.Name = "txt_" & MyList(i) ' Nombre único para cada TextBox
            txt.Location = New Point(lbl.Width + 10, startY)
            txt.Width = 100
            txt.Text = ""

            ' Añadir los controles al Panel
            EditAttributes.Controls.Add(lbl)
            EditAttributes.Controls.Add(txt)

            ToolTip1.SetToolTip(txt, "Propiedad " & MyList(i))

            ' Incrementar la posición Y para el siguiente control
            startY += spacing
        Next

        EditAttributes.ResumeLayout()

        Return startY
    End Function
End Class