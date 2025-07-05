Imports System.IO
Imports System.Text
Imports System.Xml.Serialization
Imports System.Xml.Linq

Partial Public Class frmMain
    ' ------- Acciones componentes -------
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnNewTemplate.Click
        Dim List As New List(Of String())

        Dim Discrete1 As New List(Of String)
        Dim Discrete2 As New List(Of String)
        Dim Analogic1 As New List(Of String)
        Dim Analogic2 As New List(Of String)

        Discrete1 = ConventirDatos(DiscreteTemplateData)
        Analogic1 = ConventirDatos(AnalogTemplateData)

        List.Add(Discrete1.ToArray)
        List.Add(DiscreteString8)
        List.Add(Analogic1.ToArray)
        List.Add(AnalogicString8)

        aaTemplateDiscAttrList = CreateAttrData(TPDiscFA)

        aaTemplateAnalogAttrList = CreateAttrData(TPFAAnalog)

        aaTemplateScriptList = CreateAttrData(TPScript)

        aaTemplateUDAList = CreateAttrData(TPUDA)

        aaGalaxyTools.CreateTemplate(txtDTemplate.Text,
                                     txtNewTName.Text,
                                     aaTemplateDiscAttrList,
                                     aaTemplateAnalogAttrList,
                                     aaTemplateScriptList,
                                     aaTemplateUDAList,
                                     List)

        MessageBox.Show("Plantilla creada con éxito.", "Operación completada", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub btnUploadTemplateAttr_Click(sender As Object, e As EventArgs) Handles btnUploadTemplateAttr.Click
        ActualizarPantallas(False)
    End Sub

    Private Sub btnTemplateDataAdd_Click(sender As Object, e As EventArgs) Handles btnTemplateDataAdd.Click
        Dim Discrete1 As New List(Of Integer)
        Dim Discrete2 As New List(Of Integer)
        Dim Discrete3 As New List(Of Integer)
        Dim Discrete4 As New List(Of Integer)

        Dim Integer1 As New List(Of Integer)
        Dim Integer2 As New List(Of Integer)
        Dim Integer3 As New List(Of Integer)
        Dim Integer4 As New List(Of Integer)

        Try
            Dim openFileDialog As New OpenFileDialog()
            openFileDialog.Filter = "Archivos MLSM y XLSX (*.mlsm;*.xlsx;*.xlsm)|*.mlsm;*.xlsx;*.xlsm"
            openFileDialog.Title = "Seleccionar un archivo"

            If openFileDialog.ShowDialog() = DialogResult.OK Then
                ImportExcel = openFileDialog.FileName

                For Each Item In DiscreteTemplateData
                    If Item.AttributeName <> "Nombre" Then
                        Select Case Item.Modo
                            Case 1
                                Discrete1.Add(Item.Columna(0))
                                Discrete3.Add(Item.Columna(1))
                            Case 2
                                For Each inx In Item.Columna
                                    Discrete3.Add(inx)
                                Next
                            Case 3
                                For Each inx In Item.Columna
                                    Discrete2.Add(inx)
                                Next
                            Case 4
                                For Each inx In Item.Columna
                                    Discrete3.Add(inx)
                                Next

                                For Each inx In Item.SubColumn
                                    Discrete4.Add(inx)
                                Next
                        End Select
                    End If
                Next

                For Each Item In AnalogTemplateData
                    If Item.AttributeName <> "Nombre" Then
                        Select Case Item.Modo
                            Case 1
                                Integer1.Add(Item.Columna(0))
                                Integer3.Add(Item.Columna(1))
                            Case 2
                                For Each inx In Item.Columna
                                    Integer3.Add(inx)
                                Next
                            Case 3
                                For Each inx In Item.Columna
                                    Integer2.Add(inx)
                                Next
                            Case 4
                                For Each inx In Item.Columna
                                    Integer3.Add(inx)
                                Next

                                For Each inx In Item.SubColumn
                                    Integer4.Add(inx)
                                Next
                        End Select
                    End If
                Next

                Dim InstancesData As aaExcelData.aaTemplateData = aaExcelData.CargarDatosPlantilla(ImportExcel,
                                                                                                   Discrete1.ToArray,
                                                                                                   Discrete2.ToArray,
                                                                                                   Discrete3.ToArray,
                                                                                                   Discrete4.ToArray,
                                                                                                   Integer1.ToArray,
                                                                                                   Integer2.ToArray,
                                                                                                   Integer3.ToArray,
                                                                                                   Integer4.ToArray,
                                                                                                   UDAInteger1,
                                                                                                   ScriptInteger1)

                Dim D = 0
                Dim A = 0
                Dim U = 0
                Dim S = 0

                If InstancesData.Discrete IsNot Nothing Then
                    NUPdfa.Value = InstancesData.Discrete.Count
                End If
                If InstancesData.Analogic IsNot Nothing Then
                    NUPafa.Value = InstancesData.Analogic.Count
                End If
                If InstancesData.UDA IsNot Nothing Then
                    NUPuda.Value = InstancesData.UDA.Count
                End If
                If InstancesData.Script IsNot Nothing Then
                    NUPscr.Value = InstancesData.Script.Count
                End If

                ActualizarPantallas(True, InstancesData)
            End If
        Catch y As Exception
            LogBox.Items.Add(y)
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim openFileDialog As New OpenFileDialog()
        Dim Result As New List(Of String)
        Dim ResultList As New List(Of List(Of String))

        openFileDialog.Filter = "Archivos aaPKG y CSV (*.aaPKG;*.csv)|*.aaPKG;*.csv"
        openFileDialog.Title = "Seleccionar archivo(s) que desea importar"
        openFileDialog.Multiselect = True

        If openFileDialog.ShowDialog() = DialogResult.OK Then
            For Each FileName In openFileDialog.FileNames
                Debug.WriteLine(FileName)
                Result = aaGalaxyTools.GalaxyImport(FileName, CBOverwrite.Checked)
                ResultList.Add(Result)
            Next

            For Each Result In ResultList
                For Each Message In Result
                    LogBox.Items.Add(Message)
                Next
            Next
        Else
            MsgBox("Fallo al Importar el archivo ya que no es del tipo esperado")
        End If
    End Sub

    Private Sub NuevoAttributo_Click(sender As Object, e As EventArgs) Handles btnNuevoAttributo.Click
        Dim miForm As New CustomInputBox()
        If miForm.ShowDialog() = DialogResult.OK Then
            Dim textoIngresado As String = miForm.ValorTexto
            Dim checkMarcado As Boolean = miForm.OpcionMarcada

            If textoIngresado <> "" Then
                AñadirAttributo(textoIngresado, checkMarcado)
            End If
        End If
    End Sub

    Private Sub Eliminar_Click(sender As Object, e As EventArgs) Handles btneliminarattr.Click
        Dim Index As Integer = CheckedListBox1.SelectedIndex
        If Index >= 0 Then
            CheckedListBox1.Items.RemoveAt(Index)
        End If
    End Sub

    Private Sub ModificarInstancias_Click(sender As Object, e As EventArgs) Handles btnModificarInstancias.Click
        GenerarDatosExcel(False)
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        EliminarListBox(DiscreteTemplateData)
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        ProcesarDatos(GPDiscreteTemplatesConfig, DiscreteTemplateData)

        GenerarPantallasAttributos(DiscreteTemplateData, "Disc", NUPdfa.Value, TPDiscFA, Nothing)

        Button6.Enabled = False
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim attr = InputBox("Nombre del nuevo atributo")

        If attr <> "" Then
            CargarDatosPlantillas(attr, GPDiscreteTemplatesConfig)
        End If
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Dim attr = InputBox("Nombre del nuevo atributo")

        If attr <> "" Then
            CargarDatosPlantillas(attr, GBAnalogTemplateConfig)
        End If
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        EliminarListBox(AnalogTemplateData)
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        ProcesarDatos(GBAnalogTemplateConfig, AnalogTemplateData)

        GenerarPantallasAttributos(AnalogTemplateData, "Analog", NUPafa.Value, TPFAAnalog, Nothing)

        Button7.Enabled = False
    End Sub

    Private Sub GPDiscreteTemplatesConfig_Enter(sender As Object, e As EventArgs) Handles GPDiscreteTemplatesConfig.ControlAdded
        Button6.Enabled = True
    End Sub

    Private Sub GBAnalogTemplateConfig_Enter(sender As Object, e As EventArgs) Handles GBAnalogTemplateConfig.ControlAdded
        Button7.Enabled = True
    End Sub

    Public Sub UpdateTemplateAttr(sender As Object, e As EventArgs) _
                                    Handles txtUDA1.TextChanged,
                                            txtScript1.TextChanged
        RefresTemplateTabs()
    End Sub


    ' ------- Funciones creacion de plantillas -------
    Private Function CreateAttrData(TP As TabPage)
        Dim AttrList As New List(Of aaTemplateData.aaNewField)

        For Each grupo As GroupBox In TP.Controls.OfType(Of GroupBox)()
            Dim TextBoxValues As New List(Of String)()
            Dim CheckBoxValues As New List(Of Boolean)()
            Dim GroupBoxValues As New List(Of String)()
            Dim ComboBoxValues As New List(Of String)()
            Dim AttrStrings As New List(Of String)

            ' Recopilar valores de TextBox y CheckBox dentro del GroupBox
            For Each ctrl As Control In grupo.Controls
                If TypeOf ctrl Is TextBox Then
                    TextBoxValues.Add(CType(ctrl, TextBox).Text)
                ElseIf TypeOf ctrl Is ComboBox Then
                    ComboBoxValues.Add(CType(ctrl, ComboBox).SelectedItem & "|" & CType(ctrl, ComboBox).SelectedIndex + 1)
                ElseIf TypeOf ctrl Is CheckBox Then
                    CheckBoxValues.Add(CType(ctrl, CheckBox).Checked)
                ElseIf TypeOf ctrl Is GroupBox Then
                    For Each Subctrl In ctrl.Controls
                        If TypeOf Subctrl Is TextBox Then
                            GroupBoxValues.Add((CType(Subctrl, TextBox).Text))
                        End If
                    Next
                End If
            Next

            Dim minCount As Integer = Math.Min(CheckBoxValues.Count, TextBoxValues.Count)

            ' Mostrar los valores emparejados

            Dim Name = (TextBoxValues(0) & "|" & CheckBoxValues(0))


            For i = 1 To TextBoxValues.Count - 1
                AttrStrings.Add(TextBoxValues(i) & "|" & CheckBoxValues(i))
            Next

            For i = 0 To ComboBoxValues.Count - 1
                AttrStrings.Add(ComboBoxValues(i))
            Next

            For i = minCount To CheckBoxValues.Count - 1 Step 2
                AttrStrings.Add(CheckBoxValues(i) & "|" & CheckBoxValues(i + 1))
            Next

            For i = 0 To GroupBoxValues.Count - 1
                AttrStrings.Add(GroupBoxValues(i) & "|" & "False")
            Next

            Dim attr As New aaTemplateData.aaNewField(Name, AttrStrings)
            AttrList.Add(attr)
        Next

        Return AttrList
    End Function

    Private Sub addControlsToTab(ByVal str As String,
                                 ByVal numInstancias As Integer,
                                 ByVal Page As TabPage,
                                 ByVal txtCampos As String(),
                                 ByVal txtComboBox As String,
                                 ByVal txtCheckBox As String(),
                                 GrupBox As List(Of String()),
                                 Optional WorkData As aaExcelData.aaTemplateData = Nothing)

        Page.SuspendLayout()

        Page.Controls.Clear()

        Dim nombresCampos As String() = txtCampos
        Dim CombBox = ComboBoxSplit(txtComboBox)
        Dim tagsCampos As String() = txtCampos
        Dim tagsCheckbox As String() = txtCheckBox
        Dim GlobalIndex As Integer = 0

        ' Espaciado inicial
        Dim yOffset As Integer = 10

        For i As Integer = 0 To numInstancias - 1
            Dim TextAttr As List(Of String) = Nothing
            Dim CheckAttr As List(Of String) = Nothing
            Dim GroupAttr As List(Of String) = Nothing
            Dim AttrComboText As String = Nothing

            If WorkData IsNot Nothing And str = "Disc" Then
                AttrComboText = WorkData.Discrete(i).AccessMode
                TextAttr = New List(Of String)
                CheckAttr = New List(Of String)()
                TextAttr.Add(WorkData.Discrete(i).AttrName)
                TextAttr.AddRange(WorkData.Discrete(i).TextBoxAttr)
                CheckAttr.Add("False")
                CheckAttr.AddRange(WorkData.Discrete(i).CheckBoxValue)
                GroupAttr = WorkData.Discrete(i).GroupBoxAttr
            ElseIf WorkData IsNot Nothing And str = "Analog" Then
                AttrComboText = WorkData.Analogic(i).AccessMode
                TextAttr = New List(Of String)
                CheckAttr = New List(Of String)()
                TextAttr.Add(WorkData.Analogic(i).AttrName)
                TextAttr.AddRange(WorkData.Analogic(i).TextBoxAttr)
                CheckAttr.Add("False")
                CheckAttr.AddRange(WorkData.Analogic(i).CheckBoxValue)
                GroupAttr = WorkData.Analogic(i).GroupBoxAttr
            ElseIf WorkData IsNot Nothing And str = "UDA" Then
                TextAttr = New List(Of String)
                TextAttr.Add(WorkData.UDA(i).Name)
                TextAttr.Add(WorkData.UDA(i).Value)
            ElseIf WorkData IsNot Nothing And str = "Scr" Then
                TextAttr = New List(Of String)
                GroupAttr = New List(Of String)
                TextAttr.Add(WorkData.Script(i).Name)
                TextAttr.Add(WorkData.Script(i).Expression)
                TextAttr.Add(WorkData.Script(i).ExecuteText)
                AttrComboText = WorkData.Script(i).TriggerType
            End If

            Dim Index As Integer = If(TextAttr IsNot Nothing, TextAttr.Count, 0)

            ' Crear un GroupBox para cada instancia
            Dim grupo As New GroupBox()
            grupo.Text = "Attributo " & str & (i + 1)
            grupo.Width = Page.Width - 20
            grupo.Location = New Point(10, yOffset)
            grupo.AutoSize = True

            ' Añadir los campos dentro del grupo
            Dim campoOffset As Integer = 20

            For j As Integer = 0 To nombresCampos.Count - 1

                ' Crear Label
                Dim nuevaEtiqueta As New Label()
                nuevaEtiqueta.Text = nombresCampos(j)
                nuevaEtiqueta.Location = New Point(10, campoOffset)

                ' Crear TextBox
                Dim nuevoTextBox As New TextBox()
                nuevoTextBox.Width = 150
                nuevoTextBox.Height = 100
                nuevoTextBox.Location = New Point(15 + nuevaEtiqueta.Width, campoOffset)
                nuevoTextBox.Tag = str & "_" & tagsCampos(j)
                If TextAttr IsNot Nothing AndAlso j >= 0 AndAlso j < TextAttr.Count Then
                    If TextAttr(j) IsNot Nothing Then
                        nuevoTextBox.Text = TextAttr(j)
                    End If
                End If
                If tagsCampos(j) = "Execution Text" Then
                    nuevoTextBox.Multiline = True
                End If

                ' Crear CheckBox
                Dim nuevoCheckBox As New CheckBox()
                nuevoCheckBox.Width = 100
                nuevoCheckBox.Text = "Lock"
                nuevoCheckBox.Location = New Point(25 + nuevaEtiqueta.Width + nuevoTextBox.Width, campoOffset)
                nuevoCheckBox.Tag = str & "_L_" & tagsCampos(j)
                If CheckAttr IsNot Nothing Then
                    If CheckAttr(j) IsNot Nothing And CheckAttr(j) = "True" Then
                        nuevoCheckBox.Checked = True
                    End If
                End If

                ' Agregar controles al GroupBox
                grupo.Controls.Add(nuevaEtiqueta)
                grupo.Controls.Add(nuevoTextBox)
                grupo.Controls.Add(nuevoCheckBox)

                ToolTip1.SetToolTip(nuevoCheckBox, "Bloquear Propiedad " & tagsCampos(j))
                ToolTip1.SetToolTip(nuevoTextBox, "Propiedad " & tagsCampos(j))

                ' Ajustar campoOffset basado en el tamaño real del TextBox
                campoOffset += nuevoTextBox.Height + 10 ' Espaciado base entre campos
            Next

            For Each List In CombBox
                ' Crear Label
                Dim nuevaEtiqueta As New Label()
                nuevaEtiqueta.Text = List(0)
                nuevaEtiqueta.Location = New Point(10, campoOffset)

                ' Crear TextBox
                Dim nuevoComboBox As New ComboBox
                nuevoComboBox.Width = 150
                nuevoComboBox.Location = New Point(15 + nuevaEtiqueta.Width, campoOffset)
                nuevoComboBox.Tag = str & "_" & List(0)

                If List.Length > 0 Then
                    For j As Integer = 1 To List.Length - 1
                        nuevoComboBox.Items.Add(List(j))
                    Next
                End If

                If AttrComboText IsNot Nothing Then
                    nuevoComboBox.SelectedItem = AttrComboText
                End If

                ' Agregar controles al GroupBox
                grupo.Controls.Add(nuevaEtiqueta)
                grupo.Controls.Add(nuevoComboBox)

                ToolTip1.SetToolTip(nuevoComboBox, "Propiedad " & List(0))

                ' Ajustar campoOffset basado en el tamaño real del TextBox
                campoOffset += nuevoComboBox.Height + 10 ' Espaciado base entre campos
            Next

            For j As Integer = 0 To tagsCheckbox.Count - 1
                ' Crear TextBox
                Dim nuevoCheckBox As New CheckBox()
                nuevoCheckBox.Width = 100
                nuevoCheckBox.Text = tagsCheckbox(j)
                nuevoCheckBox.Location = New Point(10, campoOffset)
                nuevoCheckBox.Tag = str & "_" & tagsCheckbox(j) ' ASIGNAR EL TAG CORRECTAMENTE
                If CheckAttr IsNot Nothing Then
                    If CheckAttr(Index) IsNot Nothing And CheckAttr(Index) = "True" Then
                        nuevoCheckBox.Checked = True
                    End If
                    Index += 1
                End If

                Dim nuevoLockCheckBox As New CheckBox()
                nuevoLockCheckBox.Width = 100
                nuevoLockCheckBox.Text = "Lock"
                nuevoLockCheckBox.Location = New Point(nuevoCheckBox.Width + 10, campoOffset)
                nuevoLockCheckBox.Tag = str & "_L_" & tagsCheckbox(j) ' ASIGNAR EL TAG CORRECTAMENTE
                If CheckAttr IsNot Nothing Then
                    If CheckAttr(Index) IsNot Nothing And CheckAttr(Index) = "True" Then
                        nuevoLockCheckBox.Checked = True
                    End If
                    Index += 1
                End If

                grupo.Controls.Add(nuevoCheckBox)
                grupo.Controls.Add(nuevoLockCheckBox)

                ToolTip1.SetToolTip(nuevoCheckBox, "Propiedad " & tagsCheckbox(j))
                ToolTip1.SetToolTip(nuevoLockCheckBox, "Bloquear Propiedad " & tagsCheckbox(j))

                ' Ajustar campoOffset basado en el tamaño real del TextBox
                campoOffset += nuevoCheckBox.Height + 1 ' Espaciado base entre campos
            Next

            Dim IndexG As Integer = 0

            For Each Subg In GrupBox
                If Subg IsNot Nothing AndAlso Subg.Length > 0 Then
                    ' Crear CheckBox
                    Dim nuevoCheckBox As New CheckBox()
                    nuevoCheckBox.Width = 100
                    nuevoCheckBox.Text = Subg(0)
                    nuevoCheckBox.Location = New Point(10, campoOffset)
                    nuevoCheckBox.Tag = str & "_" & Subg(0) & (i + 1)
                    If CheckAttr IsNot Nothing Then
                        If CheckAttr(Index) IsNot Nothing And CheckAttr(Index) = "True" Then
                            nuevoCheckBox.Checked = True
                        End If
                        Index += 1
                    End If

                    ' Crear CheckBox
                    Dim nuevoLockCheckBox As New CheckBox()
                    nuevoLockCheckBox.Width = 100
                    nuevoLockCheckBox.Text = "Lock"
                    nuevoLockCheckBox.Location = New Point(nuevoCheckBox.Width + 10, campoOffset)
                    nuevoLockCheckBox.Tag = str & "_L_" & Subg(0) & (i + 1)
                    If CheckAttr IsNot Nothing Then
                        If CheckAttr(Index) IsNot Nothing And CheckAttr(Index) = "True" Then
                            nuevoLockCheckBox.Checked = True
                        End If
                        Index += 1
                    End If

                    ' Asignar el evento CheckedChanged al CheckBox
                    AddHandler nuevoCheckBox.CheckedChanged, AddressOf CheckBox_CheckedChanged

                    grupo.Controls.Add(nuevoCheckBox)
                    grupo.Controls.Add(nuevoLockCheckBox)

                    campoOffset += nuevoCheckBox.Height + 2 ' Espaciado base entre campos

                    Dim subgrupo As New GroupBox
                    subgrupo.Name = Subg(0)
                    subgrupo.Text = Subg(0)
                    subgrupo.Width = grupo.Width - 20
                    subgrupo.Height = ((Subg.Length * 26)) + 20
                    subgrupo.Location = New Point(10, campoOffset)
                    subgrupo.Tag = str & "_" & Subg(0) & (i + 1)
                    subgrupo.Enabled = False
                    subgrupo.AutoSize = True

                    Dim subcampoOffset As Integer = 20

                    For e As Integer = 1 To Subg.Length - 1
                        ' Crear Label
                        Dim nuevaEtiqueta As New Label()
                        nuevaEtiqueta.Text = str & "_" & Subg(e)
                        nuevaEtiqueta.Location = New Point(15, subcampoOffset)
                        nuevaEtiqueta.AutoSize = True

                        ' Crear TextBox
                        Dim nuevoTextBox As New TextBox()
                        nuevoTextBox.Width = 150
                        nuevoTextBox.Location = New Point(35 + nuevaEtiqueta.Width, subcampoOffset)
                        nuevoTextBox.Tag = str & "_" & Subg(e)
                        If GroupAttr IsNot Nothing Then
                            If GroupAttr(IndexG) IsNot Nothing Then
                                nuevoTextBox.Text = GroupAttr(IndexG)
                            End If
                            IndexG += 1
                        End If

                        ' Agregar controles al GroupBox
                        subgrupo.Controls.Add(nuevaEtiqueta)
                        subgrupo.Controls.Add(nuevoTextBox)

                        subcampoOffset += nuevoTextBox.Height + 10
                    Next

                    campoOffset += subgrupo.Height + 10

                    grupo.Height = campoOffset + 20

                    grupo.Controls.Add(subgrupo)
                Else
                End If
            Next

            ' Ajustar la altura del GroupBox principal
            grupo.Height = campoOffset + 20

            yOffset += grupo.Height + 20 ' Espaciado entre grupos

            ' Agregar el GroupBox al TabPage
            Page.Controls.Add(grupo)

            Page.ResumeLayout()
        Next
    End Sub

    Public Sub RefresTemplateTabs()
        DiscreteString8 = If(String.IsNullOrEmpty(DiscreteBlockAttr.Text), New String() {}, DiscreteBlockAttr.Text.Split(","c).Select(Function(sl) sl.Trim()).ToArray())
        AnalogicString8 = If(String.IsNullOrEmpty(AnalogBlockAttr.Text), New String() {}, AnalogBlockAttr.Text.Split(","c).Select(Function(sl) sl.Trim()).ToArray())

        ' UDA
        UDAInteger1 = If(String.IsNullOrEmpty(txtUDA1.Text), New Integer() {}, txtUDA1.Text.Split(","c).Select(Function(sl) Integer.Parse(sl.Trim())).ToArray())

        ' Script
        ScriptInteger1 = If(String.IsNullOrEmpty(txtScript1.Text), New Integer() {}, txtScript1.Text.Split(","c).Select(Function(sl) Integer.Parse(sl.Trim())).ToArray())
    End Sub

    Private Sub CheckBox_CheckedChanged(sender As Object, e As EventArgs)
        ' Convierte el sender a CheckBox
        Dim checkBox As CheckBox = CType(sender, CheckBox)

        ' Busca el contenedor del CheckBox, podría ser un Panel o TabPage si los controles están dentro de uno
        Dim parentContainer As Control = checkBox.Parent

        ' Recorre los controles dentro del contenedor específico (puede ser Panel, TabPage, etc.)
        For Each ctrl As Control In parentContainer.Controls
            ' Verifica si el control es un GroupBox
            If TypeOf ctrl Is GroupBox Then
                ' Compara el Tag del GroupBox con el Tag del CheckBox
                If ctrl.Tag = checkBox.Tag Then
                    ' Habilita o deshabilita el GroupBox dependiendo del estado del CheckBox
                    ctrl.Enabled = checkBox.Checked
                End If
            End If
        Next
    End Sub

    Private Function ComboBoxSplit(input As String)
        ' Crear la lista principal
        Dim lista1 As New List(Of String())

        input = input.Replace("{", "")
        input = input.Replace("},", "|")

        Dim groups As String() = input.Trim("{"c, "}"c).Split("|"c)

        ' Recorrer cada grupo
        For Each group As String In groups
            ' Separar los elementos por comas y agregar el grupo como un array de strings
            lista1.Add(group.Split(","c).Select(Function(s) s.Trim()).ToArray())
        Next

        Return lista1
    End Function

    Private Sub AñadirAttributo(NuevoElemento As String, Array As Boolean)
        CheckedListBox1.Items.Add(NuevoElemento, Array)
    End Sub

    Private Sub CambiarAttributos()
        Dim WorkLoneAttr As New List(Of String)
        Dim WorkArrayAttr As New List(Of String)

        For i As Integer = 0 To CheckedListBox1.Items.Count - 1
            Dim itemTexto As String = CheckedListBox1.Items(i).ToString()
            Dim estaMarcado As Boolean = CheckedListBox1.GetItemChecked(i)

            If estaMarcado Then
                WorkArrayAttr.Add(itemTexto)
            Else
                WorkLoneAttr.Add(itemTexto)
            End If
        Next

        ArrayAttributes = WorkArrayAttr.ToArray()
        LoneAttributes = WorkLoneAttr.ToArray()
    End Sub

    Private Sub AñadirAttributos()
        CheckedListBox1.Items.Clear()

        For Each Alone In LoneAttributes
            AñadirAttributo(Alone, False)
        Next
        For Each Array In ArrayAttributes
            AñadirAttributo(Array, True)
        Next
    End Sub

    Private Sub CargarDatosAnteriores(Optional ResetX As Boolean = False)
        GPDiscreteTemplatesConfig.Controls.Clear()
        GBAnalogTemplateConfig.Controls.Clear()


        Dim BooleanTemp = ResetX
        For Each Item In DiscreteTemplateData
            CargarDatosPlantillas(Item.AttributeName, GPDiscreteTemplatesConfig, Item, BooleanTemp)
            If BooleanTemp Then
                BooleanTemp = False
            End If
        Next

        BooleanTemp = ResetX
        For Each Item In AnalogTemplateData
            CargarDatosPlantillas(Item.AttributeName, GBAnalogTemplateConfig, Item, BooleanTemp)
            If BooleanTemp Then
                BooleanTemp = False
            End If
        Next
    End Sub

    Private Sub CargarDatosPlantillas(Lista As String, grupo As GroupBox, Optional Attribute As TemplateData = Nothing, Optional ResetX As Boolean = False)
        Dim espacioVertical As Integer = 5
        Dim xOffset As Integer
        Static xOffsetList As Integer() = {20, 20, 20}
        Dim checkBoxCount As Integer = 4
        Static index As Integer = 0

        Dim yOffsetLocal As Integer = 20 ' Resetea para cada columna

        grupo.SuspendLayout()

        If grupo.Name = GPDiscreteTemplatesConfig.Name Then
            xOffset = xOffsetList(0)
        ElseIf grupo.Name = GBAnalogTemplateConfig.Name Then
            xOffset = xOffsetList(1)
        End If


        If Attribute IsNot Nothing And index = 0 Or ResetX Then
            xOffset = 20
            index = 1
        End If

        If Attribute Is Nothing Then
            index = 0
        End If

        ' Crear y posicionar Label
        Dim nuevaEtiqueta As New Label() With {
                .Text = Lista,
                .AutoSize = True,
                .Location = New Point(xOffset, yOffsetLocal)
            }
        grupo.Controls.Add(nuevaEtiqueta)

        yOffsetLocal += nuevaEtiqueta.Height + espacioVertical

        ' Crear CheckBoxes y TextBoxes
        For i As Integer = 1 To checkBoxCount
            Dim nuevoCheckBox As New CheckBox() With {
                    .Width = 30,
                    .Location = New Point(xOffset, yOffsetLocal),
                    .Checked = False,
                    .Name = $"chk_{i}_{Lista}"
                }
            If Attribute IsNot Nothing Then
                If Attribute.Modo = i Then
                    nuevoCheckBox.Checked = True
                End If
            End If


            grupo.Controls.Add(nuevoCheckBox)

            ' Agregar TextBox debajo de CheckBox 3 y 4
            If i = 3 Or i = 4 Then
                Dim nuevoTextBox As New TextBox() With {
                        .Width = 120,
                        .Height = 40,
                        .Location = New Point(nuevoCheckBox.Location.X + nuevoCheckBox.Width, yOffsetLocal),
                        .Tag = $"txt_{i}_{Lista}"
                    }

                AddHandler nuevoTextBox.TextChanged, AddressOf TextBox_TextChanged

                If Attribute IsNot Nothing Then
                    If Attribute.Modo = i Then
                        nuevoTextBox.Text = String.Join(",", Attribute.Attributos)
                    End If
                End If

                grupo.Controls.Add(nuevoTextBox)
            End If

            yOffsetLocal += nuevoCheckBox.Height + espacioVertical

        Next

        Dim nuevoTextBoxColumn As New TextBox() With {
                        .Width = 150,
                        .Height = 40,
                        .Location = New Point(xOffset, yOffsetLocal),
                        .Tag = $"txt_column_{Lista}"
                    }

        If Attribute IsNot Nothing Then
            nuevoTextBoxColumn.Text = String.Join(",", Attribute.Columna)
        End If

        grupo.Controls.Add(nuevoTextBoxColumn)

        yOffsetLocal += nuevoTextBoxColumn.Height + espacioVertical

        Dim nuevoTextBoxNameIDE As New TextBox() With {
                        .Width = 150,
                        .Height = 40,
                        .Location = New Point(xOffset, yOffsetLocal),
                        .Tag = $"txt_IDE_{Lista}"
                    }

        If Attribute IsNot Nothing Then
            nuevoTextBoxNameIDE.Text = Attribute.NameIDE
        End If

        grupo.Controls.Add(nuevoTextBoxNameIDE)

        yOffsetLocal += nuevoTextBoxNameIDE.Height + espacioVertical

        Dim nuevoTextBoxSubNamesIDE As New TextBox() With {
                        .Width = 150,
                        .Height = 40,
                        .Location = New Point(xOffset, yOffsetLocal),
                        .Tag = $"txt_SubNameIDE_{Lista}"
                    }

        If Attribute IsNot Nothing Then
            nuevoTextBoxSubNamesIDE.Text = String.Join(",", Attribute.SubNameIDE)
        End If

        nuevoTextBoxSubNamesIDE.Enabled = False

        grupo.Controls.Add(nuevoTextBoxSubNamesIDE)

        yOffsetLocal += nuevoTextBoxSubNamesIDE.Height + espacioVertical

        Dim nuevoTextBoxSubNamesIDEColumn As New TextBox() With {
                        .Width = 150,
                        .Height = 40,
                        .Location = New Point(xOffset, yOffsetLocal),
                        .Tag = $"txt_SubNameIDEColumn_{Lista}"
                    }

        If Attribute IsNot Nothing Then
            nuevoTextBoxSubNamesIDEColumn.Text = String.Join(",", Attribute.SubColumn)
        End If

        nuevoTextBoxSubNamesIDEColumn.Enabled = False

        grupo.Controls.Add(nuevoTextBoxSubNamesIDEColumn)

        yOffsetLocal += nuevoTextBoxSubNamesIDE.Height + espacioVertical

        ' Mover a la siguiente columna
        xOffset += 160

        If grupo.Name = GPDiscreteTemplatesConfig.Name Then
            xOffsetList(0) = xOffset
        ElseIf grupo.Name = GBAnalogTemplateConfig.Name Then
            xOffsetList(1) = xOffset
        End If

        grupo.ResumeLayout()
    End Sub

    Private Sub ProcesarDatos(grupo As GroupBox, DataList As List(Of TemplateData))
        For Each ctrl In grupo.Controls
            If TypeOf ctrl Is CheckBox Then
                Dim chk As CheckBox = CType(ctrl, CheckBox)

                If chk.Checked Then
                    Dim nameParts As String() = chk.Name.ToString().Split("_"c)

                    If nameParts.Length >= 3 Then
                        Dim i As Integer

                        ' Parsear el número i desde el nombre "chk_i_Item"
                        If Integer.TryParse(nameParts(1), i) Then
                            Dim itemName As String = nameParts(2)

                            Dim Attribute As New TemplateData
                            Attribute.AttributeName = itemName

                            Dim txtColumn = ObtenerTextBox(grupo, itemName, "column").Text.Split(","c)

                            For Each ttext In txtColumn
                                Attribute.Columna.Add(ttext)
                            Next

                            Attribute.NameIDE = ObtenerTextBox(grupo, itemName, "IDE").Text

                            ' Clasificar según el número i
                            Select Case i
                                Case 1
                                    Attribute.Modo = 1

                                Case 2
                                    Attribute.Modo = 2

                                Case 3
                                    Attribute.Modo = 3
                                    Dim txt = ObtenerTextBox(grupo, itemName, 3).Text.Split(","c)

                                    For Each Ttext In txt
                                        Attribute.Attributos.Add(Ttext)
                                    Next

                                Case 4
                                    Attribute.Modo = 4
                                    Dim txt = ObtenerTextBox(grupo, itemName, 4).Text.Split(","c)

                                    For Each Ttext In txt
                                        Attribute.Attributos.Add(Ttext)
                                    Next

                                    txt = ObtenerTextBox(grupo, itemName, "SubNameIDE").Text.Split(","c)

                                    For Each Ttext In txt
                                        Attribute.SubNameIDE.Add(Ttext)
                                    Next

                                    txt = ObtenerTextBox(grupo, nameParts(2), "SubNameIDEColumn").Text.Split(","c)

                                    For Each Ttext In txt
                                        Attribute.SubColumn.Add(Ttext)
                                    Next

                            End Select

                            ' Buscar si ya existe un elemento con el mismo nombre
                            Dim existente = DataList.FirstOrDefault(Function(item) item.AttributeName = itemName)

                            If existente IsNot Nothing Then
                                ' Actualizar valor (ajusta según lo que necesites)
                                existente.Attributos = Attribute.Attributos
                                existente.Columna = Attribute.Columna
                                existente.Modo = Attribute.Modo
                                existente.NameIDE = Attribute.NameIDE
                                existente.Columna = Attribute.Columna
                                existente.NameIDE = Attribute.NameIDE
                                existente.SubNameIDE = Attribute.SubNameIDE
                                existente.SubColumn = Attribute.SubColumn
                            Else
                                DataList.Add(Attribute)
                            End If

                        End If
                    End If
                End If
            End If
        Next

    End Sub

    Private Function ObtenerTextBox(grupo As GroupBox, itemName As String, i As String) As TextBox
        For Each ctrl In grupo.Controls
            If TypeOf ctrl Is TextBox Then
                Dim txt As TextBox = CType(ctrl, TextBox)

                ' Comparar por Tag o Name
                If txt.Tag IsNot Nothing Then
                    If txt.Tag.ToString() = $"txt_{i}_{itemName}" Then
                        Return txt
                    End If
                End If
            End If
        Next

        ' No encontrado
        Return Nothing
    End Function

    Private Sub CargarDatosTemplateData(lista As String, DataList As List(Of TemplateData))
        DataList.Clear() ' Evita duplicados
        Dim Items = lista.Split(";"c)

        For Each Item In Items
            If String.IsNullOrWhiteSpace(Item) Then Continue For

            Dim attr = Item.Split("|"c)
            If attr.Length < 6 Then
                Debug.WriteLine("Dato incompleto: " & Item)
                Continue For
            End If

            Dim Attribute As New TemplateData()

            Attribute.AttributeName = attr(0).Trim()
            Attribute.Modo = Integer.Parse(attr(1).Trim())

            For Each f In attr(2).Split(","c)
                If Not String.IsNullOrWhiteSpace(f) Then
                    Attribute.Attributos.Add(f.Trim())
                End If
            Next

            For Each f In attr(3).Split(","c)
                If Not String.IsNullOrWhiteSpace(f) Then
                    Attribute.Columna.Add(Integer.Parse(f))
                End If
            Next

            Attribute.NameIDE = attr(4).Trim()

            For Each f In attr(5).Split(","c)
                If Not String.IsNullOrWhiteSpace(f) Then
                    Attribute.SubNameIDE.Add(f.Trim())
                End If
            Next

            For Each f In attr(6).Split(","c)
                If Not String.IsNullOrWhiteSpace(f) Then
                    Attribute.SubColumn.Add(f.Trim())
                End If
            Next

            DataList.Add(Attribute)
        Next

        Dim unicos = DataList _
           .GroupBy(Function(x) x.AttributeName & "|" & x.Modo) _
           .Select(Function(g) g.First()) _
           .ToList()

        DataList.Clear()
        DataList.AddRange(unicos)
    End Sub

    Private Sub EliminarListBox(DataList As List(Of TemplateData), Optional ResetX As Boolean = True)
        Dim List As New List(Of String)()

        For Each item In DataList
            List.Add(item.AttributeName)
        Next

        Dim ListaItems As String() = List.ToArray

        Dim miForm As New CustomListBox(ListaItems)
        If miForm.ShowDialog() = DialogResult.OK Then
            Dim textoIngresado As String = miForm.ValorTexto

            If textoIngresado <> "" Then
                For Each Item In DataList
                    If Item.AttributeName = textoIngresado Then
                        DataList.Remove(Item)
                        CargarDatosAnteriores(ResetX)
                        Exit For
                    End If
                Next
            End If
        End If
    End Sub

    Private Sub GenerarPantallasAttributos(TemplateDataList As List(Of TemplateData), TextOfGame As String, NumberOf As Integer, Tab As TabPage, WorkData As aaExcelData.aaTemplateData)
        Dim Lista1 As New List(Of String)
        Dim Lista2 As New List(Of String)
        Dim Lista3 As String = "{"
        Dim Lista4 As String = "{"

        For Each Item In TemplateDataList
            If Item.Modo = 1 Then
                Lista1.Add(Item.AttributeName)
            ElseIf Item.Modo = 2 Then
                Lista2.Add(Item.AttributeName)
            ElseIf Item.Modo = 3 Then
                Lista3 &= $"{Item.AttributeName},"
                Lista3 &= String.Join((","), Item.Attributos)
                Lista3 &= "},"
            ElseIf Item.Modo = 4 Then
                Lista4 &= $"{Item.AttributeName},"
                Lista4 &= String.Join((","), Item.Attributos)
                Lista4 &= "},"
            End If
        Next

        ' Eliminar coma final antes de cerrar la llave
        Lista3 = Lista3.TrimEnd(","c) & "}"
        Lista4 = Lista4.TrimEnd(","c) & "}"

        Dim Lista4Final = ComboBoxSplit(Lista4)

        addControlsToTab(TextOfGame, NumberOf, Tab, Lista1.ToArray, Lista3, Lista2.ToArray, Lista4Final, WorkData)
    End Sub

    Private Function ConventirDatos(List As List(Of TemplateData))
        Dim List1 As New List(Of String)
        Dim Var1 As New List(Of String)
        Dim Var2 As New List(Of String)
        Dim Var3 As New List(Of String)
        Dim Var4 As New List(Of String)
        Dim Var5 As New List(Of String)

        For Each Item In List
            If Item.AttributeName <> "Nombre" Then
                Select Case Item.Modo
                    Case 1
                        Var1.Add(Item.NameIDE)
                    Case 2
                        Var3.Add(Item.NameIDE)
                    Case 3
                        Var2.Add(Item.NameIDE)
                    Case 4
                        Var4.Add(Item.NameIDE)
                        For Each SubName In Item.SubNameIDE
                            Var5.Add(SubName)
                        Next
                End Select
            End If
        Next

        List1.AddRange(Var1)
        List1.AddRange(Var2)
        List1.AddRange(Var3)
        List1.AddRange(Var4)
        List1.AddRange(Var5)

        Return List1
    End Function

    Private Sub TextBox_TextChanged(sender As Object, e As EventArgs)
        Dim txt As TextBox = DirectCast(sender, TextBox)

        If txt.Tag Is Nothing Then Exit Sub

        ' Obtener el contenedor padre (por ejemplo, el GroupBox donde está este TextBox)
        Dim parent As Control = txt.Parent
        If parent Is Nothing Then Exit Sub

        Dim nameParts As String() = txt.Tag.ToString().Split("_"c)

        Dim txtEnabled = ObtenerTextBox(parent, nameParts(2), "SubNameIDE")

        If txt.Text <> "" Then
            txtEnabled.Enabled = True
        Else
            txtEnabled.Enabled = False
        End If

        txtEnabled = ObtenerTextBox(parent, nameParts(2), "SubNameIDEColumn")

        If txt.Text <> "" Then
            txtEnabled.Enabled = True
        Else
            txtEnabled.Enabled = False
        End If
    End Sub

    Private Sub ActualizarPantallas(AddData As Boolean, Optional Data As aaExcelData.aaTemplateData = Nothing)
        Dim WorkData As aaExcelData.aaTemplateData = Nothing

        If AddData Then
            WorkData = Data
        End If

        Dim lista3 As New List(Of String()) From {New String() {}}

        ProcesarDatos(GPDiscreteTemplatesConfig, DiscreteTemplateData)
        GenerarPantallasAttributos(DiscreteTemplateData, "Disc", NUPdfa.Value, TPDiscFA, WorkData)
        ProcesarDatos(GBAnalogTemplateConfig, AnalogTemplateData)
        GenerarPantallasAttributos(AnalogTemplateData, "Analog", NUPafa.Value, TPFAAnalog, WorkData)

        addControlsToTab("UDA", NUPuda.Value, TPUDA, {"Nombre", "Valor"}, "{Prueba,Prueba1,Prueba2},{Prueba3,Prueba4}", {}, lista3, WorkData)
        addControlsToTab("Scr", NUPscr.Value, TPScript, {"Nombre", "Expresion", "Execution Text"}, "{Trigger Type,WhileTrue,WhileFalse,OnTrue,OnFalse,DataChange,Periodic}", {}, lista3, WorkData)

        If WorkData IsNot Nothing Then
            txtNewTName.Text = WorkData.TemplateName
        End If
    End Sub
End Class

' ---------------- Clase que representa datos de una plantilla ----------------
<Serializable()>
Public Class TemplateData
    Public AttributeName As String
    Public Modo As String
    Public Attributos As New List(Of String)
    Public Columna As New List(Of Integer)
    Public NameIDE As String
    Public SubNameIDE As New List(Of String)
    Public SubColumn As New List(Of Integer)

    ' Constructor vacío requerido por serialización
    Public Sub New()
    End Sub

    ' Constructor parametrizado para inicializar todos los campos
    Public Sub New(ByVal AttributeName As String, Modo As String, Attributos As List(Of String), Columna As List(Of Integer), NameIDE As String, SubNameIDE As List(Of String), SubColumn As List(Of Integer))
        Me.AttributeName = AttributeName
        Me.Modo = Modo
        Me.Attributos = Attributos
        Me.Columna = Columna
        Me.NameIDE = NameIDE
        Me.SubNameIDE = SubNameIDE
        Me.SubColumn = SubColumn
    End Sub
End Class