Imports System.IO
Imports System.Text
Imports System.Xml.Serialization
Imports System.Xml.Linq

Partial Public Class frmMain
    Public aaGalaxyTools As aaGalaxyTools
    Public aaExcelData As aaExcelData
    Public aaTemplateDiscAttrList As New List(Of aaTemplateData.aaNewField)()
    Public aaTemplateAnalogAttrList As New List(Of aaTemplateData.aaNewField)()
    Public aaTemplateScriptList As New List(Of aaTemplateData.aaNewField)()
    Public aaTemplateUDAList As New List(Of aaTemplateData.aaNewField)()
    Public DiscreteTemplateData As New List(Of TemplateData)()
    Public AnalogTemplateData As New List(Of TemplateData)()
    Public ScriptTemplateData As New List(Of TemplateData)()

    Public AuthenticationMode As String
    Public ExportFolder As String
    Public ImportExcel As String
    Public LoneAttributes As String() = {}
    Public ArrayAttributes As String() = {}
    Public LoneAttributesIndex As String() = {}
    Public ArrayAttributesIndex As String() = {}
    Public xOffsetsPorGrupo As New Dictionary(Of GroupBox, Integer)


    ' Discrete
    Public DiscreteString8 As String() = {}

    ' Analogic
    Public AnalogicString8 As String() = {}

    ' UDA
    Public UDAInteger1 As Integer() = {}

    ' Script
    Public ScriptInteger1 As Integer() = {}

    Public LogBox As ListBox

    ' ---------------- Constructor ----------------
    Public Sub New()
        InitializeComponent()

        lblStatus.Text = "Initializing"

        ' Restaura la configuración de datos guardados
        restore_configdata(False)

        ' Inicializa los atributos y pestañas de plantillas
        AñadirAttributos()
        RefresTemplateTabs()

        ' Inicializa herramientas auxiliares
        aaGalaxyTools = New aaGalaxyTools
        aaExcelData = New aaExcelData

        ' Establece nombre del nodo predeterminado
        txtNodeName.Text = Environment.MachineName

        ' Asigna el ListBox de errores
        LogBox = ErrorLog

        ' Obtiene la lista de galaxias disponibles desde el nodo local
        cmboGalaxyList.DataSource = aaGalaxyTools.getGalaxies(txtNodeName.Text)

        For Each item In cmboGalaxyList.Items
            If item.ToString = My.Settings.LastGalaxy Then
                cmboGalaxyList.SelectedItem = item
                Exit For
            End If
        Next

        ' Actualiza la información de la galaxia seleccionada
        refreshGalaxyInfo(cmboGalaxyList.SelectedValue)

        ' Refresca instancias y atributos de plantillas
        Refres_Instances(1)

        ActualizarPantallas(False)

        ' Carga datos previos si existen
        CargarDatosAnteriores()

        btnExport.Enabled = False
        lblStatus.Text = ""
    End Sub

    Private Sub frmMain_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If MessageBox.Show("¿Deseas guardar los cambios?", "Confirmar", MessageBoxButtons.YesNo) = DialogResult.Yes Then
            Dim valoresTextBox As New List(Of String)

            ' Recolectar los valores de los TextBoxs
            For Each ctrl As Control In EditAttributes.Controls
                If TypeOf ctrl Is TextBox Then
                    Dim txt As TextBox = CType(ctrl, TextBox)
                    valoresTextBox.Add(txt.Text)
                End If
            Next

            ' Verifica que la cantidad de valores coincida con el total de atributos
            If valoresTextBox.Count <> (LoneAttributes.Length + ArrayAttributes.Length) Then
                MessageBox.Show("El número de valores no coincide con los atributos.")
                Return
            End If

            ' Construir las listas nuevas con formato "NombreAtributo|valor"
            Dim nuevosLoneAttributes As New List(Of String)
            Dim nuevosArrayAttributes As New List(Of String)

            For i As Integer = 0 To LoneAttributes.Length - 1
                nuevosLoneAttributes.Add(LoneAttributes(i) & "|" & valoresTextBox(i))
            Next

            For i As Integer = 0 To ArrayAttributes.Length - 1
                nuevosArrayAttributes.Add(ArrayAttributes(i) & "|" & valoresTextBox(i + LoneAttributes.Length))
            Next

            ' Guardar en Settings
            My.Settings.AloneAttributes = String.Join(",", nuevosLoneAttributes)
            My.Settings.ArrayAttributes = String.Join(",", nuevosArrayAttributes)
            My.Settings.CSVDir = txtDirecc.Text
            My.Settings.CSVTAGS = txtTags.Text
            My.Settings.InstanceNameIndex = txtIname.Text
            My.Settings.InstanceTemplateIndex = txtItemplate.Text
            My.Settings.InstanceEngUnits1 = txtIEngUnits.Text
            My.Settings.InstanceEngUnits2 = txtMapIndex.Text
            My.Settings.ScriptInt1 = txtScript1.Text
            My.Settings.UDAInt1 = txtUDA1.Text
            My.Settings.DiscreteLockedAttr = DiscreteBlockAttr.Text
            My.Settings.AnalogLockedAttr = AnalogBlockAttr.Text
            My.Settings.LastGalaxy = cmboGalaxyList.SelectedItem.ToString

            Dim Strin As String = ""
            For Each item In DiscreteTemplateData
                Strin &= $"{item.AttributeName}|{item.Modo}|{String.Join(",", item.Attributos)}|{String.Join(",", item.Columna)}|{item.NameIDE}|{String.Join(",", item.SubNameIDE)}|{String.Join(",", item.SubColumn)};"
            Next

            My.Settings.SettingPrueba = Strin

            Strin = ""
            For Each item In AnalogTemplateData
                Strin &= $"{item.AttributeName}|{item.Modo}|{String.Join(",", item.Attributos)}|{String.Join(",", item.Columna)}|{item.NameIDE}|{String.Join(",", item.SubNameIDE)}|{String.Join(",", item.SubColumn)};"
            Next

            My.Settings.SettingPrueba1 = Strin
            My.Settings.Save()
        End If
    End Sub

    Private Sub restore_configdata(RestoreData As Boolean)
        GenerarDatosExcel(RestoreData)

        txtItemplate.Text = My.Settings.InstanceTemplateIndex
        txtTags.Text = My.Settings.CSVTAGS
        txtIname.Text = My.Settings.InstanceNameIndex
        txtDirecc.Text = My.Settings.CSVDir
        txtIEngUnits.Text = My.Settings.InstanceEngUnits1
        txtMapIndex.Text = My.Settings.InstanceEngUnits2
        txtUDA1.Text = My.Settings.UDAInt1
        txtScript1.Text = My.Settings.ScriptInt1
        DiscreteBlockAttr.Text = My.Settings.DiscreteLockedAttr
        AnalogBlockAttr.Text = My.Settings.AnalogLockedAttr

        CargarDatosTemplateData(My.Settings.SettingPrueba, DiscreteTemplateData)
        CargarDatosTemplateData(My.Settings.SettingPrueba1, AnalogTemplateData)
    End Sub

    Private Sub CambiarTamaño(sender As Object, e As EventArgs) Handles TPInstancesData.SelectedIndexChanged
        ' Verifica si la pestaña activa en TPInstancesData es "TPConfigurations"
        If TPInstancesData.SelectedTab IsNot Nothing AndAlso TPInstancesData.SelectedTab.Name = "TPConfigurations" Then
            Me.Size = New Size(790, 527)
        Else
            Me.Size = New Size(419, 643)
        End If
    End Sub

    Private Sub CambiarTamaño1(sender As Object, e As EventArgs)
        Me.Size = New Size(419, 643)
    End Sub

    Private Sub CambiarTamaño2(sender As Object, e As EventArgs) Handles TPTemplates.SelectedIndexChanged
        ' Verifica si la pestaña activa en TPInstancesData es "TPConfigurations"
        If TPTemplates.SelectedTab IsNot Nothing AndAlso TPTemplates.SelectedTab.Name = "TabPage6" Then
            Me.Size = New Size(1185, 614)
        Else
            Me.Size = New Size(419, 643)
        End If
    End Sub


    'Esta parte ocupa la opcion de importar y exportar graficos en XML
    Private Sub btnExpGraphic_Click(sender As Object, e As EventArgs) Handles btnIxpGraphic.Click
        If txtExpGraphic.Text <> "" Then
            aaGalaxyTools.ExportSymbol(txtExpGraphic.Text)
        Else
            MsgBox("Por favor introduzca el nombre del grafico deseado")
        End If
    End Sub

    Private Sub btnImpGraphic_Click(sender As Object, e As EventArgs) Handles btnImpGraphic.Click
        Dim ImportGraphic As String
        Try
            If txtImpGraphic.Text <> "" Then

                Dim openFileDialog As New OpenFileDialog()
                openFileDialog.Filter = "Archivo XML (*.xml)|*.xml"
                openFileDialog.Title = "Seleccionar un archivo"

                If openFileDialog.ShowDialog() = DialogResult.OK Then
                    ImportGraphic = openFileDialog.FileName ' Se guarda la ruta del archivo

                    aaGalaxyTools.ImportSymbol(txtImpGraphic.Text, ImportGraphic, chbOverWrite.Checked)
                Else
                    MsgBox("Por favor seleccione un archivo tipo .xml")
                End If
            Else
                MsgBox("Por favor introduzca el nombre del grafico con el que se importara")
            End If
        Catch ex As Exception

        End Try
    End Sub
End Class