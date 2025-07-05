Imports System.IO
Imports System.Text
Imports System.Xml.Serialization
Imports System.Xml.Linq

Partial Public Class frmMain
    ' ------- Pantalla Exportar Plantillas -------
    Private Sub linkLblSelectAll_Click(sender As Object, e As EventArgs) Handles linkLblSelectAll.Click
        For x = 1 To (lstTemplates.Items.Count - 1)
            lstTemplates.SetSelected(x, True)
        Next x
    End Sub

    Private Sub linkLblSelectNone_Click(sender As Object, e As EventArgs) Handles linkLblSelectNone.Click
        For x = 1 To (lstTemplates.Items.Count - 1)
            lstTemplates.SetSelected(x, False)
        Next x
    End Sub

    Private Sub btnBrowseFolders_Click(sender As Object, e As EventArgs) Handles btnBrowseFolders.Click
        If dlgFolderBrowser.ShowDialog() = DialogResult.OK Then
            ExportFolder = dlgFolderBrowser.SelectedPath
            lblFolderPath.Text = ExportFolder

            If lstTemplates.SelectedItems.Count > 0 And ExportFolder IsNot Nothing And aaGalaxyTools.loggedIn Then
                btnExport.Enabled = True
            Else
                btnExport.Enabled = False
            End If
        End If
    End Sub

    Private Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click

        Dim y As Integer
        y = lstTemplates.SelectedItems.Count - 1
        Dim TemplateNames(0 To y) As String

        Dim x As Integer = 0
        For Each Template In lstTemplates.SelectedItems
            TemplateNames(x) = Template
            x += 1
        Next

        lblStatus.Text = "Exporting " & lstTemplates.SelectedItems.Count.ToString & " Templates"

        ExportTemplatesToFile(ExportFolder, TemplateNames, ProgressBar1)

        lblStatus.Text = "Archivo CSV generado"

    End Sub

    Private Sub btnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click
        lblStatus.Text = "Logging in"

        refreshGalaxyInfo(cmboGalaxyList.SelectedValue)

        If aaGalaxyTools.login(txtUserInput.Text, txtPwdInput.Text) >= 0 Then
            lblStatus.Text = "Logged In " & cmboGalaxyList.Text & " Wich User " & txtUserInput.Text
        Else
            lblStatus.Text = "Error logging in"
        End If

    End Sub

    Private Sub btnRefreshTemplates_Click(sender As Object, e As EventArgs) Handles btnRefreshTemplates.Click
        lblStatus.Text = "Refreshing Template List"
        lstTemplates.DataSource = aaGalaxyTools.getTemplates(chkHideBaseTemplates.CheckState)
        lblStatus.Text = ""
    End Sub

    Private Sub btnRefreshGalaxies_Click(sender As Object, e As EventArgs) Handles btnRefreshGalaxies.Click
        lblStatus.Text = "Refreshing Available Galaxies"
        clearGalaxyInfo()

        ' get the list of Galaxies from the local node and fill the combo box with the collection
        cmboGalaxyList.DataSource = aaGalaxyTools.getGalaxies(txtNodeName.Text)

        ' do any UI clean up work that might be needed when the galaxy name changes
        refreshGalaxyInfo(cmboGalaxyList.SelectedValue)
        lblStatus.Text = ""
    End Sub

    ' ------- Pantalla Exportar Plantillas Funciones -------
    Private Sub cmboGalaxyList_SelectionChangeCommitted(ByVal sender As Object, ByVal e As EventArgs) Handles cmboGalaxyList.SelectionChangeCommitted
        Dim senderComboBox As ComboBox = CType(sender, ComboBox)

        If (senderComboBox.SelectedValue IsNot Nothing) Then
            refreshGalaxyInfo(senderComboBox.SelectedValue)
        Else
            clearGalaxyInfo()
        End If
    End Sub

    Public Sub refreshGalaxyInfo(ByVal GalaxyName)
        lblStatus.Text = "Refreshing Galaxy Info"
        aaGalaxyTools.setGalaxy(GalaxyName)
        grpLoginInput.Visible = aaGalaxyTools.showLogin

        If aaGalaxyTools.loggedIn Then
            ' this would happen if there's no security and we automatically login
            lstTemplates.DataSource = aaGalaxyTools.getTemplates()
        Else
            lstTemplates.Items.Clear()
        End If
        lblStatus.Text = ""
    End Sub

    Public Sub clearGalaxyInfo()
        lstTemplates.DataSource = Nothing
        txtUserInput.Clear()
        txtPwdInput.Clear()
        btnExport.Enabled = False
    End Sub

    Public Sub ExportTemplatesToFile(ByVal filePath As String, ByVal templateNames As String(), ByVal progressBar As ProgressBar)
        Dim outputFile As String = Path.Combine(filePath, "atributos_extraidos.csv")

        Try
            ' Configurar la barra de progreso
            progressBar.Minimum = 0
            progressBar.Maximum = templateNames.Length
            progressBar.Value = 0
            progressBar.Step = 1

            ' Usar StreamWriter para mejor rendimiento en escritura de archivos
            Using writer As New StreamWriter(outputFile, False, Encoding.UTF8)
                writer.WriteLine("Plantilla,Nombre,Plantilla derivada,Descripción,Historizado,Eventos,Alarm,Unidad") ' Encabezado CSV

                For Each templateName In templateNames
                    Dim templateData = aaGalaxyTools.getTemplateData(templateName)

                    ' Procesar atributos discretos
                    For Each attr As aaFieldAttributeDiscrete In templateData.FieldAttributesDiscrete
                        writer.WriteLine($"{templateName},{attr.Name},{attr.TemplateName},{attr.Description},{attr.Historized},{attr.Events},{attr.Alarm},{attr.EngUnits}")
                    Next

                    ' Procesar atributos analógicos
                    For Each attr As aaFieldAttributeAnalog In templateData.FieldAttributesAnalog
                        writer.WriteLine($"{templateName},{attr.Name},{attr.TemplateName},{attr.Description},{attr.Historized},{attr.Events},{attr.Alarm},{attr.EngUnits}")
                    Next

                    ' Incrementar la barra de progreso
                    progressBar.PerformStep()
                Next
            End Using

            ' Restablecer barra de progreso
            progressBar.Value = 0

            MessageBox.Show($"Exportación completada: {templateNames.Length} plantilla(s) exportada(s).", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As IOException
            MessageBox.Show($"Error de archivo: {ex.Message}", "Error de E/S", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Catch ex As UnauthorizedAccessException
            MessageBox.Show("Acceso denegado. Verifique los permisos del archivo.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)

        Catch ex As Exception
            MessageBox.Show($"Error inesperado: {ex.Message}{vbCrLf}{ex.StackTrace}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class