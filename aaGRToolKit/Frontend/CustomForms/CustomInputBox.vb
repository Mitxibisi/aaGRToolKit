Public Class CustomInputBox
    Public Property ValorTexto As String
    Public Property OpcionMarcada As Boolean

    Private Sub btnAceptar_Click(sender As Object, e As EventArgs) Handles btnAceptar.Click
        ValorTexto = TextBox1.Text
        OpcionMarcada = CheckBox1.Checked
        DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Private Sub btnCancelar_Click(sender As Object, e As EventArgs) Handles btnCancelar.Click
        DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub
End Class