Public Class CustomListBox

    Public Property ValorTexto As String
    Public Property OpcionMarcada As Boolean

    Public Sub New(ItemsNames As String())

        ' Esta llamada es exigida por el diseñador.
        InitializeComponent()

        ItemsList(ItemsNames)
        ' Agregue cualquier inicialización después de la llamada a InitializeComponent().

    End Sub

    Public Sub ItemsList(ItemNames As String())
        For Each item In ItemNames
            ListBox1.Items.Add(item)
        Next
    End Sub

    Public Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ValorTexto = ListBox1.SelectedItems(0)
        DialogResult = DialogResult.OK
        Me.Close()
    End Sub
End Class