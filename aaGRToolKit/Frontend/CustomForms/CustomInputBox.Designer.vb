<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class CustomInputBox
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
		Me.TextBox1 = New System.Windows.Forms.TextBox()
		Me.CheckBox1 = New System.Windows.Forms.CheckBox()
		Me.btnAceptar = New System.Windows.Forms.Button()
		Me.btnCancelar = New System.Windows.Forms.Button()
		Me.SuspendLayout()
		'
		'TextBox1
		'
		Me.TextBox1.Location = New System.Drawing.Point(13, 13)
		Me.TextBox1.Name = "TextBox1"
		Me.TextBox1.Size = New System.Drawing.Size(232, 20)
		Me.TextBox1.TabIndex = 0
		'
		'CheckBox1
		'
		Me.CheckBox1.AutoSize = True
		Me.CheckBox1.Location = New System.Drawing.Point(13, 43)
		Me.CheckBox1.Name = "CheckBox1"
		Me.CheckBox1.Size = New System.Drawing.Size(50, 17)
		Me.CheckBox1.TabIndex = 1
		Me.CheckBox1.Text = "Array"
		Me.CheckBox1.UseVisualStyleBackColor = True
		'
		'btnAceptar
		'
		Me.btnAceptar.Location = New System.Drawing.Point(111, 74)
		Me.btnAceptar.Name = "btnAceptar"
		Me.btnAceptar.Size = New System.Drawing.Size(64, 22)
		Me.btnAceptar.TabIndex = 2
		Me.btnAceptar.Text = "Aceptar"
		Me.btnAceptar.UseVisualStyleBackColor = True
		'
		'btnCancelar
		'
		Me.btnCancelar.Location = New System.Drawing.Point(180, 74)
		Me.btnCancelar.Name = "btnCancelar"
		Me.btnCancelar.Size = New System.Drawing.Size(64, 22)
		Me.btnCancelar.TabIndex = 3
		Me.btnCancelar.Text = "Cancelar"
		Me.btnCancelar.UseVisualStyleBackColor = True
		'
		'CustomInputBox
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(257, 108)
		Me.Controls.Add(Me.btnCancelar)
		Me.Controls.Add(Me.btnAceptar)
		Me.Controls.Add(Me.CheckBox1)
		Me.Controls.Add(Me.TextBox1)
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.Name = "CustomInputBox"
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
		Me.Text = "Ingrese Nuevo Atributo"
		Me.ResumeLayout(False)
		Me.PerformLayout()

	End Sub

	Friend WithEvents TextBox1 As TextBox
    Friend WithEvents CheckBox1 As CheckBox
    Friend WithEvents btnAceptar As Button
    Friend WithEvents btnCancelar As Button
End Class