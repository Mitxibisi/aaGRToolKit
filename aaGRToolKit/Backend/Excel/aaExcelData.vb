Imports System
Imports System.IO
Imports OfficeOpenXml

Public Class aaExcelData
    <Serializable()>
    Public Class aaInstanceData
        Public InstanceName As String
        Public InstanceTemplate As String
        Public InstanceAloneAttr As New List(Of String)
        Public InstanceArrayAttr As New List(Of List(Of String))
        Public EngUnitsArray As New List(Of String)

        Public Sub New()

        End Sub

        Public Sub New(ByVal InstanceName As String, ByVal InstanceTemplate As String, InstanceArrayAttr As List(Of List(Of String)), InstanceAloneAttr As List(Of String), EngsUnitsArray As List(Of String))

            Me.InstanceName = InstanceName
            Me.InstanceTemplate = InstanceTemplate
            Me.InstanceArrayAttr = InstanceArrayAttr
            Me.InstanceAloneAttr = InstanceAloneAttr
            Me.EngUnitsArray = EngsUnitsArray

        End Sub
    End Class

    <Serializable()>
    Public Class aaTemplateDiscData
        Public AttrName As String
        Public AccessMode As String
        Public TextBoxAttr As New List(Of String)
        Public CheckBoxValue As New List(Of String)
        Public GroupBoxAttr As New List(Of String)

        Public Sub New()

        End Sub

        Public Sub New(AttrName As String, AccesMode As String, TextBoxAttr As List(Of String), CheckBoxValueArray As List(Of String), GroupBoxAttr As List(Of String))

            Me.TextBoxAttr = TextBoxAttr
            Me.AccessMode = AccesMode
            Me.CheckBoxValue = CheckBoxValueArray
            Me.GroupBoxAttr = GroupBoxAttr
            Me.AttrName = AttrName
        End Sub
    End Class

    <Serializable()>
    Public Class aaTemplateAnaData
        Public AttrName As String
        Public AccessMode As String
        Public TextBoxAttr As New List(Of String)
        Public CheckBoxValue As New List(Of String)
        Public GroupBoxAttr As New List(Of String)

        Public Sub New()

        End Sub

        Public Sub New(AttrName As String, AccesMode As String, TextBoxAttr As List(Of String), CheckBoxValueArray As List(Of String), GroupBoxAttr As List(Of String))

            Me.TextBoxAttr = TextBoxAttr
            Me.AccessMode = AccesMode
            Me.CheckBoxValue = CheckBoxValueArray
            Me.GroupBoxAttr = GroupBoxAttr
            Me.AttrName = AttrName
        End Sub
    End Class

    <Serializable()>
    Public Class aaTemplateScriptData
        Public Name As String
        Public Expression As String
        Public TriggerType As String
        Public ExecuteText As String

        Public Sub New(ByVal Name As String, ByVal Expression As String, ByVal ExecuteText As String, ByVal TriggerType As String)

            Me.Name = Name
            Me.Expression = Expression
            Me.TriggerType = TriggerType
            Me.ExecuteText = ExecuteText

        End Sub

    End Class

    <Serializable()>
    Public Class aaTemplateUDAData
        Public Name As String
        Public Value As String

        Public Sub New(ByVal Name As String, ByVal Value As String)

            Me.Name = Name
            Me.Value = Value

        End Sub

    End Class

    <Serializable()>
    Public Class aaTemplateData
        Public TemplateName As String
        Public Discrete As List(Of aaTemplateDiscData)
        Public Analogic As List(Of aaTemplateAnaData)
        Public Script As List(Of aaTemplateScriptData)
        Public UDA As List(Of aaTemplateUDAData)

        Public Sub New()

        End Sub

        Public Sub New(TemplateName As String, Discrete As List(Of aaTemplateDiscData), Analogic As List(Of aaTemplateAnaData), Script As List(Of aaTemplateScriptData), UDA As List(Of aaTemplateUDAData))

            Me.TemplateName = TemplateName
            Me.Discrete = Discrete
            Me.Analogic = Analogic
            Me.Script = Script
            Me.UDA = UDA
        End Sub
    End Class

    Public Function CargarDatosPlantilla(FilPath As String,
                                         DiscreteList1 As Integer(), DiscreteList2 As Integer(), DiscreteList3 As Integer(), DiscreteList4 As Integer(),
                                         AnalogList1 As Integer(), AnalogList2 As Integer(), AnalogList3 As Integer(), AnalogList4 As Integer(),
                                         UDAList1 As Integer(),
                                         ScriptList1 As Integer())
        Dim filePath As String
        Dim package As ExcelPackage
        Dim worksheet As ExcelWorksheet
        Dim InstanceDiscData As New List(Of aaTemplateDiscData)
        Dim InstanceAnaData As New List(Of aaTemplateAnaData)
        Dim InstanceScriptData As New List(Of aaTemplateScriptData)
        Dim InstanceUDAData As New List(Of aaTemplateUDAData)
        Dim AttrNames As New List(Of String)()
        Dim TemplateData As New aaTemplateData
        Try
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial
            filePath = FilPath ' Archivo con macros
            package = New ExcelPackage(New FileInfo(filePath))
            worksheet = package.Workbook.Worksheets(0)

            AttrNames = ObtenerDatosName(2, worksheet)

            For Each name In AttrNames
                Dim IsTemplate = ObtenerDatosInstance(2, worksheet, name.ToString, 1)
                If IsTemplate.StartsWith("$") Then
                    Dim TextBoxValues As List(Of String) = ObtenerDatosFila(2, worksheet, name.ToString, DiscreteList1)
                    Dim ComboBoxValues As List(Of String) = ObtenerDatosFila(2, worksheet, name.ToString, DiscreteList2)
                    Dim CheckBoxValues As List(Of String) = ObtenerDatosFila(2, worksheet, name.ToString, DiscreteList3)
                    Dim GroupBoxValues As List(Of String) = ObtenerDatosFila(2, worksheet, name.ToString, DiscreteList4)

                    Dim NewInstanceData = New aaTemplateDiscData(name.ToString, ComboBoxValues(0), TextBoxValues, CheckBoxValues, GroupBoxValues)

                    InstanceDiscData.Add(NewInstanceData)
                End If
            Next

            worksheet = package.Workbook.Worksheets(1)

            AttrNames = ObtenerDatosName(2, worksheet)

            For Each name In AttrNames
                Dim IsTemplate = ObtenerDatosInstance(2, worksheet, name.ToString, 1)
                If IsTemplate.StartsWith("$") Then
                    Dim TextBoxValues As List(Of String) = ObtenerDatosFila(2, worksheet, name.ToString, AnalogList1)
                    Dim ComboBoxValues As List(Of String) = ObtenerDatosFila(2, worksheet, name.ToString, AnalogList2)
                    Dim CheckBoxValues As List(Of String) = ObtenerDatosFila(2, worksheet, name.ToString, AnalogList3)
                    Dim GroupBoxValues As List(Of String) = ObtenerDatosFila(2, worksheet, name.ToString, AnalogList4)

                    Dim NewInstanceData = New aaTemplateAnaData(name.ToString, ComboBoxValues(0), TextBoxValues, CheckBoxValues, GroupBoxValues)

                    InstanceAnaData.Add(NewInstanceData)
                End If
            Next

            worksheet = package.Workbook.Worksheets(2)

            AttrNames = ObtenerDatosName(2, worksheet)

            For Each name In AttrNames
                Dim IsTemplate = ObtenerDatosInstance(2, worksheet, name.ToString, 1)
                If IsTemplate.StartsWith("$") Then
                    Dim TextBoxValues As List(Of String) = ObtenerDatosFila(2, worksheet, name.ToString, ScriptList1)

                    Dim NewInstanceData = New aaTemplateScriptData(name.ToString, TextBoxValues(0), TextBoxValues(1), TextBoxValues(2))
                    InstanceScriptData.Add(NewInstanceData)
                End If
            Next

            worksheet = package.Workbook.Worksheets(3)

            AttrNames = ObtenerDatosName(2, worksheet)

            For Each name In AttrNames
                Dim IsTemplate = ObtenerDatosInstance(2, worksheet, name.ToString, 1)
                If IsTemplate.StartsWith("$") Then
                    Dim TextBoxValues As List(Of String) = ObtenerDatosFila(2, worksheet, name.ToString, UDAList1)

                    Dim NewInstanceData = New aaTemplateUDAData(name.ToString, TextBoxValues(0))

                    InstanceUDAData.Add(NewInstanceData)
                End If
            Next

            worksheet = package.Workbook.Worksheets(0)
            Dim TemplateName = ObtenerDatosName(1, worksheet)
            TemplateData = New aaTemplateData(TemplateName(1), InstanceDiscData, InstanceAnaData, InstanceScriptData, InstanceUDAData)
        Catch e As Exception
            Console.WriteLine(e)
        End Try
        Return TemplateData
    End Function

    Public Function CargarDatosMapeado(FilPath As String, InstanceNameTitle As String, InstanceTemplateTitle As String, EditableAloneElementsIndex As List(Of String), EditableArrayElementsIndex As List(Of String), EngUnitsTitle As String, MapIndexTitle As String)
        Dim filePath As String
        Dim package As ExcelPackage
        Dim worksheet As ExcelWorksheet
        Dim InstanceData As New List(Of aaInstanceData)
        Dim InstancesNames As New List(Of String)()
        Dim NewInstanceData As New aaInstanceData

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial
        filePath = FilPath ' Archivo con macros
        package = New ExcelPackage(New FileInfo(filePath))

        worksheet = package.Workbook.Worksheets(0)

        If worksheet.Tables.Count = 0 Then
            ' Si hay una segunda hoja, la probamos
            If package.Workbook.Worksheets.Count > 1 Then
                worksheet = package.Workbook.Worksheets(1)
            End If

            ' Si aún así no hay tablas, error
            If worksheet.Tables.Count = 0 Then
                MsgBox("El Excel otorgado no está correctamente configurado (no se encontró ninguna tabla).")
                Return InstanceData
            End If
        End If


        Dim InstanceNameColumnIndex As Integer = ObtenerIndicePorCabecera(worksheet, InstanceNameTitle)
        Dim InstanceTemplateColumnIndex As Integer = ObtenerIndicePorCabecera(worksheet, InstanceTemplateTitle)
        Dim EngUnits As Integer = ObtenerIndicePorCabecera(worksheet, EngUnitsTitle)
        Dim MapIndex As Integer = ObtenerIndicePorCabecera(worksheet, MapIndexTitle)

        InstancesNames = ObtenerDatosName(InstanceNameColumnIndex, worksheet)

        If InstancesNames.Count <= 0 Then
            MessageBox.Show("No se encontró ninguna instancia en la primera hoja. Verifica que el archivo contenga datos válidos y que la hoja esté en la posición correcta.")
            Return InstanceData
        End If

        For Each name In InstancesNames
            Dim strname As String = name.ToString
            Dim Template = ObtenerDatosInstance(InstanceNameColumnIndex, worksheet, strname, InstanceTemplateColumnIndex)

            Dim NewArrayElementList As New List(Of List(Of String))

            For Each column As String In EditableArrayElementsIndex
                Dim attrList = ObtenerDatosColumnaPorCabecera(InstanceNameColumnIndex, worksheet, strname, column)

                NewArrayElementList.Add(attrList)
            Next

            Dim NewAloneElementList As New List(Of String)

            For Each column As String In EditableAloneElementsIndex
                Dim attrList = ObtenerDatosPorCabecera(InstanceNameColumnIndex, worksheet, strname, column)

                NewAloneElementList.Add(attrList)
            Next

            Dim EngUnitsValue = ObtenerDatosColumnaEngUnits(InstanceNameColumnIndex, worksheet, strname, EngUnits)
            Dim NewMapList = ObtenerDatosColumna(InstanceNameColumnIndex, worksheet, strname, MapIndex)
            Dim CombinedList As New List(Of String)()

            For i As Integer = 0 To Math.Min(EngUnitsValue.Count, NewMapList.Count) - 1
                CombinedList.Add(NewMapList(i) & " | " & EngUnitsValue(i))
            Next

            If Template.StartsWith("$") Then
                NewInstanceData = New aaInstanceData(strname,
                                             Template,
                                             NewArrayElementList,
                                             NewAloneElementList,
                                             CombinedList)
                InstanceData.Add(NewInstanceData)
            End If
        Next

        If InstanceData.Count <= 0 Then
            MessageBox.Show("No se ha podido registrar ningún dato.")
        End If

        Return InstanceData
    End Function

    Public Sub GenerarCsv(FilPath As String, InstanceNameTitle As String, InstanceTemplateTitle As String, csvIndexTitle As String, csv1indexTitle As String, generateCSV As Boolean, generateDoubleCSV As Boolean, DoubleText1 As String, DoubleText2 As String)
        Dim filePath As String
        Dim package As ExcelPackage
        Dim worksheet As ExcelWorksheet
        Dim InstancesNames As New List(Of String)()
        Dim CSVBase As New List(Of String)()

        Try
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial
            filePath = FilPath ' Archivo con macros
            package = New ExcelPackage(New FileInfo(filePath))

            worksheet = package.Workbook.Worksheets(0)

            If worksheet.Tables.Count = 0 Then
                ' Si hay una segunda hoja, la probamos
                If package.Workbook.Worksheets.Count > 1 Then
                    worksheet = package.Workbook.Worksheets(1)
                End If

                ' Si aún así no hay tablas, error
                If worksheet.Tables.Count = 0 Then
                    MsgBox("El Excel otorgado no está correctamente configurado (no se encontró ninguna tabla).")
                    Exit Sub
                End If
            End If

            Dim InstanceNameColumnIndex As Integer = ObtenerIndicePorCabecera(worksheet, InstanceNameTitle)
            Dim InstanceTemplateColumnIndex As Integer = ObtenerIndicePorCabecera(worksheet, InstanceTemplateTitle)
            Dim csvIndex As Integer = ObtenerIndicePorCabecera(worksheet, csvIndexTitle)
            Dim csv1index As Integer = ObtenerIndicePorCabecera(worksheet, csv1indexTitle)

            InstancesNames = ObtenerDatosName(InstanceNameColumnIndex, worksheet)

            For Each name In InstancesNames
                If generateCSV Then
                    Dim Tags = ObtenerDatosCSV(InstanceNameColumnIndex, worksheet, csvIndex, csv1index, InstanceTemplateColumnIndex)

                    For I As Integer = 0 To Tags.Count - 1
                        Dim original As String = Tags(I)
                        ' Aplicar las sustituciones como en la fórmula de Excel
                        original = Tags(I).Replace(".DBX", ",X") _
                                       .Replace(".DBDINT", ",DINT") _
                                       .Replace(".DBW", ",INT") _
                                       .Replace(".DBB", ",BYTE") _
                                       .Replace(".DBD", ",REAL")

                        CSVBase.Add(original)
                    Next
                End If
            Next

            If generateCSV Then
                Dim CSV As List(Of String) = CSVBase.Distinct().ToList()
                If generateDoubleCSV Then
                    If generateCSV Then
                        ' Obtener la ruta del escritorio del usuario actual
                        Dim desktopPath As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)

                        ' Inicializar listas para cada tipo de DB
                        Dim db100List As New List(Of String)()
                        Dim db101List As New List(Of String)()

                        ' Clasificar las líneas según el tipo de DB
                        For Each linea As String In CSV
                            If linea.Contains(DoubleText1) Then
                                db100List.Add(linea)
                            ElseIf linea.Contains(DoubleText2) Then
                                db101List.Add(linea)
                            End If
                        Next

                        ' Guardar DB100 en su archivo
                        Dim rutaCSV_DB100 As String = Path.Combine(desktopPath, "resultado_" & DoubleText1 & ".csv")
                        Using writer As New StreamWriter(rutaCSV_DB100, False, System.Text.Encoding.UTF8)
                            For Each linea As String In db100List
                                writer.WriteLine(linea)
                            Next
                        End Using

                        ' Guardar DB101 en su archivo
                        Dim rutaCSV_DB101 As String = Path.Combine(desktopPath, "resultado_" & DoubleText2 & ".csv")
                        Using writer As New StreamWriter(rutaCSV_DB101, False, System.Text.Encoding.UTF8)
                            For Each linea As String In db101List
                                writer.WriteLine(linea)
                            Next
                        End Using

                        ' Mensaje de confirmación
                        MessageBox.Show("CSV generado correctamente en: " & vbCrLf &
                        rutaCSV_DB100 & vbCrLf & rutaCSV_DB101)
                    End If

                Else
                    ' Obtener la ruta del escritorio del usuario actual
                    Dim desktopPath As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                    Dim rutaCSV As String = Path.Combine(desktopPath, "resultado.csv")

                    ' Guardar la lista CSV en el archivo
                    Using writer As New StreamWriter(rutaCSV, False, System.Text.Encoding.UTF8)
                        For Each linea As String In CSV
                            writer.WriteLine(linea)
                        Next
                    End Using

                    ' Mensaje de confirmación
                    MessageBox.Show("CSV generado correctamente en: " & rutaCSV)
                End If
            End If

        Catch e As Exception
            MsgBox(e.Message)
        End Try
    End Sub

    Private Function ObtenerDatosName(columnIndex As Integer, worksheet As ExcelWorksheet)
        Dim value As Object
        Dim Names As New List(Of String)()

        For row As Integer = 1 To worksheet.Dimension.End.Row
            value = worksheet.Cells(row, columnIndex).Value
            If value IsNot Nothing Then
                Dim trimmedValue As String = value.ToString().Trim()
                If Not String.IsNullOrEmpty(trimmedValue) Then
                    Names.Add(trimmedValue)
                End If
            End If
        Next

        Names = Names.Distinct().ToList

        Return Names
    End Function

    Private Function ObtenerDatosInstance(columnIndex As Integer, worksheet As ExcelWorksheet, instanceFilter As String, templatecolumnIndex As Integer) As String
        Dim value As Object
        Dim filteredValues As String = String.Empty ' Inicializamos la variable como una cadena vacía
        Dim columnValue As Object = "Valor Nulo"

        For row As Integer = 1 To worksheet.Dimension.End.Row
            value = worksheet.Cells(row, columnIndex).Value

            If value IsNot Nothing Then
                Dim trimmedValue As String = value.ToString().Trim()

                ' Verificar si la celda contiene el filtro
                If Not String.IsNullOrEmpty(trimmedValue) AndAlso trimmedValue = instanceFilter Then
                    If templatecolumnIndex <> 10000 Then
                        columnValue = worksheet.Cells(row, templatecolumnIndex).Value
                    End If

                    If columnValue IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(columnValue.ToString()) Then
                        filteredValues = columnValue.ToString().Trim() ' Asignamos el valor de la columna
                        Return filteredValues ' Salimos de la función y devolvemos el valor encontrado
                    End If
                End If
            End If
        Next

        Return filteredValues ' Retorna una cadena vacía si no encuentra el filtro
    End Function

    Private Function ObtenerDatosColumna(columnIndex As Integer, worksheet As ExcelWorksheet, instanceFilter As String, mapcolumnIndex As Integer) As List(Of String)
        Dim value As Object
        Dim filteredValues As New List(Of String)() ' Inicializamos la lista
        Dim columnValue As Object = "Valor Nulo"

        For row As Integer = 1 To worksheet.Dimension.End.Row
            value = worksheet.Cells(row, columnIndex).Value

            If value IsNot Nothing Then
                Dim trimmedValue As String = value.ToString().Trim()

                ' Verificar si la celda contiene el filtro
                If Not String.IsNullOrEmpty(trimmedValue) AndAlso trimmedValue = instanceFilter Then
                    ' Obtener el valor de la columna de la misma fila
                    If mapcolumnIndex <> 10000 Then
                        columnValue = worksheet.Cells(row, mapcolumnIndex).Value
                    End If

                    If columnValue IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(columnValue.ToString()) Then
                        filteredValues.Add(columnValue.ToString().Trim()) ' Añadir el valor a la lista
                    End If
                End If
            End If
        Next

        Return filteredValues
    End Function

    Private Function ObtenerDatosFila(columnIndex As Integer, worksheet As ExcelWorksheet, instanceFilter As String, columndIndex As Integer()) As List(Of String)
        Dim value As Object
        Dim filteredValues As New List(Of String) ' Inicializamos la lista
        Dim columnValue As Object = "Valor Nulo"

        For row As Integer = 1 To worksheet.Dimension.End.Row
            value = worksheet.Cells(row, columnIndex).Value

            If value IsNot Nothing Then
                Dim trimmedValue As String = value.ToString().Trim()

                ' Verificar si la celda contiene el filtro
                If Not String.IsNullOrEmpty(trimmedValue) AndAlso trimmedValue = instanceFilter Then
                    ' Obtener el valor de la columna de la misma fila
                    For Each column In columndIndex
                        columnValue = worksheet.Cells(row, column).Value

                        If columnValue IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(columnValue.ToString()) Then
                            filteredValues.Add(columnValue.ToString().Trim()) ' Añadir el valor a la lista
                        End If
                    Next
                End If
            End If
        Next

        Return filteredValues
    End Function

    Private Function ObtenerDatosColumnaEngUnits(columnIndex As Integer, worksheet As ExcelWorksheet, instanceFilter As String, mapcolumnIndex As Integer) As List(Of String)
        Dim value As Object
        Dim filteredValues As New List(Of String)() ' Inicializamos la lista
        Dim columnValue As Object = "Valor Nulo"

        For row As Integer = 1 To worksheet.Dimension.End.Row
            value = worksheet.Cells(row, columnIndex).Value

            If value IsNot Nothing Then
                Dim trimmedValue As String = value.ToString().Trim()

                ' Verificar si la celda contiene el filtro
                If Not String.IsNullOrEmpty(trimmedValue) AndAlso trimmedValue = instanceFilter Then
                    ' Obtener el valor de la columna de la misma fila
                    If mapcolumnIndex <> 10000 Then
                        columnValue = worksheet.Cells(row, mapcolumnIndex).Value
                    End If

                    If columnValue IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(columnValue.ToString()) Then
                        filteredValues.Add(columnValue.ToString().Trim()) ' Añadir el valor a la lista
                    Else
                        filteredValues.Add("n/d")
                    End If
                End If
            End If
        Next

        Return filteredValues
    End Function

    Private Function ObtenerDatosCSV(columnIndex As Integer, worksheet As ExcelWorksheet, columnIndex1 As Integer, ColumnIndex2 As Integer, ColumnIndex3 As Integer) As List(Of String)
        Dim value As Object
        Dim filteredValues As New List(Of String)() ' Inicializamos la lista
        Dim columnValue1 As Object = "Valor Nulo"
        Dim columnValue2 As Object = "Valor Nulo"
        Dim columnValue3 As Object = "T Nulo"

        For row As Integer = 1 To worksheet.Dimension.End.Row
            value = worksheet.Cells(row, columnIndex).Value

            If value IsNot Nothing Then
                Dim trimmedValue As String = value.ToString().Trim()

                ' Verificar si la celda contiene el filtro
                If Not String.IsNullOrEmpty(trimmedValue) Then
                    ' Obtener el valor de la columna de la misma fila
                    If columnIndex1 <> 10000 Then
                        columnValue1 = worksheet.Cells(row, columnIndex1).Value
                    End If
                    If ColumnIndex2 <> 10000 Then
                        columnValue2 = worksheet.Cells(row, ColumnIndex2).Value
                    End If

                    If ColumnIndex3 <> 10000 Then
                        columnValue3 = worksheet.Cells(row, ColumnIndex3).Value
                    End If

                    If columnValue1 IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(columnValue1.ToString()) And columnValue2 IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(columnValue2.ToString()) And columnValue3 IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(columnValue2.ToString()) Then
                        If columnValue3.ToString.Contains("$") Then
                            filteredValues.Add("""" & columnValue1.ToString().Trim() & """" & "," & """" & columnValue2.ToString.Trim() & """") ' Añadir el valor a la lista
                        End If
                    End If
                End If
            End If
        Next

        Return filteredValues
    End Function

    Private Function ObtenerDatosPorCabecera(columnIndex As Integer, worksheet As ExcelWorksheet, instanceFilter As String, TitleName As String) As String
        Dim value As Object
        Dim filteredValues As String = String.Empty ' Inicializamos la variable como una cadena vacía
        Dim columnValue As Object = "Valor Nulo"
        Dim templatecolumnIndex As Integer = ObtenerIndicePorCabecera(worksheet, TitleName)

        For row As Integer = 1 To worksheet.Dimension.End.Row
            value = worksheet.Cells(row, columnIndex).Value

            If value IsNot Nothing Then
                Dim trimmedValue As String = value.ToString().Trim()

                ' Verificar si la celda contiene el filtro
                If Not String.IsNullOrEmpty(trimmedValue) AndAlso trimmedValue = instanceFilter Then
                    If templatecolumnIndex <> 10000 Then
                        columnValue = worksheet.Cells(row, templatecolumnIndex).Value
                    End If

                    If columnValue IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(columnValue.ToString()) Then
                        filteredValues = columnValue.ToString().Trim() ' Asignamos el valor de la columna
                        Return filteredValues ' Salimos de la función y devolvemos el valor encontrado
                    End If
                End If
            End If
        Next

        Return filteredValues ' Retorna una cadena vacía si no encuentra el filtro
    End Function

    Private Function ObtenerDatosColumnaPorCabecera(columnIndex As Integer, worksheet As ExcelWorksheet, instanceFilter As String, TitleName As String) As List(Of String)
        Dim value As Object
        Dim filteredValues As New List(Of String)() ' Inicializamos la lista
        Dim columnValue As Object = "Valor Nulo"

        Dim titlecolumnIndex As Integer = ObtenerIndicePorCabecera(worksheet, TitleName)

        For row As Integer = 1 To worksheet.Dimension.End.Row
            value = worksheet.Cells(row, columnIndex).Value

            If value IsNot Nothing Then
                Dim trimmedValue As String = value.ToString().Trim()

                ' Verificar si la celda contiene el filtro
                If Not String.IsNullOrEmpty(trimmedValue) AndAlso trimmedValue = instanceFilter Then
                    ' Obtener el valor de la columna de la misma fila
                    If titlecolumnIndex <> 10000 Then
                        columnValue = worksheet.Cells(row, titlecolumnIndex).Value
                    End If

                    If columnValue IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(columnValue.ToString()) Then
                        filteredValues.Add(columnValue.ToString().Trim()) ' Añadir el valor a la lista
                    End If
                End If
            End If
        Next

        Return filteredValues
    End Function

    Private Function ObtenerIndicePorCabecera(ws As ExcelWorksheet, nombreColumna As String) As Integer
        If nombreColumna = "10000" Then
            Return 10000
        End If


        ' Asegurarse de que existe al menos una tabla en la hoja
        If ws.Tables.Count = 0 Then
            Throw New Exception("La hoja no contiene ninguna tabla definida.")
        End If

        ' Obtener la primera tabla (asumiendo que solo hay una)
        Dim tabla = ws.Tables(0)

        ' Recorrer las columnas definidas en la tabla
        For i As Integer = 0 To tabla.Columns.Count - 1
            Dim nombreActual As String = tabla.Columns(i).Name.Trim()

            If nombreActual.Equals(nombreColumna, StringComparison.OrdinalIgnoreCase) Then
                ' Devolvemos el índice de la columna relativa al worksheet, no solo dentro de la tabla
                Return tabla.Range.Start.Column + i
            End If
        Next

        Throw New Exception($"No se encontró la columna con el encabezado '{nombreColumna}' en la tabla.")
    End Function
End Class