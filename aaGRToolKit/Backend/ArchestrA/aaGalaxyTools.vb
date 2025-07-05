Imports System.IO
Imports aaGRAccessApp
Imports ArchestrA.Visualization.GraphicAccess
Imports System.Runtime.InteropServices

''' <summary>
''' The aaTemplateExtract class interfaces with GRAccess to connect, collect, and parse data from the Galaxy.
''' All the hard work is in this class, but aaTemplateData contains the actual data structure.
''' </summary>
Public Class aaGalaxyTools
    Public showLogin As Boolean
    Public errorMessage As String
    Public loggedIn As Boolean

    Private resultStatus As Boolean
    Private resultCount As Integer
    Private galaxyNode As String
    Private galaxyName As String
    Private templateName As String
    Private grAccess As aaGRAccessApp.GRAccessApp
    Private myGalaxy As aaGRAccessApp.IGalaxy

    ' Carpeta centralizada para gráficos
    Private ReadOnly GraphicsFolder As String = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Graphics")

    Public Sub ExportSymbol(SymbolName As String)
        Dim ff As New ArchestrA.Visualization.GraphicAccess.GraphicAccess()
        Try
            If Not Directory.Exists(GraphicsFolder) Then
                Directory.CreateDirectory(GraphicsFolder)
            End If

            Dim outputPath As String = Path.Combine(GraphicsFolder, SymbolName & ".xml")
            ff.ExportGraphicToXml(myGalaxy, SymbolName, outputPath)
            Dim cmd = ff.CommandResult

            If Not cmd.Successful Then
                MsgBox("Error exportando: " & cmd.CustomMessage, MsgBoxStyle.Critical, "Exportación de símbolo")
            Else
                MsgBox("Exportación correcta, XML en: " & outputPath, MsgBoxStyle.Information, "Exportación de símbolo")
            End If
        Catch ex As Exception
            MsgBox("Excepción exportando símbolo: " & ex.Message, MsgBoxStyle.Critical, "Exportación de símbolo")
        End Try
    End Sub

    Public Sub ImportSymbol(SymbolName As String, folderPath As String, OverWritte As Boolean)
        Dim ff As New ArchestrA.Visualization.GraphicAccess.GraphicAccess()
        Try
            ff.ImportGraphicFromXml(myGalaxy, SymbolName, folderPath, OverWritte)
            Dim cmd = ff.CommandResult

            If Not cmd.Successful Then
                MsgBox("Error importando: " & cmd.CustomMessage, MsgBoxStyle.Critical, "Importación de símbolo")
            Else
                MsgBox("Importación correcta, añadido el gráfico: " & SymbolName, MsgBoxStyle.Information, "Importación de símbolo")
            End If
        Catch ex As Exception
            MsgBox("Excepción importando símbolo: " & ex.Message, MsgBoxStyle.Critical, "Importación de símbolo")
        End Try
    End Sub

    ' ------- DLL PATH CONFIG ----------
    <DllImport("kernel32.dll", SetLastError:=True)>
    Private Shared Function SetDllDirectory(lpPathName As String) As Boolean
    End Function

    ''' <summary>
    ''' Initializes and sets up our GRAccess client app.
    ''' </summary>
    Public Sub New()
        ' Intenta establecer la ruta donde probablemente está GRAccess.dll
        Dim dllPaths As String() = {
            "C:\Program Files (x86)\ArchestrA\Framework\Bin",
            "D:\ArchestrA\Common Files",
            Application.StartupPath
        }

        For Each path As String In dllPaths
            If Directory.Exists(path) Then
                SetDllDirectory(path)
                Exit For
            End If
        Next

        ' Ahora creamos el objeto normalmente
        grAccess = New aaGRAccessApp.GRAccessApp
        loggedIn = False
        showLogin = True
    End Sub

    ''' <summary>
    ''' Queries for all the galaxy names on this node.
    ''' </summary>
    ''' <param name="NodeName">The node name as a string (e.g. "localhost").</param>
    ''' <returns>A collection of galaxy names as strings.</returns>
    Public Function getGalaxies(ByVal NodeName As String) As Collection
        Dim galaxies As aaGRAccessApp.IGalaxies
        Dim galaxy As aaGRAccessApp.IGalaxy
        Dim galaxyList As New Collection

        galaxyNode = NodeName

        Try
            galaxies = grAccess.QueryGalaxies(galaxyNode)
            resultStatus = grAccess.CommandResult.Successful
            If resultStatus And (galaxies IsNot Nothing) And galaxies.count > 0 Then
                For Each galaxy In galaxies
                    galaxyList.Add(galaxy.Name)
                Next
            Else
                Throw New ApplicationException("No Galaxies detected on this node")
            End If

        Catch e As Exception
            MessageBox.Show("Ocurrió un error al obtener las galaxias: " & e.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Return galaxyList
    End Function

    ''' <summary>
    ''' Importa objetos a la galaxia desde un archivo especificado.
    ''' </summary>
    ''' <param name="FileInfo">Ruta completa del archivo a importar.</param>
    ''' <param name="Overwrite">Indica si se deben sobrescribir los objetos existentes.</param>
    ''' <returns>Lista de mensajes que indican el resultado de la importación, 
    ''' incluyendo éxito, advertencias, errores o excepciones.</returns>
    Public Function GalaxyImport(FileInfo As String, Overwrite As Boolean) As List(Of String)
        Dim Results As New List(Of String)
        Try
            ' Ejecutar la importación
            myGalaxy.ImportObjects(FileInfo, Overwrite)

            ' Verificar si la importación fue completamente exitosa
            If myGalaxy.CommandResults IsNot Nothing AndAlso myGalaxy.CommandResults.CompletelySuccessful Then
                Results.Add(FileInfo & " importación exitosa.")
            ElseIf myGalaxy.CommandResults IsNot Nothing Then
                ' Recorrer los mensajes de error o advertencia
                For i As Integer = 0 To myGalaxy.CommandResults.count - 1
                    Dim msg As String = myGalaxy.CommandResults.Item(i).CustomMessage
                    Results.Add("Error o advertencia: " & msg)
                    Debug.WriteLine(msg)
                Next
            Else
                Results.Add("No se recibieron resultados de la importación.")
            End If

        Catch ex As Exception
            Results.Add("Excepción durante importación: " & ex.Message)
        End Try

        Return Results
    End Function

    ''' <summary>
    ''' Given a Galaxy Name, sets our Client to that Galaxy.
    ''' </summary>
    ''' <param name="galaxyName">A string that is the Galaxy name.</param>
    ''' <returns>Nothing</returns>
    Public Function setGalaxy(ByVal galaxyName)
        showLogin = True
        loggedIn = False
        Try
            If (galaxyNode IsNot Nothing) And (galaxyName IsNot Nothing) Then
                myGalaxy = grAccess.QueryGalaxies(galaxyNode)(galaxyName)
            Else
                Throw New ApplicationException("No Node or Galaxy Selected")
            End If
        Catch e As Exception
            MessageBox.Show("Ocurrió un error: " & e.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return 0
    End Function

    ''' <summary>
    ''' Tries to login to the Galaxy.
    ''' </summary>
    ''' <param name="user">User name, include Domain (e.g. "Domain\user") if authenticating against a domain.</param>
    ''' <param name="password">Password</param>
    ''' <returns>Status of login attempt.</returns>
    Public Function login(ByVal user, ByVal password) As Integer
        Try
            myGalaxy.Login(user, password)
            If myGalaxy.CommandResult.Successful Then
                loggedIn = True
                Return 0
            Else
                loggedIn = False
                Return -1
            End If
        Catch e As Exception
            loggedIn = False
            MessageBox.Show("Ocurrió un error durante el inicio de sesión: " & e.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return -2
        End Try

    End Function


    ' ------- Templates ----------
    ''' <summary>
    ''' Discovers all of the templates in this Galaxy.
    ''' </summary>
    ''' <param name="HideBaseTemplates">If set, will not return the base templates that are in every Galaxy.</param>
    ''' <returns>A collection of strings with the template names listed.</returns>
    ''' <remarks>Does not return instances or checked in templates.</remarks>
    Public Function getTemplates(Optional ByVal HideBaseTemplates As Boolean = False) As Collection
        Dim templateList As New Collection
        Dim gTemplates As aaGRAccessApp.IgObjects
        Dim gTemplate As aaGRAccessApp.IgObject

        Dim BaseTemplates() As String = New String() {"$Boolean", "$Integer", "$Double", "$Float", "$String", "$FieldReference", "$UserDefined", "$AnalogDevice", "$AppEngine", "$Area", "$DDESuiteLinkClient", "$DiscreteDevice", "$InTouchProxy", "$InTouchViewApp", "$OPCClient", "$RedundantDIObject", "$Sequencer", "$SQLData", "$Switch", "$ViewEngine", "$WinPlatform"}

        Try
            If loggedIn Then
                gTemplates = myGalaxy.QueryObjects(aaGRAccessApp.EgObjectIsTemplateOrInstance.gObjectIsTemplate,
                                                aaGRAccessApp.EConditionType.checkoutStatusIs,
                                                aaGRAccessApp.ECheckoutStatus.notCheckedOut,
                                                aaGRAccessApp.EMatch.MatchCondition)
                resultStatus = myGalaxy.CommandResult.Successful
                If resultStatus And (gTemplates IsNot Nothing) And (gTemplates.count > 0) Then
                    For Each gTemplate In gTemplates
                        If Not (HideBaseTemplates And BaseTemplates.Contains(gTemplate.Tagname)) Then
                            templateList.Add(gTemplate.Tagname)
                        End If
                    Next
                Else
                    Throw New ApplicationException("No templates found")
                End If
            End If

        Catch e As Exception
            MessageBox.Show("Ocurrió un error: " & e.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
        Return templateList
    End Function

    ''' <summary>
    ''' This is the master function that is initiated for each template. Sub-functions will gather all of the attributes and scripts.
    ''' </summary>
    ''' <param name="TemplateName">The template name that is desired to get data from.</param>
    ''' <returns>An aaTemplate class of all of the template data (scripts, UDAs, field attributes).</returns>
    Public Function getTemplateData(ByVal TemplateName As String) As aaTemplate
        Dim templateList(1) As String
        Dim gTemplates As aaGRAccessApp.IgObjects
        Dim gTemplate As aaGRAccessApp.ITemplate
        Dim templateData As New aaTemplate()
        Dim gDerivedTemplate As String

        Try
            If loggedIn Then
                ' Convert the individual template name to an array. 
                templateList(0) = TemplateName
                ' query the galaxy for this template's data
                gTemplates = myGalaxy.QueryObjectsByName(aaGRAccessApp.EgObjectIsTemplateOrInstance.gObjectIsTemplate, templateList)
                resultStatus = myGalaxy.CommandResult.Successful
                If resultStatus And (gTemplates IsNot Nothing) And (gTemplates.count > 0) Then
                    gTemplate = gTemplates(1)

                    ' get all of the Configurable Attributes
                    Dim gAttributes = gTemplate.Attributes

                    ' Get the current template data
                    templateData = New aaTemplate(gTemplate.Tagname,
                                              DataTypeGet("ConfigVersion", gAttributes),
                                              GetFieldAttributesDiscrete(gAttributes),
                                              GetFieldAttributesAnalog(gAttributes))

                    ' Check for derived templates and recursively get their data
                    gDerivedTemplate = gTemplate.DerivedFrom
                    If Not String.IsNullOrEmpty(gDerivedTemplate) Then
                        Dim derivedTemplateData As aaTemplate = getTemplateData(gDerivedTemplate)
                        ' Sumar los FieldAttributes de la plantilla derivada
                        templateData.AddAttributes(derivedTemplateData)
                    End If
                End If
            Else
                Throw New ApplicationException("Not Logged In")
            End If

        Catch e As Exception
            frmMain.LogBox.Items.Add("Error occurred: " & e.Message)
        End Try

        Return templateData
    End Function

    ''' <summary>
    ''' Crea una nueva plantilla en la galaxia basada en una plantilla existente y le asigna atributos personalizados.
    ''' </summary>
    ''' <param name="TemplateName">Nombre de la plantilla base desde la cual se clonará la nueva plantilla.</param>
    ''' <param name="NewTemplateName">Nombre de la nueva plantilla que se creará.</param>
    ''' <param name="DiscAttrList">Lista de atributos digitales (discretos) que se agregarán a la nueva plantilla.</param>
    ''' <param name="AnalogAttrList">Lista de atributos analógicos que se agregarán a la nueva plantilla.</param>
    ''' <param name="ScriptList">Lista de scripts que se asignarán a la nueva plantilla.</param>
    ''' <param name="UDAList">Lista de atributos definidos por el usuario (UDA) que se asignarán a la plantilla.</param>
    ''' <param name="List">Lista de arrays de strings con parámetros adicionales utilizados por la función AddNewTemAttr.</param>
    ''' <remarks>
    ''' La función verifica si el usuario está logueado en la galaxia antes de realizar operaciones.
    ''' En caso de error (por ejemplo, plantilla base no encontrada), se informa en el log de la interfaz.
    ''' </remarks>
    Public Sub CreateTemplate(ByVal TemplateName As String,
                              ByVal NewTemplateName As String,
                              ByVal DiscAttrList As List(Of aaTemplateData.aaNewField),
                              ByVal AnalogAttrList As List(Of aaTemplateData.aaNewField),
                              ByVal ScriptList As List(Of aaTemplateData.aaNewField),
                              ByVal UDAList As List(Of aaTemplateData.aaNewField),
                              ByVal List As List(Of String()))

        Try
            If loggedIn Then
                Dim templateList(0) As String
                templateList(0) = TemplateName

                ' Consultar plantilla
                Dim gTemplates As aaGRAccessApp.IgObjects = myGalaxy.QueryObjectsByName(aaGRAccessApp.EgObjectIsTemplateOrInstance.gObjectIsTemplate, templateList)
                If gTemplates Is Nothing OrElse gTemplates.count = 0 Then
                    frmMain.LogBox.Items.Add("Error: Plantilla '" & TemplateName & "' no encontrada.")
                    Exit Sub
                End If

                Dim baseTemplate As aaGRAccessApp.ITemplate = CType(gTemplates(1), aaGRAccessApp.ITemplate)
                Dim newTemplate As aaGRAccessApp.ITemplate = baseTemplate.CreateTemplate(NewTemplateName, True)

                AddNewTemAttr(newTemplate,
                              DiscAttrList,
                              AnalogAttrList,
                              ScriptList,
                              UDAList,
                              List)

                frmMain.LogBox.Items.Add("Plantilla y atributos guardados y check-in realizado.")
            Else
                Throw New ApplicationException("No se ha iniciado sesión en la galaxia.")
            End If
        Catch ex As Exception
            frmMain.LogBox.Items.Add("Error general: " & ex.Message)
        End Try
    End Sub

    ' ------- Instances ----------

    ''' <summary>
    ''' Crea una nueva instancia en la galaxia a partir de una plantilla existente, asignándole atributos personalizados y una ubicación (área).
    ''' </summary>
    ''' <param name="TemplateName">Nombre de la plantilla base desde la que se generará la instancia.</param>
    ''' <param name="InstanceName">Nombre que se asignará a la nueva instancia creada.</param>
    ''' <param name="AreaName">Nombre del área de la galaxia donde se ubicará la instancia.</param>
    ''' <param name="LoneAttrText">Lista de valores para atributos individuales que se deben asignar a la instancia.</param>
    ''' <param name="ArrayAttrText">Lista de valores para atributos tipo array que se deben asignar a la instancia.</param>
    ''' <param name="AloneAttributes">Arreglo de nombres de los atributos individuales que se configurarán.</param>
    ''' <param name="arrayAttributes">Arreglo de nombres de los atributos tipo array que se configurarán.</param>
    ''' <param name="EngUnits">Arreglo de unidades de ingeniería asociadas a los atributos (si aplica).</param>
    ''' <remarks>
    ''' La función requiere que el usuario esté autenticado (loggedIn = True). 
    ''' En caso de error (como plantilla no encontrada o no autenticado), se muestra un mensaje descriptivo.
    ''' Llama internamente a la función EditarCfg para configurar los atributos en la nueva instancia.
    ''' </remarks>
    Public Sub CreateInstance(ByVal TemplateName As String, ByVal InstanceName As String, ByVal AreaName As String, ByVal LoneAttrText As List(Of String), ByVal ArrayAttrText As List(Of String), AloneAttributes As String(), arrayAttributes As String(), EngUnits() As String)
        Dim templateList(1) As String
        Dim gTemplates As aaGRAccessApp.IgObjects
        Dim gTemplate As aaGRAccessApp.ITemplate
        Dim instance As aaGRAccessApp.IInstance
        Dim templateData As New aaTemplate()
        Dim FinalMap As String() = {}

        Try
            ' Convert the individual template name to an array. 
            templateList(0) = TemplateName
            ' query the galaxy for this template's data
            gTemplates = myGalaxy.QueryObjectsByName(aaGRAccessApp.EgObjectIsTemplateOrInstance.gObjectIsTemplate, templateList)
            resultStatus = myGalaxy.CommandResult.Successful
            If resultStatus And (gTemplates IsNot Nothing) And (gTemplates.count > 0) Then
                gTemplate = gTemplates(1)
                instance = gTemplate.CreateInstance(InstanceName, True)
                instance.CheckOut()
                EditarCfg(LoneAttrText, ArrayAttrText, instance, AloneAttributes, arrayAttributes, EngUnits)
                instance.Save()
                instance.Area = AreaName
            Else
                MessageBox.Show("Error: Template " & TemplateName & " not found")
            End If

        Catch e As Exception
            MessageBox.Show("Ocurrió un error: " & e.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' Despliega o retira el despliegue (undeploy) de una instancia dentro de la galaxia, según el modo especificado.
    ''' </summary>
    ''' <param name="InstanceName">Nombre de la instancia que se desea desplegar o retirar.</param>
    ''' <param name="Cascade">Indica si la acción debe aplicarse de forma recursiva (en cascada) a objetos relacionados.</param>
    ''' <param name="mode">
    ''' Modo de operación:
    ''' 1 = Undeploy (retirar despliegue),
    ''' Cualquier otro valor = Deploy (desplegar instancia).
    ''' </param>
    ''' <remarks>
    ''' Requiere que el usuario esté autenticado previamente (loggedIn = True).
    ''' La función consulta la instancia por nombre y realiza la acción correspondiente.
    ''' En caso de error o instancia no encontrada, se muestra un mensaje descriptivo.
    ''' </remarks>
    Public Sub deployUndeployInstance(ByVal InstanceName As String, Cascade As Boolean, mode As Integer)
        Dim InstanceList(1) As String
        Dim gInstances As aaGRAccessApp.IgObjects
        Dim instance As aaGRAccessApp.IInstance

        Try
            If loggedIn Then

                InstanceList(0) = InstanceName
                gInstances = myGalaxy.QueryObjectsByName(aaGRAccessApp.EgObjectIsTemplateOrInstance.gObjectIsInstance, InstanceList)
                resultStatus = myGalaxy.CommandResult.Successful

                If mode = 1 Then
                    If resultStatus Then
                        instance = gInstances(1)
                        instance.Undeploy(EForceOffScan.doForceOffScan, ECascade.doCascade, markAsUndeployedOnFailure:=False)
                    End If
                Else
                    If resultStatus Then
                        instance = gInstances(1)
                        If Cascade Then
                            instance.Deploy(EActionForCurrentlyDeployedObjects.deployChanges, ESkipIfCurrentlyUndeployed.dontSkipIfCurrentlyUndeployed, EDeployOnScan.doDeployOnScan, EForceOffScan.doForceOffScan, ECascade.doCascade, True)
                        Else
                            instance.Deploy(EActionForCurrentlyDeployedObjects.deployChanges, ESkipIfCurrentlyUndeployed.dontSkipIfCurrentlyUndeployed, EDeployOnScan.doDeployOnScan, EForceOffScan.doForceOffScan, ECascade.dontCascade, True)
                        End If
                    Else
                        MessageBox.Show("Error: Instance" & InstanceName & "not found")
                    End If
                End If
            Else
                Throw New ApplicationException("Not Logged In")
            End If

        Catch e As Exception
            MessageBox.Show("Ocurrió un error: " & e.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    ' ------- Instance Attributes ----------
    ''' <summary>
    ''' Asigna un conjunto de valores a un atributo array del tipo <c>Cfg_Mapeado</c> en ArchestrA.
    ''' </summary>
    ''' <param name="gAttributes">Colección de atributos de la instancia en la que se desea modificar el array.</param>
    ''' <param name="valores">Arreglo de strings que contiene los valores a insertar en el atributo array.</param>
    ''' <param name="ArrayAttributes">Nombre del atributo array que se desea modificar.</param>
    ''' <remarks>
    ''' - Se admite un máximo de 50 elementos; si hay menos, se completa con cadenas vacías.
    ''' - Si un valor es <c>null</c> o vacío, se inserta una cadena vacía en su lugar.
    ''' - Los índices del array son 1-based según el estándar de ArchestrA.
    ''' - Registra el resultado o cualquier error en el LogBox principal de la aplicación.
    ''' </remarks>
    Private Sub AgregarElementosACfgMapeado(gAttributes As aaGRAccess.IAttributes, valores As String(), ArrayAttributes As String)
        Try
            Dim Attr As aaGRAccess.IAttribute = gAttributes.Item(ArrayAttributes)

            If Attr Is Nothing Then
                frmMain.LogBox.Items.Add($"El atributo {ArrayAttributes} no existe.")
                Return
            End If

            ' Crear un nuevo MxValue para el array
            Dim MXVal_ As New aaGRAccess.MxValue()

            ' Iniciar el array estableciendo el tipo de datos, por ejemplo, String
            If valores.Length > 0 Then
                MXVal_.PutString(valores(0)) ' Establecer el tipo de datos usando el primer valor
            End If

            ' Obtener el tamaño máximo del array
            Dim maxItems As Integer = 50

            ' Asegurarse de no agregar más de 50 elementos
            Dim elementosAAgregar As Integer = Math.Min(valores.Length, maxItems)

            ' Recorrer los valores y agregar cada uno al array utilizando PutElement
            For i As Integer = 0 To elementosAAgregar - 1 ' Asegurarse de no exceder el tamaño de 'valores'
                Dim Elemento As New aaGRAccess.MxValue()

                ' Verificar si el valor está vacío o nulo
                If String.IsNullOrEmpty(valores(i)) Then
                    ' Establecer el valor vacío
                    Elemento.PutString("")
                Else
                    ' Establecer el valor real
                    Elemento.PutString(valores(i))
                End If

                ' Usar i + 1 para asegurar índices secuenciales (1-based)
                MXVal_.PutElement(i + 1, Elemento)
            Next

            ' Si el tamaño de 'valores' es menor que 'maxItems', completa el array con valores vacíos
            For i As Integer = elementosAAgregar To maxItems - 1
                Dim Elemento As New aaGRAccess.MxValue()
                Elemento.PutString("") ' Valor vacío
                MXVal_.PutElement(i + 1, Elemento) ' Usar i + 1 para asegurar índices secuenciales
            Next

            ' Asignar el array al atributo Cfg_Mapeado
            Attr.SetValue(MXVal_)

            frmMain.LogBox.Items.Add("Los elementos se han agregado correctamente al atributo " & ArrayAttributes)
        Catch ex As Exception
            frmMain.LogBox.Items.Add("Error al agregar los elementos: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Actualiza los atributos configurables, arrays y unidades de ingeniería de una instancia dada.
    ''' </summary>
    ''' <param name="Aloneattr">Lista de valores individuales para atributos simples.</param>
    ''' <param name="arrayeattr">Lista de cadenas que contienen valores separados por líneas para atributos tipo array.</param>
    ''' <param name="Instance">Instancia de ArchestrA que se va a modificar.</param>
    ''' <param name="ConfigurableAttr">Arreglo con los nombres de atributos simples a actualizar.</param>
    ''' <param name="ArrayAttributes">Arreglo con los nombres de atributos tipo array a actualizar.</param>
    ''' <param name="EngUnits">Arreglo con valores de unidades de ingeniería para actualizar atributos específicos.</param>
    ''' <remarks>
    ''' - Utiliza las funciones <c>ToMxValue</c>, <c>AgregarElementosACfgMapeado</c> y <c>ToMxValueEngUnits</c> para establecer los valores.
    ''' - Guarda y hace check-in de la instancia al finalizar.
    ''' - Cualquier excepción se registra en el LogBox principal.
    ''' </remarks>
    Private Sub EditarCfg(Aloneattr As List(Of String), arrayeattr As List(Of String), Instance As aaGRAccessApp.IInstance, ConfigurableAttr As String(), ArrayAttributes As String(), EngUnits() As String)
        Dim gAttributes = Instance.Attributes

        Try
            For j As Integer = 0 To ConfigurableAttr.Length - 1
                ToMxValue(Aloneattr(j), gAttributes, Instance, ConfigurableAttr(j))
            Next

            For i As Integer = 0 To ArrayAttributes.Length - 1
                Dim FinalMap = arrayeattr(i).Split(New String() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)
                AgregarElementosACfgMapeado(gAttributes, FinalMap, ArrayAttributes(i))
            Next

            For e As Integer = 0 To EngUnits.Length - 1
                Dim Valor = EngUnits(e).Split("|"c)
                Dim ValorTrim As String() = Valor.Select(Function(v) v.Trim()).ToArray()

                If ValorTrim(1) <> "n/d" Then
                    ToMxValueEngUnits(ValorTrim(1), gAttributes, Instance, ValorTrim(0))
                End If
            Next

            Instance.Save()
            Instance.CheckIn()
        Catch ex As Exception
            frmMain.LogBox.Items.Add("Error " & ex.Message)
        End Try
    End Sub


    ' ------- Template Attributes ----------
    ''' <summary>
    ''' Gets all of the Analog Field Attributes for a given template. 
    ''' </summary>
    ''' <param name="gAttributes">All of the Configurable Attributes that were found for this template.</param>
    ''' <returns>A collection of all of the Analog Field Attributes using the aaFieldAttributeAnalog class.</returns>
    ''' <remarks></remarks>
    Private Function GetFieldAttributesAnalog(gAttributes As aaGRAccess.IAttributes) As Collection
        Dim FieldAttributes As New Collection()

        ' A list of the Field Attributes are stored in an XML fragment in the UserAttrData attribute
        Dim UserAttrData As XElement = XElement.Parse(gAttributes.Item("UserAttrData").value.GetString)

        Dim attrList = UserAttrData.<DiscreteAttr>.Attributes("Name")

        For Each attr In attrList
            Dim attrName = attr.Value

            ' Now, put all of the info together into one data set for this attribute
            Dim AnalogAttrData = New aaFieldAttributeAnalog(attrName,
                DataTypeGet("Tagname", gAttributes),
                DataTypeGet(attrName + ".Desc", gAttributes),
                DataTypeGet(attrName + ".Historized", gAttributes),
                DataTypeGet(attrName + ".LogDataChangeEvent", gAttributes),
                DataTypeGet(attrName + ".Alarmed", gAttributes),
                DataTypeGet(attrName + ".EngUnits", gAttributes))

            ' Finally, add it to a (growing) collection of field attributes
            FieldAttributes.Add(AnalogAttrData)
        Next

        Return FieldAttributes
    End Function

    ''' <summary>
    ''' Gets all of the Discrete Field Attributes for a given template. 
    ''' </summary>
    ''' <param name="gAttributes">All of the Configurable Attributes that were found for this template.</param>
    ''' <returns>A collection of all of the Discrete Field Attributes using the aaFieldAttributeDiscrete class.</returns>
    ''' <remarks></remarks>
    Private Function GetFieldAttributesDiscrete(gAttributes As aaGRAccess.IAttributes) As Collection
        Dim FieldAttributes As New Collection()

        ' A list of the Field Attributes are stored in an XML fragment in the UserAttrData attribute
        Dim UserAttrData As XElement = XElement.Parse(gAttributes.Item("UserAttrData").value.GetString)

        Dim attrList = UserAttrData.<AnalogAttr>.Attributes("Name")

        For Each attr In attrList
            Dim attrName = attr.Value

            ' Now, put all of the info together into one data set for this attribute
            Dim DiscreteAttrData = New aaFieldAttributeDiscrete(
                attrName,
                DataTypeGet("Tagname", gAttributes),
                DataTypeGet(attrName + ".Desc", gAttributes),
                DataTypeGet(attrName + ".Historized", gAttributes),
                DataTypeGet(attrName + ".LogDataChangeEvent", gAttributes),
                DataTypeGet(attrName + ".Alarmed", gAttributes),
                DataTypeGet(attrName + ".EngUnits", gAttributes))

            ' Finally, add it to a (growing) collection of field attributes
            FieldAttributes.Add(DiscreteAttrData)
        Next

        Return FieldAttributes
    End Function

    ''' <summary>
    ''' Obtiene una lista de nombres de atributos configurables para una plantilla dada.
    ''' </summary>
    ''' <param name="templatename">Nombre de la plantilla a consultar.</param>
    ''' <returns>
    ''' Lista de cadenas con los nombres de los atributos configurables de la plantilla.
    ''' Si la plantilla no existe, devuelve una lista con el valor "BadTemplate".
    ''' Si no se ha iniciado sesión, devuelve una lista con "NotLoggedIn".
    ''' En caso de error, devuelve una lista con "Error".
    ''' </returns>
    ''' <remarks>
    ''' - Filtra atributos cuyos nombres comienzan con "Scr_" o "Src_".
    ''' - Utiliza la API de ArchestrA para consultar los objetos y sus atributos.
    ''' - Muestra un mensaje de error en un MessageBox si ocurre una excepción.
    ''' </remarks>
    Public Function getTemplateAttributes(templatename As String) As List(Of String)
        Dim templateList(0) As String
        Dim gTemplates As aaGRAccessApp.IgObjects
        Dim gTemplate As aaGRAccessApp.ITemplate
        Dim gAttributes As aaGRAccess.IAttributes
        Dim attrList As New List(Of String)()

        Try
            If loggedIn Then
                ' Convert the individual template name to an array. 
                templateList(0) = templatename

                ' Query the galaxy for this template's data
                gTemplates = myGalaxy.QueryObjectsByName(aaGRAccessApp.EgObjectIsTemplateOrInstance.gObjectIsTemplate, templateList)
                resultStatus = myGalaxy.CommandResult.Successful

                If resultStatus AndAlso gTemplates IsNot Nothing AndAlso gTemplates.count > 0 Then
                    gTemplate = CType(gTemplates(1), aaGRAccessApp.ITemplate)

                    ' Get all of the Configurable Attributes
                    gAttributes = gTemplate.Attributes

                    ' Iterate correctly over attributes using index
                    Dim prefixes As String() = {"Scr_" & "Src_"}
                    For i As Integer = 1 To gAttributes.count
                        Dim attr = gAttributes(i)
                        If Not prefixes.Any(Function(p) attr.Name.StartsWith(p)) Then
                            attrList.Add(attr.Name)
                        End If
                    Next

                    Return attrList
                Else
                    attrList.Add("BadTemplate")
                    Return attrList
                End If
            Else
                attrList.Add("NotLoggedIn")
                Return attrList
            End If
        Catch ex As Exception
            MessageBox.Show("Error Ocurred: " & ex.Message & vbCrLf & ex.StackTrace)
            attrList.Add("Error")
            Return attrList
        End Try
    End Function

    ''' <summary>
    ''' Agrega nuevos atributos a una plantilla dada, incluyendo atributos discretos, analógicos, scripts y UDA.
    ''' </summary>
    ''' <param name="newTemplate">La plantilla (ITemplate) a la que se agregarán los atributos.</param>
    ''' <param name="DiscAttrList">Lista de atributos discretos a agregar.</param>
    ''' <param name="AnalogAttrList">Lista de atributos analógicos a agregar.</param>
    ''' <param name="ScriptList">Lista de atributos de tipo script a agregar.</param>
    ''' <param name="UDAList">Lista de atributos definidos por el usuario (UDA) a agregar.</param>
    ''' <param name="List">Lista de arreglos de cadenas que contiene parámetros para la creación de nuevos Field Attributes.</param>
    ''' <remarks>
    ''' - Realiza un CheckOut de la plantilla antes de modificarla y un CheckIn al finalizar.
    ''' - El orden de creación de UDA, atributos discretos y analógicos es importante para evitar errores.
    ''' - Actualiza atributos XML específicos como 'UserAttrData' y 'Extensions' con los nuevos atributos.
    ''' - Guarda la plantilla varias veces para asegurar la correcta aplicación de cambios.
    ''' </remarks>
    Public Sub AddNewTemAttr(ByVal newTemplate As aaGRAccessApp.ITemplate,
                             ByVal DiscAttrList As List(Of aaTemplateData.aaNewField),
                             ByVal AnalogAttrList As List(Of aaTemplateData.aaNewField),
                             ByVal ScriptList As List(Of aaTemplateData.aaNewField),
                             ByVal UDAList As List(Of aaTemplateData.aaNewField),
                             ByVal List As List(Of String()))

        ' Check-Out para editar
        newTemplate.CheckOut()

        ' El orden de escritura es importante si intentas crear los UDA despues de los FA dara fallo, desconozco la razon pero al seguir el orden no da problemas.
        For Each UDA In UDAList
            If UDA.NameAttr IsNot Nothing Or UDA.NameAttr <> "" Then
                createNewUDA(newTemplate, UDA.NameAttr.Split("|"c)(0), UDA.Values(0))
            End If
        Next

        newTemplate.Save()

        Dim UDOAttributes As IAttributes = newTemplate.Attributes
        Dim UDOUserAttrDataAttribute As IAttribute = UDOAttributes("UserAttrData")
        Dim MxVal As New MxValueClass()

        Dim xmlString As String = "<AttrXML>"

        For Each sublist In DiscAttrList
            xmlString &= "<DiscreteAttr Name=""" & sublist.NameAttr.Split("|"c)(0) & """ />"
        Next

        For Each sublist In AnalogAttrList
            xmlString &= "<AnalogAttr Name=""" & sublist.NameAttr.Split("|"c)(0) & """ />"
        Next

        xmlString &= "</AttrXML>"

        MxVal.PutString(xmlString)
        UDOUserAttrDataAttribute.SetValue(MxVal)

        newTemplate.Save()

        For Each name In DiscAttrList
            If name IsNot Nothing Then
                createNewFA(newTemplate, name, List(0), List(1))
            End If
        Next

        For Each name In AnalogAttrList
            If name IsNot Nothing Then
                createNewFA(newTemplate, name, List(2), List(3))
            End If
        Next

        newTemplate.Save()

        UDOUserAttrDataAttribute = UDOAttributes("Extensions")

        xmlString = "<ExtensionInfo><ObjectExtension>"

        For Each sublist In ScriptList
            xmlString &= "<Extension Name=""" & sublist.NameAttr.Split("|"c)(0) & """ ExtensionType=""ScriptExtension"" InheritedFromTagName=""""/>"
        Next

        xmlString &= "</ObjectExtension><AttributeExtension/></ExtensionInfo>"

        MxVal.PutString(xmlString)
        UDOUserAttrDataAttribute.SetValue(MxVal)

        newTemplate.Save()

        For Each Script In ScriptList
            If Script IsNot Nothing Then
                createNewFA(newTemplate, Script, {".Expression", ".ExecuteText", ".TriggerType"}, {".DeclarationsText", ".ExecuteText", ".Aliases", ".Expression", ".ExecuteText", ".TriggerType", ".TriggerPeriod", ".DataChangeDeadband", ".State.Historized", ".RunsAsync", ".ExecuteTimeout.Limit", ".ExecutionError.Alarmed", ".ExecutionError.Priority"})
            End If
        Next

        ' ---- Guardar y Check-In ----
        newTemplate.Save()
        newTemplate.CheckIn("Plantilla y atributos creados correctamente.")
    End Sub

    ''' <summary>
    ''' Crea y configura un nuevo Field Attribute (FA) en la plantilla dada.
    ''' </summary>
    ''' <param name="NewTemplate">La plantilla (ITemplate) donde se agregarán los atributos.</param>
    ''' <param name="FA">Objeto que representa el nuevo atributo a crear, contiene nombre y valores.</param>
    ''' <param name="TextBoxAttributes">Arreglo de sufijos de atributos que se configurarán con valores del atributo FA.</param>
    ''' <param name="GroupBoxBlockAttributes">Arreglo de sufijos de atributos que serán bloqueados (locked) tras su configuración.</param>
    ''' <remarks>
    ''' - Para cada sufijo en TextBoxAttributes, configura el atributo correspondiente concatenando el nombre base y el sufijo.
    ''' - Se realiza una conversión de tipos a través de DataTypeConversion para asignar valores correctamente.
    ''' - Configura atributos relacionados con Input/Output según el valor asignado.
    ''' - Para el atributo ".LevelAlarmed", activa algunos atributos específicos de alarma.
    ''' - Bloquea los atributos indicados en GroupBoxBlockAttributes para evitar modificaciones posteriores.
    ''' </remarks>
    Private Sub createNewFA(ByVal NewTemplate As aaGRAccessApp.ITemplate,
                                    ByVal FA As aaTemplateData.aaNewField,
                                    ByVal TextBoxAttributes As String(),
                                    ByVal GroupBoxBlockAttributes As String())

        Dim ValueList As New List(Of String)
        ValueList.AddRange(TextBoxAttributes)

        ' Configurar los atributos como de costumbre
        Dim DiscreteFieldAttribute As IAttribute
        Dim MxVal As New MxValueClass()
        Dim UDOAttributes As IAttributes = NewTemplate.Attributes
        Dim i As Integer = 0
        Debug.WriteLine(ValueList.Count & " y " & FA.Values.Count)
        For Each name In ValueList
            DiscreteFieldAttribute = UDOAttributes(FA.NameAttr.Split("|"c)(0) & name)
            DataTypeConversion(DiscreteFieldAttribute, FA.Values(i))

            If FA.Values(i) = "Output|3" Then
                DiscreteFieldAttribute = UDOAttributes(FA.NameAttr.Split("|"c)(0) & ".Output.OutputDest")

                DataTypeConversion(DiscreteFieldAttribute, "---|False")
            Else
                DiscreteFieldAttribute = UDOAttributes(FA.NameAttr.Split("|"c)(0) & ".Input.InputSource")

                DataTypeConversion(DiscreteFieldAttribute, "---|False")

            End If

            If name = ".LevelAlarmed" Then
                Dim Exceptions As String() = {".HiHi.Alarmed", ".Hi.Alarmed", ".Lo.Alarmed", ".LoLo.Alarmed"}
                For Each ex In Exceptions
                    DiscreteFieldAttribute = UDOAttributes(FA.NameAttr.Split("|"c)(0) & ex)
                    DataTypeConversion(DiscreteFieldAttribute, "True|False")
                Next
            End If
            i += 1
        Next

        For Each Name In GroupBoxBlockAttributes
            DiscreteFieldAttribute = UDOAttributes(FA.NameAttr.Split("|"c)(0) & Name)
            Try
                If DiscreteFieldAttribute IsNot Nothing Then
                    DiscreteFieldAttribute.SetLocked(MxPropertyLockedEnum.MxPropertyLockedEnumEND)
                End If
            Catch Ex As Exception
                frmMain.LogBox.Items.Add("Error general: " & Ex.Message)
            End Try
        Next
    End Sub

    ''' <summary>
    ''' Crea un nuevo User Defined Attribute (UDA) en la plantilla y le asigna un valor inicial.
    ''' </summary>
    ''' <param name="template">Plantilla donde se añadirá el UDA.</param>
    ''' <param name="UDAName">Nombre del nuevo UDA.</param>
    ''' <param name="UDAValue">Valor inicial que se asignará al UDA.</param>
    Private Sub createNewUDA(ByVal template As aaGRAccessApp.ITemplate, UDAName As String, UDAValue As String)
        Try
            ' Añadir el nuevo UDA
            template.AddUDA(
            UDAName,
            MxDataType.MxString,
            MxAttributeCategory.MxCategoryWriteable_USC_Lockable,
            MxSecurityClassification.MxSecurityOperate,
            False,
            0)

            template.Save()

            ' Configurar el valor del nuevo atributo
            Dim newAttribute As aaGRAccessApp.IAttribute = template.Attributes(UDAName)

            If newAttribute IsNot Nothing Then
                DataTypeConversion(newAttribute, UDAValue)
            Else
                frmMain.LogBox.Items.Add("No se encontró el atributo recién creado: " & UDAName)
            End If

        Catch ex As Exception
            frmMain.LogBox.Items.Add("Error: " & ex.Message)
        End Try
    End Sub


    ' ------- Data Conversion ----------
    'Cambiar esta y la proxima (Creacion de instancias necesita modificarse
    Private Sub ToMxValue(Atribute As String, gAttributes As aaGRAccess.IAttributes, Instance As aaGRAccessApp.IInstance, COMAtribute As String)
        Try
            Dim MXVal_ As aaGRAccess.MxValue = New aaGRAccess.MxValue()
            Dim Attr As aaGRAccess.IAttribute = gAttributes.Item(COMAtribute)

            If gAttributes.Item(COMAtribute) Is Nothing Then
                frmMain.LogBox.Items.Add("El atributo " & COMAtribute & " no existe en la instancia " & Instance.Tagname)
                Return
            End If
            MXVal_.PutString(Atribute)

            Attr.SetValue(MXVal_)
        Catch ex As Exception
            frmMain.LogBox.Items.Add(ex)
        End Try
    End Sub

    Private Sub ToMxValueEngUnits(Atribute As String, gAttributes As aaGRAccess.IAttributes, Instance As aaGRAccessApp.IInstance, COMAtribute As String)
        Try
            Dim MXVal_ As aaGRAccess.MxValue = New aaGRAccess.MxValue()
            Dim Attr As aaGRAccess.IAttribute = gAttributes.Item(COMAtribute & ".EngUnits")

            If gAttributes.Item(COMAtribute) Is Nothing Then
                frmMain.LogBox.Items.Add("El atributo " & COMAtribute & " no existe en la instancia " & Instance.Tagname)
                Return
            End If

            If Attr.Locked = aaGRAccess.MxPropertyLockedEnum.MxUnLocked Then
                MXVal_.PutString(Atribute)
                Attr.SetValue(MXVal_)
            Else
                frmMain.LogBox.Items.Add("El atributo " & COMAtribute & " está bloqueado en la instancia " & Instance.Tagname)
            End If

        Catch ex As Exception
            frmMain.LogBox.Items.Add(ex)
        End Try
    End Sub

    Private Function DataTypeGet(ByVal AttributeName As String, gAttributes As aaGRAccess.IAttributes)
        If String.IsNullOrWhiteSpace(AttributeName) Then
            Debug.WriteLine("AttributeName es null o vacío.")
            Return Nothing
        End If

        Try
            If gAttributes Is Nothing Then
                Debug.WriteLine("gAttributes es null.")
                Return Nothing
            End If

            Dim Attr As IAttribute = Nothing

            ' Proteger el acceso a gAttributes.Item
            Try
                Attr = gAttributes.Item(AttributeName)
            Catch ex As Exception
                MessageBox.Show($"Atributo '{AttributeName}' no encontrado en gAttributes: {ex.Message}")
                Return Nothing
            End Try

        If Attr Is Nothing Then Return Nothing

            Dim value = Attr.value
            If value Is Nothing Then Return Nothing

            Select Case Attr.DataType
                Case aaGRAccessApp.MxDataType.MxString,
                    aaGRAccessApp.MxDataType.MxInternationalizedString,
                    aaGRAccessApp.MxDataType.MxBigString
                    Return value.GetString()

                Case aaGRAccessApp.MxDataType.MxInteger
                    Return value.GetInteger()

                Case aaGRAccessApp.MxDataType.MxDouble
                    Return value.GetDouble()

                Case aaGRAccessApp.MxDataType.MxBoolean
                    Return value.GetBoolean()

                Case aaGRAccessApp.MxDataType.MxReferenceType
                    Return value.GetMxReference()

                Case Else
                    Debug.WriteLine($"Tipo de dato no soportado: {Attr.DataType}")
                    Return Nothing
            End Select

        Catch ex As Exception
            ' Protege si Attr es null antes de usar .Name
            Dim attrName As String = If(AttributeName, "Desconocido")
            MessageBox.Show("Error al leer el atributo: " & attrName & " - " & ex.Message)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Realiza la conversión del valor proporcionado a un tipo de dato compatible con el atributo COM recibido 
    ''' y lo asigna al mismo, aplicando bloqueo si es necesario.
    ''' </summary>
    ''' <param name="Attr">Objeto IAttribute que recibirá el valor convertido y el estado de bloqueo.</param>
    ''' <param name="ValueString">
    ''' Cadena de entrada con el formato "Valor|Locked", donde:
    ''' - "Valor" representa el contenido a convertir según el tipo de dato del atributo.
    ''' - "Locked" (opcional) puede ser "True" o "False" para indicar si el atributo debe ser bloqueado después.
    ''' </param>
    ''' <remarks>
    ''' Esta función maneja múltiples tipos de datos: MxString, MxInternationalizedString, MxInteger,
    ''' MxDouble, MxBoolean, MxReferenceType, MxQualifiedEnum y MxBigString.
    ''' También maneja errores silenciosamente mediante Debug.WriteLine.
    ''' </remarks>
    Private Sub DataTypeConversion(Attr As IAttribute, ValueString As String)
        Dim MxVal As New MxValueClass()
        Dim partes() As String = ValueString.Split("|"c)
        Dim Value As String = partes(0)
        Dim Locked As String = If(partes.Length > 1, partes(1), "False")

        If Locked <> "True" Then
            If Locked <> "False" Then
                Locked = "False"
            End If
        End If

        Try
            If Attr IsNot Nothing Then
                If Not Attr.Locked Then
                    Dim MxDataType As aaGRAccessApp.MxDataType = Attr.DataType
                    If MxDataType.ToString = "MxString" Then
                        MxVal.PutString(Value)
                    ElseIf MxDataType.ToString = "MxInternationalizedString" Then
                        MxVal.PutInternationalString(1033, Value)
                    ElseIf MxDataType.ToString = "MxInteger" Then
                        Dim intValue As Integer
                        If Not String.IsNullOrEmpty(Value) AndAlso Integer.TryParse(Value, intValue) Then
                            MxVal.PutInteger(intValue)
                        End If
                    ElseIf MxDataType.ToString = "MxDouble" Then
                        Dim doubleValue As Double
                        If Not String.IsNullOrEmpty(Value) AndAlso Double.TryParse(Value, doubleValue) Then
                            MxVal.PutDouble(doubleValue)
                        End If
                    ElseIf MxDataType.ToString = "MxBoolean" Then
                        Dim boolValue As Boolean
                        If Boolean.TryParse(Value, boolValue) Then
                            MxVal.PutBoolean(boolValue)
                        End If
                    ElseIf MxDataType.ToString = "MxReferenceType" Then
                        Dim MXRef As IMxReference
                        MXRef = Attr.value.GetMxReference()
                        MXRef.FullReferenceString = Value
                        MxVal.PutMxReference(MXRef)
                    ElseIf MxDataType.ToString = "MxQualifiedEnum" Then
                        Dim primitiveId As Short = 0             ' Generalmente 0 si no hay primitiva específica
                        Dim attributeId As Short = 0             ' Generalmente 0 si no hay un atributo específico
                        MxVal.PutCustomEnum(Value, ValueString.Split("|"c)(1), primitiveId, attributeId)
                    ElseIf MxDataType.ToString = "MxBigString" Then
                        MxVal.PutString(Value)
                    End If

                    Attr.SetValue(MxVal)

                    If Locked = "True" Then
                        Attr.SetLocked(MxPropertyLockedEnum.MxPropertyLockedEnumEND)
                    End If
                End If
            End If
        Catch e As Exception
            MessageBox.Show("Error en el atributo: " & Attr.Name & " Con el siguiente error generado " & e.Message)
        End Try
    End Sub
End Class