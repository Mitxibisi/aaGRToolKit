# REFERENCIA TÉCNICA DE APIs - aaGRToolKit

## REFERENCIA COMPLETA DE APIS Y ESTRUCTURAS

### CLASE: aaGalaxyTools

**Propósito:** Interfaz principal para conectividad y operaciones con Galaxy ArchestrA

#### PROPIEDADES PÚBLICAS

```vb
Public showLogin As Boolean          ' Indica si mostrar diálogo de login
Public errorMessage As String        ' Último mensaje de error
Public loggedIn As Boolean          ' Estado de autenticación actual
```

#### MÉTODOS PRINCIPALES

##### **getGalaxies(NodeName As String) As Collection**
```vb
''' <summary>
''' Consulta todas las galaxias disponibles en un nodo específico
''' </summary>
''' <param name="NodeName">Nombre del nodo (ej: "localhost")</param>
''' <returns>Collection de nombres de galaxias como strings</returns>
''' <exception cref="ApplicationException">Si no se detectan galaxias</exception>
Public Function getGalaxies(ByVal NodeName As String) As Collection
```
**Uso típico:**
```vb
Dim galaxias = aaGalaxyTools.getGalaxies("localhost")
For Each galaxyName In galaxias
    Console.WriteLine($"Galaxy disponible: {galaxyName}")
Next
```

##### **login(user As String, password As String) As Integer**
```vb
''' <summary>
''' Autentica usuario en la galaxy seleccionada
''' </summary>
''' <param name="user">Nombre de usuario</param>
''' <param name="password">Contraseña</param>
''' <returns>Código de estado: >= 0 éxito, < 0 error</returns>
Public Function login(user As String, password As String) As Integer
```

##### **getTemplates(Optional hideBase As CheckState = CheckState.Unchecked) As Collection**
```vb
''' <summary>
''' Obtiene lista de plantillas disponibles en la galaxy
''' </summary>
''' <param name="hideBase">Si ocultar plantillas base del sistema</param>
''' <returns>Collection de nombres de plantillas</returns>
Public Function getTemplates(Optional hideBase As CheckState = CheckState.Unchecked) As Collection
```

##### **getTemplateData(templateName As String) As aaTemplate**
```vb
''' <summary>
''' Extrae todos los datos de una plantilla específica
''' </summary>
''' <param name="templateName">Nombre de la plantilla a consultar</param>
''' <returns>Objeto aaTemplate con atributos completos</returns>
Public Function getTemplateData(templateName As String) As aaTemplate
```

##### **deployUndeployInstance(instanceName As String, cascade As Boolean, action As Integer)**
```vb
''' <summary>
''' Despliega o retrae una instancia en Galaxy
''' </summary>
''' <param name="instanceName">Nombre de la instancia</param>
''' <param name="cascade">Si propagar a instancias dependientes</param>
''' <param name="action">0 = deploy, 1 = undeploy</param>
Public Sub deployUndeployInstance(instanceName As String, cascade As Boolean, action As Integer)
```

##### **CreateTemplate(derivedTemplate As String, newTemplateName As String, ...)**
```vb
''' <summary>
''' Crea nueva plantilla en Galaxy con atributos especificados
''' </summary>
''' <param name="derivedTemplate">Plantilla base de la cual derivar</param>
''' <param name="newTemplateName">Nombre de la nueva plantilla</param>
''' <param name="discreteAttrs">Lista de atributos discretos</param>
''' <param name="analogAttrs">Lista de atributos analógicos</param>
''' <param name="scriptAttrs">Lista de scripts</param>
''' <param name="udaAttrs">Lista de UDAs</param>
Public Sub CreateTemplate(derivedTemplate As String, 
                         newTemplateName As String,
                         discreteAttrs As List(Of aaTemplateData.aaNewField),
                         analogAttrs As List(Of aaTemplateData.aaNewField),
                         scriptAttrs As List(Of aaTemplateData.aaNewField),
                         udaAttrs As List(Of aaTemplateData.aaNewField))
```

##### **ExportSymbol(SymbolName As String)**
```vb
''' <summary>
''' Exporta símbolo gráfico a archivo XML
''' </summary>
''' <param name="SymbolName">Nombre del símbolo en Galaxy</param>
Public Sub ExportSymbol(SymbolName As String)
```

##### **ImportSymbol(SymbolName As String, folderPath As String, OverWrite As Boolean)**
```vb
''' <summary>
''' Importa símbolo gráfico desde archivo XML
''' </summary>
''' <param name="SymbolName">Nombre del símbolo</param>
''' <param name="folderPath">Ruta del archivo XML</param>
''' <param name="OverWrite">Si sobrescribir símbolo existente</param>
Public Sub ImportSymbol(SymbolName As String, folderPath As String, OverWrite As Boolean)
```

---

### CLASE: aaExcelData

**Propósito:** Manipulación y procesamiento de datos Excel

#### CLASES ANIDADAS

##### **aaInstanceData**
```vb
<Serializable()>
Public Class aaInstanceData
    Public InstanceName As String                          ' Nombre único de instancia
    Public InstanceTemplate As String                      ' Plantilla asociada
    Public InstanceAloneAttr As New List(Of String)        ' Atributos individuales
    Public InstanceArrayAttr As New List(Of List(Of String)) ' Atributos array
    Public EngUnitsArray As New List(Of String)            ' Unidades de ingeniería
    
    ' Constructor completo
    Public Sub New(InstanceName As String, InstanceTemplate As String, 
                   InstanceArrayAttr As List(Of List(Of String)), 
                   InstanceAloneAttr As List(Of String), 
                   EngsUnitsArray As List(Of String))
End Class
```

##### **aaTemplateDiscData**
```vb
<Serializable()>
Public Class aaTemplateDiscData
    Public AttrName As String                              ' Nombre del atributo
    Public AccessMode As String                            ' Modo de acceso
    Public TextBoxAttr As New List(Of String)              ' Valores TextBox
    Public CheckBoxValue As New List(Of String)            ' Valores CheckBox
    Public GroupBoxAttr As New List(Of String)             ' Valores GroupBox
End Class
```

##### **aaTemplateAnalogData**
```vb
<Serializable()>
Public Class aaTemplateAnalogData
    Public AttrName As String                              ' Nombre del atributo
    Public AccessMode As String                            ' Modo de acceso
    Public TextBoxAttr As New List(Of String)              ' Valores TextBox
    Public CheckBoxValue As New List(Of String)            ' Valores CheckBox
    Public GroupBoxAttr As New List(Of String)             ' Valores GroupBox
End Class
```

#### MÉTODOS PRINCIPALES

##### **CargarDatosMapeado(...) As List(Of aaInstanceData)**
```vb
''' <summary>
''' Carga datos de instancias desde Excel con mapeo configurable
''' </summary>
''' <param name="FilPath">Ruta del archivo Excel</param>
''' <param name="InstanceNameColumnIndex">Índice columna nombre instancia</param>
''' <param name="InstanceTemplateColumnIndex">Índice columna plantilla</param>
''' <param name="EditableAloneElementsIndex">Índices atributos individuales</param>
''' <param name="EditableArrayElementsIndex">Índices atributos array</param>
''' <param name="EngUnits">Índice columna unidades ingeniería</param>
''' <param name="MapIndex">Índice columna mapeo</param>
''' <returns>Lista de instancias procesadas</returns>
Public Function CargarDatosMapeado(FilPath As String, 
                                  InstanceNameColumnIndex As Integer,
                                  InstanceTemplateColumnIndex As Integer,
                                  EditableAloneElementsIndex As List(Of Integer),
                                  EditableArrayElementsIndex As List(Of Integer),
                                  EngUnits As Integer, 
                                  MapIndex As Integer) As List(Of aaInstanceData)
```

##### **CargarDatosPlantilla(...) As aaTemplateData**
```vb
''' <summary>
''' Carga configuración completa de plantilla desde Excel
''' </summary>
''' <param name="FilPath">Ruta archivo Excel</param>
''' <param name="DiscreteList1-4">Arrays índices para datos discretos</param>
''' <param name="AnalogList1-4">Arrays índices para datos analógicos</param>
''' <param name="UDAList1">Array índices para UDAs</param>
''' <param name="ScriptList1">Array índices para Scripts</param>
''' <returns>Objeto aaTemplateData con configuración completa</returns>
Public Function CargarDatosPlantilla(FilPath As String,
                                    DiscreteList1 As Integer(), DiscreteList2 As Integer(), 
                                    DiscreteList3 As Integer(), DiscreteList4 As Integer(),
                                    AnalogList1 As Integer(), AnalogList2 As Integer(), 
                                    AnalogList3 As Integer(), AnalogList4 As Integer(),
                                    UDAList1 As Integer(),
                                    ScriptList1 As Integer()) As aaTemplateData
```

##### **GenerarCsv(...)**
```vb
''' <summary>
''' Genera archivos CSV para importación PLC con transformaciones
''' </summary>
''' <param name="FilPath">Ruta archivo Excel origen</param>
''' <param name="InstanceNameColumnIndex">Índice columna nombres</param>
''' <param name="InstanceTemplateColumnIndex">Índice columna plantillas</param>
''' <param name="csvIndex">Índice columna datos CSV</param>
''' <param name="csv1index">Índice columna adicional CSV</param>
''' <param name="generateCSV">Si generar archivo CSV</param>
''' <param name="generateDoubleCSV">Si generar CSV doble</param>
''' <param name="DoubleText1">Texto filtro 1</param>
''' <param name="DoubleText2">Texto filtro 2</param>
Public Sub GenerarCsv(FilPath As String, 
                     InstanceNameColumnIndex As Integer,
                     InstanceTemplateColumnIndex As Integer,
                     csvIndex As String, csv1index As String,
                     generateCSV As Boolean, generateDoubleCSV As Boolean,
                     DoubleText1 As String, DoubleText2 As String)
```

---

### MÓDULO: aaTemplateData

**Propósito:** Estructuras de datos serializables para plantillas

#### CLASES PRINCIPALES

##### **aaTemplate**
```vb
<Serializable()>
Public Class aaTemplate
    <XmlAttribute("name")> Public Name As String           ' Nombre plantilla
    <XmlAttribute("revision")> Public Revision As Integer  ' Número revisión
    Public FieldAttributesDiscrete As New Collection()     ' Atributos discretos
    Public FieldAttributesAnalog As New Collection()       ' Atributos analógicos
    
    ''' <summary>
    ''' Fusiona atributos de otra plantilla
    ''' </summary>
    Public Sub AddAttributes(ByVal otherTemplate As aaTemplate)
End Class
```

##### **aaFieldAttributeDiscrete**
```vb
<Serializable()>
Public Class aaFieldAttributeDiscrete
    Public Name As String               ' Nombre atributo
    Public TemplateName As String        ' Plantilla origen
    Public Description As String         ' Descripción
    Public Historized As String          ' Configuración historización
    Public Events As String              ' Eventos asociados
    Public Alarm As Boolean              ' Estado alarma
    Public EngUnits As String            ' Unidades ingeniería
    Public AccessMode As String          ' Modo acceso
    Public Category As String            ' Categoría
    Public DataType As String            ' Tipo de dato
    Public InitialValue As String        ' Valor inicial
    Public MinValue As String            ' Valor mínimo
    Public MaxValue As String            ' Valor máximo
End Class
```

##### **aaFieldAttributeAnalog**
```vb
<Serializable()>
Public Class aaFieldAttributeAnalog
    Public Name As String               ' Nombre atributo
    Public TemplateName As String        ' Plantilla origen
    Public Description As String         ' Descripción
    Public Historized As String          ' Configuración historización
    Public Events As String              ' Eventos asociados
    Public Alarm As Boolean              ' Estado alarma
    Public EngUnits As String            ' Unidades ingeniería
    Public AccessMode As String          ' Modo acceso
    Public DataType As String            ' Tipo de dato (Integer, Float, Double)
    Public InitialValue As String        ' Valor inicial
    Public MinValue As String            ' Valor mínimo
    Public MaxValue As String            ' Valor máximo
    Public Scaling As String             ' Configuración escalado
    Public AlarmLimits As String         ' Límites de alarma
End Class
```

##### **aaNewField**
```vb
''' <summary>
''' Estructura para nuevos campos en creación de plantillas
''' </summary>
Public Class aaNewField
    Public FieldName As String          ' Nombre del campo
    Public FieldType As String          ' Tipo de campo
    Public FieldValue As String         ' Valor del campo
    Public FieldDescription As String   ' Descripción del campo
End Class
```

---

### FORMULARIOS PERSONALIZADOS

#### **CustomInputBox**
```vb
Public Class CustomInputBox
    Public Property ValorTexto As String        ' Texto ingresado
    Public Property OpcionMarcada As Boolean    ' Estado checkbox
    
    ''' <summary>
    ''' Muestra diálogo y retorna DialogResult
    ''' </summary>
    Public Function ShowDialog() As DialogResult
End Class
```

**Uso típico:**
```vb
Dim inputBox As New CustomInputBox()
If inputBox.ShowDialog() = DialogResult.OK Then
    Dim textoIngresado = inputBox.ValorTexto
    Dim opcionSeleccionada = inputBox.OpcionMarcada
End If
```

#### **CustomListBox**
```vb
Public Class CustomListBox
    Public Property ValorTexto As String        ' Item seleccionado
    Public Property OpcionMarcada As Boolean    ' Estado selección
    
    ''' <summary>
    ''' Constructor con items predefinidos
    ''' </summary>
    Public Sub New(ItemsNames As String())
End Class
```

**Uso típico:**
```vb
Dim opciones() As String = {"Opción 1", "Opción 2", "Opción 3"}
Dim listBox As New CustomListBox(opciones)
If listBox.ShowDialog() = DialogResult.OK Then
    Dim opcionSeleccionada = listBox.ValorTexto
End If
```

---

## CONFIGURACIONES AVANZADAS

### FORMATO DE CONFIGURACIÓN DE CONTROLES

**Estructura general:**
```
NombreControl|TipoControl|OpcionesCombo|ColumnasExcel|AtributoGalaxy|SubAtributos|ColumnasAdicionales
```

**Tipos de control:**
- **1:** TextBox simple
- **2:** CheckBox
- **3:** ComboBox con opciones predefinidas
- **4:** GroupBox con múltiples sub-controles

**Ejemplos de configuración:**

```xml
<!-- TextBox simple -->
Nombre|1||2|||

<!-- CheckBox -->
Historizado|2||12,13|.Historized||

<!-- ComboBox con opciones -->
AccessMode|3|Input,InputOutput,Output|3|.AccessMode||

<!-- GroupBox complejo -->
IOScaling|4|RawMin,RawMax,EUMin,EUMax|10,11|.Scaled|.RawMin,.RawMax,.EngUnitsMin,.EngUnitsMax|12,15,13,16
```

### ÍNDICES DE COLUMNA EXCEL

**Configuraciones estándar:**
```xml
<setting name="InstanceNameIndex" serializeAs="String">
    <value>8</value>  <!-- Columna H -->
</setting>
<setting name="InstanceTemplateIndex" serializeAs="String">
    <value>6</value>  <!-- Columna F -->
</setting>
<setting name="InstanceEngUnits1" serializeAs="String">
    <value>7</value>  <!-- Columna G -->
</setting>
```

### ATRIBUTOS BLOQUEADOS

**Analógicos:**
```
.ConversionMode, .HiHi.Priority, .Hi.Priority, .Lo.Priority, .LoLo.Priority, 
.HiHi.DescAttrName, .Hi.DescAttrName, .Lo.DescAttrName, .LoLo.DescAttrName, 
.LevelAlarm.ValueDeadband, .LevelAlarm.TimeDeadband
```

**Discretos:**
```
.ActiveAlarmState, .DescAttrName, .BooleanAlarm.TimeDeadBand, .Category
```

---

## CÓDIGOS DE ERROR Y ESTADOS

### CÓDIGOS DE RETORNO LOGIN
- **>= 0:** Autenticación exitosa
- **-1:** Credenciales inválidas
- **-2:** Galaxy no disponible
- **-3:** Error de conectividad
- **-4:** Permisos insuficientes

### ESTADOS DE PROCESAMIENTO EXCEL
- **Success:** Procesamiento completado
- **FileNotFound:** Archivo Excel no encontrado
- **InvalidFormat:** Formato de archivo inválido
- **ColumnMismatch:** Índices de columna incorrectos
- **DataCorrupted:** Datos corruptos o incompletos

### CÓDIGOS GALAXY OPERATIONS
- **0:** Deploy exitoso
- **1:** Undeploy exitoso
- **-1:** Instancia no encontrada
- **-2:** Permisos insuficientes
- **-3:** Dependencias no resueltas

---

Esta referencia técnica proporciona documentación completa de todas las APIs, estructuras de datos y configuraciones disponibles en el sistema aaGRToolKit.
