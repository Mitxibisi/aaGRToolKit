# DOCUMENTACIÓN TÉCNICA COMPLETA - aaGRToolKit

## ANÁLISIS FUNCIONAL DEL SISTEMA

### RESUMEN EJECUTIVO

**aaGRToolKit** es una herramienta especializada para la gestión de plantillas y datos en sistemas ArchestrA Galaxy, desarrollada como aplicación Windows Forms en VB.NET. Su arquitectura modular permite la extracción, manipulación y sincronización de datos entre Galaxy y Excel, facilitando la gestión masiva de configuraciones industriales.

---

## ARQUITECTURA DEL SISTEMA

### PATRÓN ARQUITECTURAL
El sistema implementa una **arquitectura en capas** con separación clara de responsabilidades:

```
┌─────────────────────────────────────────┐
│             FRONTEND LAYER              │
│  ┌─────────────┐ ┌──────────────────┐   │
│  │ frmMain.vb  │ │  CustomForms/    │   │
│  │ (Principal) │ │  (Diálogos)      │   │
│  └─────────────┘ └──────────────────┘   │
└─────────────────────────────────────────┘
┌─────────────────────────────────────────┐
│             BUSINESS LAYER              │
│  ┌─────────────────┐ ┌───────────────┐  │
│  │ frmMainFunctions│ │ TemplateData  │  │
│  │ (Lógica UI)     │ │ (Estructuras) │  │
│  └─────────────────┘ └───────────────┘  │
└─────────────────────────────────────────┘
┌─────────────────────────────────────────┐
│             DATA ACCESS LAYER           │
│  ┌─────────────────┐ ┌───────────────┐  │
│  │ aaGalaxyTools   │ │ aaExcelData   │  │
│  │ (Galaxy API)    │ │ (Excel I/O)   │  │
│  └─────────────────┘ └───────────────┘  │
└─────────────────────────────────────────┘
```

---

## FLUJOS DE TRABAJO PRINCIPALES

### 1. FLUJO DE CONEXIÓN A GALAXY

```mermaid
graph TD
    A[Usuario ingresa credenciales] --> B[aaGalaxyTools.login()]
    B --> C{Autenticación exitosa?}
    C -->|Sí| D[Obtener lista de plantillas]
    C -->|No| E[Mostrar error]
    D --> F[Habilitar funciones de exportación]
    E --> A
```

**Código relevante:**
```vb
' En frmMain.PantallasGalaxia.vb
Private Sub btnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click
    lblStatus.Text = "Logging in"
    refreshGalaxyInfo(cmboGalaxyList.SelectedValue)
    If aaGalaxyTools.login(txtUserInput.Text, txtPwdInput.Text) >= 0 Then
        lblStatus.Text = "Logged In " & cmboGalaxyList.Text & " Wich User " & txtUserInput.Text
    Else
        lblStatus.Text = "Error logging in"
    End If
End Sub
```

### 2. FLUJO DE EXTRACCIÓN DE PLANTILLAS

```mermaid
graph TD
    A[Seleccionar plantillas] --> B[Configurar carpeta destino]
    B --> C[Ejecutar exportación]
    C --> D[Para cada plantilla seleccionada]
    D --> E[getTemplateData()]
    E --> F[Procesar atributos discretos]
    F --> G[Procesar atributos analógicos]
    G --> H[Escribir línea CSV]
    H --> I{¿Más plantillas?}
    I -->|Sí| D
    I -->|No| J[Finalizar archivo CSV]
```

**Implementación en el código:**
```vb
' En frmMain.PantallasGalaxia.vb
Public Sub ExportTemplatesToFile(ByVal filePath As String, ByVal templateNames As String(), ByVal progressBar As ProgressBar)
    Dim outputFile As String = Path.Combine(filePath, "atributos_extraidos.csv")
    
    Using writer As New StreamWriter(outputFile, False, Encoding.UTF8)
        writer.WriteLine("Plantilla,Nombre,Plantilla derivada,Descripción,Historizado,Eventos,Alarm,Unidad")
        
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
            
            progressBar.PerformStep()
        Next
    End Using
End Sub
```

### 3. FLUJO DE IMPORTACIÓN DESDE EXCEL

```mermaid
graph TD
    A[Seleccionar archivo Excel] --> B[Configurar índices de columnas]
    B --> C[aaExcelData.CargarDatosMapeado()]
    C --> D[Leer hoja de cálculo]
    D --> E[Para cada instancia encontrada]
    E --> F[Extraer atributos individuales]
    F --> G[Extraer atributos array]
    G --> H[Crear aaInstanceData]
    H --> I{¿Más instancias?}
    I -->|Sí| E
    I -->|No| J[Retornar lista de instancias]
    J --> K[Generar controles UI dinámicos]
```

---

## COMPONENTES TÉCNICOS DETALLADOS

### FRONTEND LAYER

#### **frmMain.vb - Formulario Principal**
**Responsabilidad:** Coordinación general y gestión de estado
```vb
' Variables principales de estado
Public aaGalaxyTools As aaGalaxyTools          ' Conexión a Galaxy
Public aaExcelData As aaExcelData              ' Manipulación Excel
Public DiscreteTemplateData As List(Of TemplateData)  ' Datos discretos
Public AnalogTemplateData As List(Of TemplateData)    ' Datos analógicos
```

**Funciones clave:**
- Inicialización de componentes backend
- Gestión de configuraciones de usuario
- Coordinación entre módulos funcionales

#### **CustomForms/ - Diálogos Personalizados**

**CustomInputBox.vb:**
```vb
' Propiedades principales
Public Property ValorTexto As String      ' Texto ingresado
Public Property OpcionMarcada As Boolean  ' Estado del checkbox
```
- **Propósito:** Captura de entrada de usuario con opción adicional
- **Uso:** Configuración de parámetros dinámicos

**CustomListBox.vb:**
```vb
' Constructor que recibe arrays de opciones
Public Sub New(ItemsNames As String())
    InitializeComponent()
    ItemsList(ItemsNames)
End Sub
```
- **Propósito:** Selección de elementos de listas dinámicas
- **Uso:** Selección de plantillas, atributos, etc.

#### **frmMainFunctions/ - Módulos Funcionales**

**1. frmMain.PantallasGalaxia.vb**
- **Responsabilidad:** Gestión de conexión y exportación Galaxy
- **Funciones principales:**
  - `btnLogin_Click()`: Autenticación en Galaxy
  - `ExportTemplatesToFile()`: Exportación masiva de plantillas
  - `refreshGalaxyInfo()`: Actualización de información de Galaxy

**2. frmMain.PantallaPlantillas.vb**
- **Responsabilidad:** Creación y gestión de plantillas
- **Funciones principales:**
  - `btnNewTemplate_Click()`: Creación de nuevas plantillas
  - `btnTemplateDataAdd_Click()`: Carga de datos desde Excel
  - Procesamiento de configuraciones complejas

**3. frmMain.PantallaInstancias.vb**
- **Responsabilidad:** Gestión de instancias individuales
- **Funciones principales:**
  - `btnLoadData_Click()`: Carga masiva de instancias
  - `btnDeploy_Click()`: Despliegue de instancias
  - `btnUndeploy_Click()`: Retracción de instancias

### DATA ACCESS LAYER

#### **aaGalaxyTools.vb - Conectividad Galaxy**

**Patrón de diseño:** Singleton con gestión de estado
```vb
Private grAccess As aaGRAccessApp.GRAccessApp    ' Cliente COM
Private myGalaxy As aaGRAccessApp.IGalaxy        ' Instancia Galaxy
```

**Métodos principales:**

**1. Gestión de conexión:**
```vb
Public Function getGalaxies(ByVal NodeName As String) As Collection
    ' Consulta todas las galaxias disponibles en un nodo
    ' Retorna: Collection de nombres de Galaxy
```

**2. Autenticación:**
```vb
Public Function login(user As String, password As String) As Integer
    ' Autenticación en Galaxy seleccionada
    ' Retorna: Código de estado (>= 0 = éxito)
```

**3. Extracción de datos:**
```vb
Public Function getTemplateData(templateName As String) As aaTemplate
    ' Extrae todos los atributos de una plantilla específica
    ' Retorna: Objeto aaTemplate con atributos discretos y analógicos
```

**4. Gestión de instancias:**
```vb
Public Sub deployUndeployInstance(instanceName As String, cascade As Boolean, action As Integer)
    ' action: 0 = deploy, 1 = undeploy
    ' cascade: propagar a instancias dependientes
```

#### **aaExcelData.vb - Manipulación Excel**

**Patrón de diseño:** Factory con clases de datos especializadas

**Clases internas principales:**

**1. aaInstanceData:**
```vb
Public Class aaInstanceData
    Public InstanceName As String                          ' Nombre único
    Public InstanceTemplate As String                      ' Plantilla base
    Public InstanceAloneAttr As New List(Of String)        ' Atributos individuales
    Public InstanceArrayAttr As New List(Of List(Of String)) ' Atributos array
    Public EngUnitsArray As New List(Of String)            ' Unidades ingeniería
End Class
```

**2. aaTemplateDiscData:**
```vb
Public Class aaTemplateDiscData
    Public AttrName As String                    ' Nombre atributo
    Public AccessMode As String                  ' Modo acceso
    Public TextBoxAttr As New List(Of String)    ' Valores texto
    Public CheckBoxValue As New List(Of String)  ' Valores booleanos
    Public GroupBoxAttr As New List(Of String)   ' Valores agrupados
End Class
```

**Métodos principales:**

**1. Carga de datos mapeados:**
```vb
Public Function CargarDatosMapeado(FilPath As String, 
                                  InstanceNameColumnIndex As Integer,
                                  InstanceTemplateColumnIndex As Integer,
                                  EditableAloneElementsIndex As List(Of Integer),
                                  EditableArrayElementsIndex As List(Of Integer),
                                  EngUnits As Integer, 
                                  MapIndex As Integer) As List(Of aaInstanceData)
```
- **Propósito:** Importación masiva desde Excel con configuración flexible
- **Parámetros:** Índices de columnas configurables
- **Retorna:** Lista de instancias procesadas

**2. Generación de CSV:**
```vb
Public Sub GenerarCsv(FilPath As String, ...)
    ' Genera archivos CSV específicos para importación PLC
    ' Incluye transformaciones de formatos de datos
```

#### **aaTemplateData.vb - Estructuras de Datos**

**Patrón de diseño:** Data Transfer Objects (DTO) con serialización XML

**Clase principal:**
```vb
<Serializable()>
Public Class aaTemplate
    <XmlAttribute("name")> Public Name As String
    <XmlAttribute("revision")> Public Revision As Integer
    Public FieldAttributesDiscrete As New Collection()
    Public FieldAttributesAnalog As New Collection()
    
    Public Sub AddAttributes(ByVal otherTemplate As aaTemplate)
        ' Fusión de plantillas para configuraciones complejas
    End Sub
End Class
```

**Clases de atributos especializadas:**
```vb
<Serializable()>
Public Class aaFieldAttributeDiscrete
    Public Name As String              ' Nombre único
    Public TemplateName As String       ' Plantilla origen
    Public Description As String        ' Descripción funcional
    Public Historized As String         ' Configuración historización
    Public Events As String             ' Eventos asociados
    Public Alarm As Boolean             ' Estado alarma
    Public EngUnits As String           ' Unidades ingeniería
End Class
```

---

## CONFIGURACIÓN Y PERSISTENCIA

### SISTEMA DE CONFIGURACIÓN AVANZADO

El sistema utiliza **App.config** para configuraciones complejas con formato especializado:

#### **Formato de configuración de plantillas:**
```xml
<setting name="SettingPrueba" serializeAs="String">
    <value>Nombre|1||2|||;Descripcion|1||4,5|.Desc||;AccessMode|3|Input,InputOutput,Output|3|.AccessMode||</value>
</setting>
```

**Estructura del formato:**
```
NombreControl|TipoControl|OpcionesCombo|ColumnasExcel|AtributoGalaxy|SubAtributos|ColumnasAdicionales
```

**Tipos de control:**
- **1:** TextBox simple
- **2:** CheckBox
- **3:** ComboBox con opciones
- **4:** GroupBox con múltiples elementos

#### **Mapeo de columnas Excel:**
```xml
<!-- Índices configurables para mapeo flexible -->
<setting name="InstanceNameIndex" serializeAs="String">
    <value>8</value>  <!-- Columna H en Excel -->
</setting>
<setting name="InstanceTemplateIndex" serializeAs="String">
    <value>6</value>  <!-- Columna F en Excel -->
</setting>
```

---

## FLUJOS DE DATOS COMPLEJOS

### 1. PROCESAMIENTO DE ATRIBUTOS DINÁMICOS

```vb
' En frmMain.PantallaPlantillas.vb
Private Function CreateAttrData(container As Panel) As List(Of aaTemplateData.aaNewField)
    Dim fieldList As New List(Of aaTemplateData.aaNewField)
    
    For Each group As GroupBox In container.Controls.OfType(Of GroupBox)()
        ' Procesar controles dinámicos por tipo
        Dim textBoxes = group.Controls.OfType(Of TextBox)()
        Dim checkBoxes = group.Controls.OfType(Of CheckBox)()
        Dim comboBoxes = group.Controls.OfType(Of ComboBox)()
        
        ' Crear estructura de datos basada en configuración
        Dim newField As New aaTemplateData.aaNewField()
        ' ... lógica de procesamiento
        fieldList.Add(newField)
    Next
    
    Return fieldList
End Function
```

### 2. GENERACIÓN DINÁMICA DE INTERFAZ

```vb
' Creación de controles basada en configuración
Private Sub GenerarControlesDinamicos(configString As String)
    Dim parts() As String = configString.Split("|"c)
    Dim tipoControl As Integer = Integer.Parse(parts(1))
    
    Select Case tipoControl
        Case 1 ' TextBox
            Dim txt As New TextBox()
            txt.Name = "txt_" & parts(0)
            ' ... configuración específica
            
        Case 2 ' CheckBox
            Dim chk As New CheckBox()
            chk.Name = "chk_" & parts(0)
            ' ... configuración específica
            
        Case 3 ' ComboBox
            Dim cmb As New ComboBox()
            cmb.Items.AddRange(parts(2).Split(","c))
            ' ... configuración específica
    End Select
End Sub
```

---

## GESTIÓN DE ERRORES Y LOGGING

### ESTRATEGIA DE MANEJO DE ERRORES

**1. Errores de conectividad Galaxy:**
```vb
Try
    galaxies = grAccess.QueryGalaxies(galaxyNode)
    resultStatus = grAccess.CommandResult.Successful
    If Not resultStatus Then
        Throw New ApplicationException("No Galaxies detected on this node")
    End If
Catch e As Exception
    MessageBox.Show("Ocurrió un error al obtener las galaxias: " & e.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
End Try
```

**2. Errores de procesamiento Excel:**
```vb
Try
    ExcelPackage.LicenseContext = LicenseContext.NonCommercial
    package = New ExcelPackage(New FileInfo(filePath))
    worksheet = package.Workbook.Worksheets(0)
Catch e As Exception
    Console.WriteLine(e)
    LogBox.Items.Add("Error procesando Excel: " & e.Message)
End Try
```

**3. Sistema de logging integrado:**
```vb
' LogBox para feedback visual inmediato
Public LogBox As ListBox

' Uso en operaciones críticas
lblStatus.Text = "Logging in"
' ... operación
lblStatus.Text = "Logged In " & cmboGalaxyList.Text
```

---

## OPTIMIZACIONES Y RENDIMIENTO

### 1. GESTIÓN EFICIENTE DE MEMORIA

**Uso de `Using` para recursos:**
```vb
Using writer As New StreamWriter(outputFile, False, Encoding.UTF8)
    ' Procesamiento de archivo
    ' Liberación automática de recursos
End Using
```

**Limpieza de referencias COM:**
```vb
' Liberación explícita de objetos COM
Marshal.ReleaseComObject(grAccess)
Marshal.ReleaseComObject(myGalaxy)
```

### 2. PROCESAMIENTO ASÍNCRONO

**Barras de progreso para operaciones largas:**
```vb
progressBar.Minimum = 0
progressBar.Maximum = templateNames.Length
For Each templateName In templateNames
    ' ... procesamiento
    progressBar.PerformStep()
Next
progressBar.Value = 0
```

---

## EXTENSIBILIDAD Y MANTENIMIENTO

### PUNTOS DE EXTENSIÓN

**1. Nuevos tipos de atributos:**
- Agregar clases en `aaTemplateData.vb`
- Extender configuración en `App.config`
- Implementar procesamiento en `aaExcelData.vb`

**2. Nuevos formatos de exportación:**
- Extender `ExportTemplatesToFile()` con strategy pattern
- Agregar factories para diferentes formatos

**3. Conectividad adicional:**
- Implementar interfaces en `aaGalaxyTools.vb`
- Agregar abstracciones para diferentes sistemas

### CONSIDERACIONES DE MANTENIMIENTO

**1. Versionado de configuración:**
```vb
Public SchemaVersion As Double = 1.2
```

**2. Compatibilidad hacia atrás:**
- Migración automática de configuraciones
- Validación de esquemas XML

**3. Testing y debugging:**
- Logging extensivo en operaciones críticas
- Validación de datos en puntos clave
- Manejo graceful de errores

---

Esta documentación proporciona una visión técnica completa del sistema aaGRToolKit, incluyendo su arquitectura, flujos de trabajo, componentes técnicos y consideraciones de implementación.
