Imports System.Xml.Serialization

''' <summary>
''' Stores ArchestrA Galaxy template data in an XML serializable format.
''' </summary>
Public Module aaTemplateData

    ''' <summary>
    ''' The current XML schema version. This is useful in exports to determine 
    ''' if data may be missing from an older version export.
    ''' </summary>
    Public SchemaVersion As Double = 1.2

    ''' <summary>
    ''' The template itself. Contains scripts, field attributes, and user defined attributes.
    ''' </summary>
    <Serializable()>
    Public Class aaTemplate
        <XmlAttribute("name")> Public Name As String
        <XmlAttribute("revision")> Public Revision As Integer
        Public FieldAttributesDiscrete As New Collection()
        Public FieldAttributesAnalog As New Collection()

        Public Sub New()

        End Sub

        Public Sub New(ByVal Name As String, ByVal Revision As Integer, ByVal FieldAttributesDiscrete As Collection, ByVal FieldAttributesAnalog As Collection)
            Me.Name = Name
            Me.Revision = Revision
            Me.FieldAttributesDiscrete = FieldAttributesDiscrete
            Me.FieldAttributesAnalog = FieldAttributesAnalog
        End Sub

        ' Método para agregar los atributos de otra instancia de aaTemplate
        Public Sub AddAttributes(ByVal otherTemplate As aaTemplate)
            For Each item In otherTemplate.FieldAttributesDiscrete
                FieldAttributesDiscrete.Add(item)
            Next
            For Each item In otherTemplate.FieldAttributesAnalog
                FieldAttributesAnalog.Add(item)
            Next
        End Sub

    End Class


    ''' <summary>
    ''' Discrete Field Attribute data structure.
    ''' </summary>
    ''' <remarks></remarks>
    <Serializable()>
    Public Class aaFieldAttributeDiscrete
        Public Name As String
        Public TemplateName As String
        Public Description As String
        Public Historized As String
        Public Events As String
        Public Alarm As Boolean
        Public EngUnits As String

        Public Sub New()

        End Sub

        Public Sub New(ByVal Name As String, ByVal TemplateName As String,
                       ByVal Description As String, ByVal Historized As Boolean, ByVal Events As Boolean, ByVal Alarm As Boolean, ByVal EngUnits As String)

            Me.Name = Name
            Me.TemplateName = TemplateName
            Me.Description = Description
            Me.Historized = Historized
            Me.Events = Events
            Me.Alarm = Alarm
            Me.EngUnits = EngUnits

        End Sub

    End Class

    ''' <summary>
    ''' Analog Field Attribute data structure.
    ''' </summary>
    <Serializable()>
    Public Class aaFieldAttributeAnalog
        Public Name As String
        Public TemplateName As String
        Public Description As String
        Public Historized As String
        Public Events As String
        Public Alarm As Boolean
        Public EngUnits As String

        Public Sub New()

        End Sub

        Public Sub New(ByVal Name As String, ByVal TemplateName As String,
                       ByVal Description As String, ByVal Historized As Boolean, ByVal Events As Boolean, ByVal Alarm As Boolean, ByVal EngUnits As String)

            Me.Name = Name
            Me.TemplateName = TemplateName
            Me.Description = Description
            Me.Historized = Historized
            Me.Events = Events
            Me.Alarm = Alarm
            Me.EngUnits = EngUnits

        End Sub

    End Class

    <Serializable()>
    Public Class aaNewField
        Public NameAttr As String
        Public Values As New List(Of String)

        Public Sub New(ByVal NameAttr As String, ByVal Values As List(Of String))

            Me.NameAttr = NameAttr
            Me.Values = Values

        End Sub

    End Class

    Public Class FieldAttributesContainer
        Public Property DiscreteAttributes As List(Of aaFieldAttributeDiscrete)

        Public Property AnalogAttributes As List(Of aaFieldAttributeAnalog)

        Public Sub New()
            DiscreteAttributes = New List(Of aaFieldAttributeDiscrete)()
            AnalogAttributes = New List(Of aaFieldAttributeAnalog)()
        End Sub
    End Class
End Module