Imports ESRI.ArcGIS.Geodatabase
Public Class clsTransitLineAtts
    Private _transitLine As IFeature
    Public Sub New(ByVal TransitLineFeature As IFeature)
        _transitLine = TransitLineFeature
    End Sub
    Public ReadOnly Property OID() As Long
        Get
            'Dim intPos As Integer
            'intPos = modeAttributeRow.Fields.FindField("PSRCEDGEID")
            OID = _transitLine.Value(0)
        End Get
    End Property

    Public Property LineID() As Long
        Get
            Dim intPos As Integer
            intPos = _transitLine.Fields.FindField("LineID")
            LineID = _transitLine.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _transitLine.Fields.FindField("LineID")
            _transitLine.Value(intPos) = value
            _transitLine.Store()
        End Set
    End Property
    Public Property TimePeriod() As Short
        Get
            Dim intPos As Integer
            intPos = _transitLine.Fields.FindField("TimePeriod")
            TimePeriod = _transitLine.Value(intPos)
        End Get
        Set(ByVal value As Short)
            Dim intPos As Integer
            intPos = _transitLine.Fields.FindField("TimePeriod")
            _transitLine.Value(intPos) = value
            _transitLine.Store()
        End Set
    End Property
    Public Property TransitLineNo() As String
        Get
            Dim intPos As Integer
            intPos = _transitLine.Fields.FindField("TransLineN")
            TransitLineNo = _transitLine.Value(intPos)
        End Get
        Set(ByVal value As String)
            Dim intPos As Integer
            intPos = _transitLine.Fields.FindField("TransLineN")
            _transitLine.Value(intPos) = value
            _transitLine.Store()
        End Set
    End Property
    Public Property Mode() As String
        Get
            Dim intPos As Integer
            intPos = _transitLine.Fields.FindField("Mode")
            Mode = _transitLine.Value(intPos)
        End Get
        Set(ByVal value As String)
            Dim intPos As Integer
            intPos = _transitLine.Fields.FindField("Mode")
            _transitLine.Value(intPos) = value
            _transitLine.Store()
        End Set
    End Property
    Public Property VehicleType() As Short
        Get
            Dim intPos As Integer
            intPos = _transitLine.Fields.FindField("VehicleTyp")
            VehicleType = _transitLine.Value(intPos)
        End Get
        Set(ByVal value As Short)
            Dim intPos As Integer
            intPos = _transitLine.Fields.FindField("VehicleTyp")
            _transitLine.Value(intPos) = value
            _transitLine.Store()
        End Set
    End Property
    Public Property Headway() As Single
        Get
            Dim intPos As Integer
            intPos = _transitLine.Fields.FindField("Headway")
            Headway = _transitLine.Value(intPos)
        End Get
        Set(ByVal value As Single)
            Dim intPos As Integer
            intPos = _transitLine.Fields.FindField("Headway")
            _transitLine.Value(intPos) = value
            _transitLine.Store()
        End Set
    End Property
    Public Property Speed() As Single

        Get
            Dim intPos As Integer
            intPos = _transitLine.Fields.FindField("Speed")
            Speed = _transitLine.Value(intPos)
        End Get
        Set(ByVal value As Single)
            Dim intPos As Integer
            intPos = _transitLine.Fields.FindField("Speed")
            _transitLine.Value(intPos) = value
            _transitLine.Store()
        End Set
    End Property
    Public Property Description() As String
        Get
            Dim intPos As Integer
            intPos = _transitLine.Fields.FindField("Descriptio")
            Description = _transitLine.Value(intPos)
        End Get
        Set(ByVal value As String)
            Dim intPos As Integer
            intPos = _transitLine.Fields.FindField("Descriptio")
            _transitLine.Value(intPos) = value
            _transitLine.Store()
        End Set
    End Property
    Public Property Company() As Short
        Get
            Dim intPos As Integer
            intPos = _transitLine.Fields.FindField("Operator")
            Company = _transitLine.Value(intPos)
        End Get
        Set(ByVal value As Short)
            Dim intPos As Integer
            intPos = _transitLine.Fields.FindField("Operator")
            _transitLine.Value(intPos) = value
            _transitLine.Store()
        End Set
    End Property
    Public Property InServiceDate() As Short
        Get
            Dim intPos As Integer
            intPos = _transitLine.Fields.FindField("InServiceD")
            InServiceDate = _transitLine.Value(intPos)
        End Get
        Set(ByVal value As Short)
            Dim intPos As Integer
            intPos = _transitLine.Fields.FindField("InServiceD")
            _transitLine.Value(intPos) = value
            _transitLine.Store()
        End Set
    End Property
    Public Property OutServiceDate() As Short
        Get
            Dim intPos As Integer
            intPos = _transitLine.Fields.FindField("OutService")
            OutServiceDate = _transitLine.Value(intPos)
        End Get
        Set(ByVal value As Short)
            Dim intPos As Integer
            intPos = _transitLine.Fields.FindField("OutService")
            _transitLine.Value(intPos) = value
            _transitLine.Store()
        End Set
    End Property
    Public Property UL2() As Short
        Get
            Dim intPos As Integer
            intPos = _transitLine.Fields.FindField("UL2")
            UL2 = _transitLine.Value(intPos)
        End Get
        Set(ByVal value As Short)
            Dim intPos As Integer
            intPos = _transitLine.Fields.FindField("UL2")
            _transitLine.Value(intPos) = value
            _transitLine.Store()
        End Set
    End Property

    Public Property Headway_AM() As Double
        Get
            Dim intPos As Integer
            intPos = _transitLine.Fields.FindField("Headway_AM")
            Headway_AM = _transitLine.Value(intPos)
        End Get
        Set(ByVal value As Double)
            Dim intPos As Integer
            intPos = _transitLine.Fields.FindField("Headway_AM")
            _transitLine.Value(intPos) = value
            _transitLine.Store()
        End Set
    End Property
    Public Property Headway_MD() As Double
        Get
            Dim intPos As Integer
            intPos = _transitLine.Fields.FindField("Headway_MD")
            Headway_MD = _transitLine.Value(intPos)
        End Get
        Set(ByVal value As Double)
            Dim intPos As Integer
            intPos = _transitLine.Fields.FindField("Headway_MD")
            _transitLine.Value(intPos) = value
            _transitLine.Store()
        End Set
    End Property
    Public Property Headway_PM() As Double
        Get
            Dim intPos As Integer
            intPos = _transitLine.Fields.FindField("Headway_PM")
            Headway_PM = _transitLine.Value(intPos)
        End Get
        Set(ByVal value As Double)
            Dim intPos As Integer
            intPos = _transitLine.Fields.FindField("Headway_PM")
            _transitLine.Value(intPos) = value
            _transitLine.Store()
        End Set
    End Property
    Public Property Headway_EV() As Double
        Get
            Dim intPos As Integer
            intPos = _transitLine.Fields.FindField("Headway_EV")
            Headway_EV = _transitLine.Value(intPos)
        End Get
        Set(ByVal value As Double)
            Dim intPos As Integer
            intPos = _transitLine.Fields.FindField("Headway_EV")
            _transitLine.Value(intPos) = value
            _transitLine.Store()
        End Set
    End Property

End Class
