Imports ESRI.ArcGIS.Geodatabase
Public Class clsBaseEdgeFeatAttributes
    Private _edge As IFeature
    Public Sub New(ByVal EdgeFeature As IFeature)
        _edge = EdgeFeature
    End Sub
    Public ReadOnly Property OID() As Long
        Get
            'Dim intPos As Integer
            'intPos = modeAttributeRow.Fields.FindField("PSRCEDGEID")
            OID = _edge.Value(0)
        End Get
    End Property
    Public ReadOnly Property Enabled() As Boolean
        Get
            Dim intPos As Integer
            intPos = _edge.Fields.FindField("Enabled")
            Enabled = _edge.Value(intPos)
        End Get

    End Property
    Public ReadOnly Property PSRCEdgeID() As Long
        Get
            Dim intPos As Integer
            intPos = _edge.Fields.FindField("PSRCEdgeID")
            PSRCEdgeID = _edge.Value(intPos)
        End Get

    End Property
    Public Property FTRsegID() As Long
        Get
            Dim intPos As Integer
            intPos = _edge.Fields.FindField("FTRsegID")
            FTRsegID = _edge.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _edge.Fields.FindField("FTRsegID")
            _edge.Value(intPos) = value
            _edge.Store()
        End Set
    End Property
    Public ReadOnly Property FacilityType() As String
        Get
            Dim intPos As Integer
            intPos = _edge.Fields.FindField("FacilityType")
            FacilityType = _edge.Value(intPos)
        End Get

    End Property
    Public ReadOnly Property NewFacilityType() As Long
        Get
            Dim intPos As Integer
            intPos = _edge.Fields.FindField("NewFacilityType")
            NewFacilityType = _edge.Value(intPos)
        End Get

    End Property
    Public ReadOnly Property INode() As Long
        Get
            Dim intPos As Integer
            intPos = _edge.Fields.FindField("INode")
            INode = _edge.Value(intPos)
        End Get

    End Property
    Public ReadOnly Property JNode() As Long
        Get
            Dim intPos As Integer
            intPos = _edge.Fields.FindField("JNode")
            JNode = _edge.Value(intPos)
        End Get

    End Property
    Public ReadOnly Property InServiceDate() As Long
        Get
            Dim intPos As Integer
            intPos = _edge.Fields.FindField("InServiceDate")
            InServiceDate = _edge.Value(intPos)
        End Get

    End Property
    Public ReadOnly Property OutServiceDate() As Long
        Get
            Dim intPos As Integer
            intPos = _edge.Fields.FindField("OutServiceDate")
            OutServiceDate = _edge.Value(intPos)
        End Get

    End Property
    Public ReadOnly Property DateLastUpdated() As Long
        Get
            Dim intPos As Integer
            intPos = _edge.Fields.FindField("DateLastUpdated")
            DateLastUpdated = _edge.Value(intPos)
        End Get

    End Property
    Public ReadOnly Property ActiveLink() As Short
        Get
            Dim intPos As Integer
            intPos = _edge.Fields.FindField("ActiveLink")
            ActiveLink = _edge.Value(intPos)
        End Get

    End Property

    Public ReadOnly Property LinkType() As Short
        Get
            Dim intPos As Integer
            intPos = _edge.Fields.FindField("LinkType")
            LinkType = _edge.Value(intPos)
        End Get

    End Property
    Public ReadOnly Property FunctionalClass() As String
        Get
            Dim intPos As Integer
            intPos = _edge.Fields.FindField("FunctionalClass")
            FunctionalClass = _edge.Value(intPos)
        End Get

    End Property
    Public ReadOnly Property Modes() As String
        Get
            Dim intPos As Integer
            intPos = _edge.Fields.FindField("Modes")
            Modes = _edge.Value(intPos)
        End Get

    End Property
    Public ReadOnly Property MTS() As String
        Get
            Dim intPos As Integer
            intPos = _edge.Fields.FindField("MTS")
            MTS = _edge.Value(intPos)
        End Get

    End Property
    Public ReadOnly Property CMSlinkID() As String
        Get
            Dim intPos As Integer
            intPos = _edge.Fields.FindField("CMSlinkID")
            CMSlinkID = _edge.Value(intPos)
        End Get

    End Property
    Public ReadOnly Property CMScriticalLinkID() As String
        Get
            Dim intPos As Integer
            intPos = _edge.Fields.FindField("CMScriticalLinkID")
            CMScriticalLinkID = _edge.Value(intPos)
        End Get

    End Property
    Public ReadOnly Property Oneway() As Short
        Get
            Dim intPos As Integer
            intPos = _edge.Fields.FindField("Oneway")
            Oneway = _edge.Value(intPos)
        End Get

    End Property
    Public ReadOnly Property Fullname() As String
        Get
            Dim intPos As Integer
            intPos = _edge.Fields.FindField("Fullname")
            Fullname = _edge.Value(intPos)
        End Get

    End Property
    Public ReadOnly Property DateCreated() As Long
        Get
            Dim intPos As Integer
            intPos = _edge.Fields.FindField("DateCreated")
            DateCreated = _edge.Value(intPos)
        End Get

    End Property
    Public ReadOnly Property LastEditor() As String
        Get
            Dim intPos As Integer
            intPos = _edge.Fields.FindField("LastEditor")
            LastEditor = _edge.Value(intPos)
        End Get

    End Property
    Public ReadOnly Property EditNotes() As String
        Get
            Dim intPos As Integer
            intPos = _edge.Fields.FindField("EditNotes")
            EditNotes = _edge.Value(intPos)
        End Get

    End Property
    Public ReadOnly Property Tonnage() As Long
        Get
            Dim intPos As Integer
            intPos = _edge.Fields.FindField("Tonnage")
            Tonnage = _edge.Value(intPos)
        End Get

    End Property
    Public ReadOnly Property GlobalEID() As Long
        Get
            Dim intPos As Integer
            intPos = _edge.Fields.FindField("GlobalEID")
            GlobalEID = _edge.Value(intPos)
        End Get

    End Property
    Public ReadOnly Property Processing() As Long
        Get
            Dim intPos As Integer
            intPos = _edge.Fields.FindField("Processing")
            Processing = _edge.Value(intPos)
        End Get

    End Property
    Public ReadOnly Property CountID() As Long
        Get
            Dim intPos As Integer
            intPos = _edge.Fields.FindField("CountID")
            CountID = _edge.Value(intPos)
        End Get

    End Property
    Public ReadOnly Property CountyID() As Long
        Get
            Dim intPos As Integer
            intPos = _edge.Fields.FindField("CountyID")
            CountyID = _edge.Value(intPos)
        End Get

    End Property
 
End Class
