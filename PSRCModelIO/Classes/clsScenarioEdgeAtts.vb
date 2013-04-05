Imports ESRI.ArcGIS.Geodatabase
Public Class clsScenarioEdgeAtts
    Inherits clsBaseEdgeFeatAttributes
    Private _scenarioEdge As IFeature
    Public Sub New(ByVal ScenarioEdgeFeature As IFeature)
        MyBase.New(ScenarioEdgeFeature)
        _scenarioEdge = ScenarioEdgeFeature
    End Sub

    Public Property PSRC_E2ID() As Long
        Get
            Dim intPos As Integer
            intPos = _scenarioEdge.Fields.FindField("PSRC_E2ID")
            PSRC_E2ID = _scenarioEdge.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _scenarioEdge.Fields.FindField("PSRC_E2ID")
            _scenarioEdge.Value(intPos) = value
            _scenarioEdge.Store()
        End Set
    End Property
    Public Property ScenarioID() As Long
        Get
            Dim intPos As Integer
            intPos = _scenarioEdge.Fields.FindField("ScenarioID")
            ScenarioID = _scenarioEdge.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _scenarioEdge.Fields.FindField("ScenarioID")
            _scenarioEdge.Value(intPos) = value
            _scenarioEdge.Store()
        End Set
    End Property
    Public Property Scen_Link() As Long
        Get
            Dim intPos As Integer
            intPos = _scenarioEdge.Fields.FindField("Scen_Link")
            Scen_Link = _scenarioEdge.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _scenarioEdge.Fields.FindField("Scen_Link")
            _scenarioEdge.Value(intPos) = value
            _scenarioEdge.Store()
        End Set
    End Property
    Public Property TR_I() As Long
        Get
            Dim intPos As Integer
            intPos = _scenarioEdge.Fields.FindField("TR_I")
            TR_I = _scenarioEdge.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _scenarioEdge.Fields.FindField("TR_I")
            _scenarioEdge.Value(intPos) = value
            _scenarioEdge.Store()
        End Set
    End Property
    Public Property TR_J() As Long
        Get
            Dim intPos As Integer
            intPos = _scenarioEdge.Fields.FindField("TR_J")
            TR_J = _scenarioEdge.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _scenarioEdge.Fields.FindField("TR_J")
            _scenarioEdge.Value(intPos) = value
            _scenarioEdge.Store()
        End Set
    End Property
    Public Property HOV_I() As Long
        Get
            Dim intPos As Integer
            intPos = _scenarioEdge.Fields.FindField("HOV_I")
            HOV_I = _scenarioEdge.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _scenarioEdge.Fields.FindField("HOV_I")
            _scenarioEdge.Value(intPos) = value
            _scenarioEdge.Store()
        End Set
    End Property
    Public Property HOV_J() As Long
        Get
            Dim intPos As Integer
            intPos = _scenarioEdge.Fields.FindField("HOV_J")
            HOV_J = _scenarioEdge.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _scenarioEdge.Fields.FindField("HOV_J")
            _scenarioEdge.Value(intPos) = value
            _scenarioEdge.Store()
        End Set
    End Property
    Public Property TK_I() As Long
        Get
            Dim intPos As Integer
            intPos = _scenarioEdge.Fields.FindField("TK_I")
            TK_I = _scenarioEdge.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _scenarioEdge.Fields.FindField("TK_I")
            _scenarioEdge.Value(intPos) = value
            _scenarioEdge.Store()
        End Set
    End Property
    Public Property TK_J() As Long
        Get
            Dim intPos As Integer
            intPos = _scenarioEdge.Fields.FindField("TK_J")
            TK_J = _scenarioEdge.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _scenarioEdge.Fields.FindField("TK_J")
            _scenarioEdge.Value(intPos) = value
            _scenarioEdge.Store()
        End Set
    End Property
    Public Property UseEmmeN() As Long
        Get
            Dim intPos As Integer
            intPos = _scenarioEdge.Fields.FindField("UseEmmeN")
            UseEmmeN = _scenarioEdge.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _scenarioEdge.Fields.FindField("UseEmmeN")
            _scenarioEdge.Value(intPos) = value
            _scenarioEdge.Store()
        End Set
    End Property
    Public Property SplitHOV() As String
        Get

            Dim intPos As Integer
            intPos = _scenarioEdge.Fields.FindField("SplitHOV")
            If IsDBNull(_scenarioEdge.Value(intPos)) Then
                SplitHOV = "Null"
            Else
                SplitHOV = _scenarioEdge.Value(intPos)
            End If


        End Get
        Set(ByVal value As String)
            Dim intPos As Integer
            intPos = _scenarioEdge.Fields.FindField("SplitHOV")
            _scenarioEdge.Value(intPos) = value
            _scenarioEdge.Store()
        End Set
    End Property
    Public Property SplitTR() As String
        Get

            Dim intPos As Integer
            intPos = _scenarioEdge.Fields.FindField("SplitTR")
            If IsDBNull(_scenarioEdge.Value(intPos)) Then
                SplitTR = "Null"
            Else
                SplitTR = _scenarioEdge.Value(intPos)
            End If


        End Get
        Set(ByVal value As String)
            Dim intPos As Integer
            intPos = _scenarioEdge.Fields.FindField("SplitTR")
            _scenarioEdge.Value(intPos) = value
            _scenarioEdge.Store()
        End Set
    End Property
    Public Property SplitTK() As String
        Get

            Dim intPos As Integer
            intPos = _scenarioEdge.Fields.FindField("SplitTK")
            If IsDBNull(_scenarioEdge.Value(intPos)) Then
                SplitTK = "Null"
            Else
                SplitTK = _scenarioEdge.Value(intPos)
            End If


        End Get
        Set(ByVal value As String)
            Dim intPos As Integer
            intPos = _scenarioEdge.Fields.FindField("SplitTK")
            _scenarioEdge.Value(intPos) = value
            _scenarioEdge.Store()
        End Set
    End Property
    Public Property Updated1() As String
        Get

            Dim intPos As Integer
            intPos = _scenarioEdge.Fields.FindField("Updated1")
            If IsDBNull(_scenarioEdge.Value(intPos)) Then
                Updated1 = "Null"
            Else
                Updated1 = _scenarioEdge.Value(intPos)
            End If


        End Get
        Set(ByVal value As String)
            Dim intPos As Integer
            intPos = _scenarioEdge.Fields.FindField("Updated1")
            _scenarioEdge.Value(intPos) = value
            _scenarioEdge.Store()
        End Set
    End Property
    Public Property prjRte() As String
        Get

            Dim intPos As Integer
            intPos = _scenarioEdge.Fields.FindField("prjRte")
            If IsDBNull(_scenarioEdge.Value(intPos)) Then
                prjRte = "Null"
            Else
                prjRte = _scenarioEdge.Value(intPos)
            End If


        End Get
        Set(ByVal value As String)
            Dim intPos As Integer
            intPos = _scenarioEdge.Fields.FindField("prjRte")
            _scenarioEdge.Value(intPos) = value
            _scenarioEdge.Store()
        End Set
    End Property
    Public Property shptype() As String
        Get

            Dim intPos As Integer
            intPos = _scenarioEdge.Fields.FindField("shptype")
            If IsDBNull(_scenarioEdge.Value(intPos)) Then
                shptype = "Null"
            Else
                shptype = _scenarioEdge.Value(intPos)
            End If


        End Get
        Set(ByVal value As String)
            Dim intPos As Integer
            intPos = _scenarioEdge.Fields.FindField("shptype")
            _scenarioEdge.Value(intPos) = value
            _scenarioEdge.Store()
        End Set
    End Property
    Public Property Dissolve() As String
        Get

            Dim intPos As Integer
            intPos = _scenarioEdge.Fields.FindField("Dissolve")
            If IsDBNull(_scenarioEdge.Value(intPos)) Then
                Dissolve = "Null"
            Else
                Dissolve = _scenarioEdge.Value(intPos)
            End If


        End Get
        Set(ByVal value As String)
            Dim intPos As Integer
            intPos = _scenarioEdge.Fields.FindField("Dissolve")
            _scenarioEdge.Value(intPos) = value
            _scenarioEdge.Store()
        End Set
    End Property
    Public Property Direction() As String
        Get

            Dim intPos As Integer
            intPos = _scenarioEdge.Fields.FindField("Direction")
            If IsDBNull(_scenarioEdge.Value(intPos)) Then
                Direction = "Null"
            Else
                Direction = _scenarioEdge.Value(intPos)
            End If


        End Get
        Set(ByVal value As String)
            Dim intPos As Integer
            intPos = _scenarioEdge.Fields.FindField("Direction")
            _scenarioEdge.Value(intPos) = value
            _scenarioEdge.Store()
        End Set
    End Property


End Class
