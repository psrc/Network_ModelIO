Imports ESRI.ArcGIS.Geodatabase
Public Class clsTransRefEdgeAtts
    Inherits clsBaseEdgeFeatAttributes
    Private refEdge As IFeature
    Public Sub New(ByVal transRefEdgeFeature As IFeature)
        MyBase.New(transRefEdgeFeature)
        refEdge = transRefEdgeFeature
    End Sub

End Class
