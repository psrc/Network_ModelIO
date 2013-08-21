Imports ESRI.ArcGIS.Geodatabase
Public Class clsProjectAttributes
    Private projectFeature As IFeature
    Public Sub New(ByVal ProjectLine As IFeature)
        projectFeature = ProjectLine



    End Sub
    Public ReadOnly Property PROJRTEID() As Long

        Get
            Dim intPos As Integer
            intPos = projectFeature.Fields.FindField("PROJRTEID")
            PROJRTEID = projectFeature.Value(intPos)
        End Get

    End Property
End Class
