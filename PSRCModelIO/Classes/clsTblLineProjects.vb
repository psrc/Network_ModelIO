Imports ESRI.ArcGIS.Geodatabase

Public Class clsTblLineProjects
    Inherits clsBaseTblAttributes
    Private tblLineProjectRow As IRow
    Public Sub New(ByVal row As IRow)
        MyBase.New(row)
        tblLineProjectRow = row
    End Sub
    Public ReadOnly Property projRteID() As Long
        Get
            Dim intPos As Integer
            intPos = tblLineProjectRow.Fields.FindField("PSRCEDGEID")
            projRteID = tblLineProjectRow.Value(intPos)
        End Get
    End Property
    Public ReadOnly Property projDBS() As Long
        Get
            Dim intPos As Integer
            intPos = tblLineProjectRow.Fields.FindField("projDBS")
            projDBS = tblLineProjectRow.Value(intPos)
        End Get
    End Property
    Public ReadOnly Property projID() As String
        Get
            Dim intPos As Integer
            intPos = tblLineProjectRow.Fields.FindField("projID")
            projID = tblLineProjectRow.Value(intPos)
        End Get
    End Property
    Public ReadOnly Property version() As Integer
        Get
            Dim intPos As Integer
            intPos = tblLineProjectRow.Fields.FindField("version")
            version = tblLineProjectRow.Value(intPos)
        End Get
    End Property
    Public ReadOnly Property InServiceDate() As Short
        Get
            Dim intPos As Integer
            intPos = tblLineProjectRow.Fields.FindField("InServiceDate")
            InServiceDate = tblLineProjectRow.Value(intPos)
        End Get
    End Property
    Public ReadOnly Property OutServiceDate() As Short
        Get
            Dim intPos As Integer
            intPos = tblLineProjectRow.Fields.FindField("OutServiceDate")
            OutServiceDate = tblLineProjectRow.Value(intPos)
        End Get
    End Property
    Public ReadOnly Property CompletionDate() As Short
        Get
            Dim intPos As Integer
            intPos = tblLineProjectRow.Fields.FindField("CompletionDate")
            CompletionDate = tblLineProjectRow.Value(intPos)
        End Get
    End Property

    Public ReadOnly Property Modes() As String
        Get
            Dim intPos As Integer
            intPos = tblLineProjectRow.Fields.FindField("Modes")
            Modes = tblLineProjectRow.Value(intPos)
        End Get
    End Property
End Class
