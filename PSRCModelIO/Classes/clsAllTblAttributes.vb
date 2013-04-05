Imports ESRI.ArcGIS.Geodatabase
Public Class clsAllTblAttributes
    Private _AttributesRow As IRow
    Public Sub New(ByVal row As Irow)
        _AttributesRow = row
    End Sub
    Public ReadOnly Property OID() As Long
        Get
            'Dim intPos As Integer
            'intPos = modeAttributeRow.Fields.FindField("PSRCEDGEID")
            OID = _AttributesRow.Value(0)
        End Get
    End Property

    Public Property IJLANESGPAM() As Long
        Get
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("IJLANESGPAM")
            IJLANESGPAM = _AttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("IJLANESGPAM")
            _AttributesRow.Value(intPos) = value
            _AttributesRow.Store()
        End Set
    End Property






    Public Property IJLANESGPMD() As Long
        Get
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("IJLANESGPMD")
            IJLANESGPMD = _AttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("IJLANESGPMD")
            _AttributesRow.Value(intPos) = value
            _AttributesRow.Store()
        End Set
    End Property


    Public Property IJLANESGPPM() As Long
        Get
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("IJLANESGPPM")
            IJLANESGPPM = _AttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("IJLANESGPPM")
            _AttributesRow.Value(intPos) = value
            _AttributesRow.Store()
        End Set
    End Property


    Public Property IJLANESGPEV() As Long
        Get
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("IJLANESGPEV")
            IJLANESGPEV = _AttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("IJLANESGPEV")
            _AttributesRow.Value(intPos) = value
            _AttributesRow.Store()
        End Set
    End Property



    Public Property IJLANESGPNI() As Long
        Get
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("IJLANESGPNI")
            IJLANESGPNI = _AttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("IJLANESGPNI")
            _AttributesRow.Value(intPos) = value
            _AttributesRow.Store()
        End Set
    End Property



    Public Property JILANESGPAM() As Long
        Get
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("JILANESGPAM")
            JILANESGPAM = _AttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("JILANESGPAM")
            _AttributesRow.Value(intPos) = value
            _AttributesRow.Store()
        End Set
    End Property


    Public Property JILANESGPMD() As Long
        Get
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("JILANESGPMD")
            JILANESGPMD = _AttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("JILANESGPMD")
            _AttributesRow.Value(intPos) = value
            _AttributesRow.Store()
        End Set
    End Property


    Public Property JILANESGPPM() As Long
        Get
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("JILANESGPPM")
            JILANESGPPM = _AttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("JILANESGPPM")
            _AttributesRow.Value(intPos) = value
            _AttributesRow.Store()
        End Set
    End Property

    Public Property JILANESGPEV() As Long
        Get
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("JILANESGPEV")
            JILANESGPEV = _AttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intpos As Integer
            intpos = _AttributesRow.Fields.FindField("JILANESGPEV")
            _AttributesRow.Value(intpos) = value
            _AttributesRow.Store()
        End Set
    End Property


    Public Property JILANESGPNI() As Long
        Get
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("JILANESGPNI")
            JILANESGPNI = _AttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("JILANESGPNI")
            _AttributesRow.Value(intPos) = value
            _AttributesRow.Store()
        End Set
    End Property


    Public Property IJLANESGPADJUST() As Double
        Get
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("IJLANESGPADJUST")
            IJLANESGPADJUST = _AttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Double)
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("IJLANESGPADJUST")
            _AttributesRow.Value(intPos) = value
            _AttributesRow.Store()
        End Set
    End Property



    Public Property JILANESGPADJUST() As Double
        Get
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("JILANESGPADJUST")
            JILANESGPADJUST = _AttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Double)
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("JILANESGPADJUST")
            _AttributesRow.Value(intPos) = value
            _AttributesRow.Store()
        End Set
    End Property



    Public Property IJLANESHOVAM() As Long
        Get
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("IJLANESHOVAM")
            IJLANESHOVAM = _AttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("IJLANESHOVAM")
            _AttributesRow.Value(intPos) = value
            _AttributesRow.Store()
        End Set
    End Property


    Public Property IJLANESHOVMD() As Long
        Get
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("IJLANESHOVMD")
            IJLANESHOVMD = _AttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("IJLANESHOVMD")
            _AttributesRow.Value(intPos) = value
            _AttributesRow.Store()
        End Set
    End Property


    Public Property IJLANESHOVPM() As Long
        Get
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("IJLANESHOVPM")
            IJLANESHOVPM = _AttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("IJLANESHOVPM")
            _AttributesRow.Value(intPos) = value
            _AttributesRow.Store()
        End Set
    End Property




    Public Property IJLANESHOVEV() As Long
        Get
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("IJLANESHOVEV")
            IJLANESHOVEV = _AttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("IJLANESHOVEV")
            _AttributesRow.Value(intPos) = value
            _AttributesRow.Store()
        End Set
    End Property


    Public Property IJLANESHOVNI() As Long
        Get
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("IJLANESHOVNI")
            IJLANESHOVNI = _AttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("IJLANESHOVNI")
            _AttributesRow.Value(intPos) = value
            _AttributesRow.Store()
        End Set
    End Property

    Public Property JILANESHOVAM() As Long
        Get
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("JILANESHOVAM")
            JILANESHOVAM = _AttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("JILANESHOVAM")
            _AttributesRow.Value(intPos) = value
            _AttributesRow.Store()
        End Set
    End Property


    Public Property JILANESHOVMD() As Long
        Get
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("JILANESHOVMD")
            JILANESHOVMD = _AttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("JILANESHOVMD")
            _AttributesRow.Value(intPos) = value
            _AttributesRow.Store()
        End Set
    End Property


    Public Property JILANESHOVPM() As Long
        Get
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("JILANESHOVPM")
            JILANESHOVPM = _AttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("JILANESHOVPM")
            _AttributesRow.Value(intPos) = value
            _AttributesRow.Store()
        End Set
    End Property


    Public Property JILANESHOVEV() As Long
        Get
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("JILANESHOVEV")
            JILANESHOVEV = _AttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("JILANESHOVEV")
            _AttributesRow.Value(intPos) = value
            _AttributesRow.Store()
        End Set
    End Property


    Public Property JILANESHOVNI() As Long
        Get
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("JILANESHOVNI")
            JILANESHOVNI = _AttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("JILANESHOVNI")
            _AttributesRow.Value(intPos) = value
            _AttributesRow.Store()
        End Set
    End Property


    Public Property IJSPEEDLIMIT() As Long
        Get
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("IJSPEEDLIMIT")
            IJSPEEDLIMIT = _AttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("IJSPEEDLIMIT")
            _AttributesRow.Value(intPos) = value
            _AttributesRow.Store()
        End Set
    End Property


    Public Property JISPEEDLIMIT() As Long
        Get
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("JISPEEDLIMIT")
            JISPEEDLIMIT = _AttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("JISPEEDLIMIT")
            _AttributesRow.Value(intPos) = value
            _AttributesRow.Store()
        End Set
    End Property

    Public Property IJFFS() As Double
        Get
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("IJFFS")
            IJFFS = _AttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Double)
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("IJFFS")
            _AttributesRow.Value(intPos) = value
            _AttributesRow.Store()
        End Set
    End Property



    Public Property JIFFS() As Long
        Get
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("JIFFS")
            JIFFS = _AttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("JIFFS")
            _AttributesRow.Value(intPos) = value
            _AttributesRow.Store()
        End Set
    End Property


    Public Property IJVDFUNC() As Long
        Get
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("IJVDFUNC")
            IJVDFUNC = _AttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("IJVDFUNC")
            _AttributesRow.Value(intPos) = value
            _AttributesRow.Store()
        End Set
    End Property



    Public Property JIVDFUNC() As Long
        Get
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("JIVDFUNC")
            JIVDFUNC = _AttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("JIVDFUNC")
            _AttributesRow.Value(intPos) = value
            _AttributesRow.Store()
        End Set
    End Property


    Public Property IJLANECAPGP() As Long
        Get
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("IJLANECAPGP")
            IJLANECAPGP = _AttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("IJLANECAPGP")
            _AttributesRow.Value(intPos) = value
            _AttributesRow.Store()
        End Set
    End Property

    Public Property IJLANECAPHOV() As Long
        Get
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("IJLANECAPHOV")
            IJLANECAPHOV = _AttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("IJLANECAPHOV")
            _AttributesRow.Value(intPos) = value
            _AttributesRow.Store()
        End Set
    End Property


    Public Property JILANECAPGP() As Long
        Get
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("JILANECAPGP")
            JILANECAPGP = _AttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("JILANECAPGP")
            _AttributesRow.Value(intPos) = value
            _AttributesRow.Store()
        End Set
    End Property


    Public Property JILANECAPHOV() As Long
        Get
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("JILANECAPHOV")
            JILANECAPHOV = _AttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("JILANECAPHOV")
            _AttributesRow.Value(intPos) = value
            _AttributesRow.Store()
        End Set
    End Property


    Public Property IJSIDEWALKS() As Long
        Get
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("IJSIDEWALKS")
            IJSIDEWALKS = _AttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("IJSIDEWALKS")
            _AttributesRow.Value(intPos) = value
            _AttributesRow.Store()
        End Set
    End Property


    Public Property JISIDEWALKS() As Long
        Get
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("JISIDEWALKS")
            JISIDEWALKS = _AttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("JISIDEWALKS")
            _AttributesRow.Value(intPos) = value
            _AttributesRow.Store()
        End Set
    End Property


    Public Property IJBIKELANES() As Long
        Get
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("IJBIKELANES")
            IJBIKELANES = _AttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("IJBIKELANES")
            _AttributesRow.Value(intPos) = value
            _AttributesRow.Store()
        End Set
    End Property



    Public Property JIBIKELANES() As Long
        Get
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("JIBIKELANES")
            JIBIKELANES = _AttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("JIBIKELANES")
            _AttributesRow.Value(intPos) = value
            _AttributesRow.Store()
        End Set
    End Property


    Public Property IJLANESTR() As Long
        Get
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("IJLANESTR")
            IJLANESTR = _AttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("IJLANESTR")
            _AttributesRow.Value(intPos) = value
            _AttributesRow.Store()
        End Set
    End Property


    Public Property JILANESTR() As Long
        Get
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("JILANESTR")
            JILANESTR = _AttributesRow.Value(intPos)

        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("JILANESTR")
            _AttributesRow.Value(intPos) = value
            _AttributesRow.Store()
        End Set
    End Property


    Public Property IJLANESTK() As Long
        Get
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("IJLANESTK")
            IJLANESTK = _AttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("IJLANESTK")
            _AttributesRow.Value(intPos) = value
            _AttributesRow.Store()
        End Set
    End Property



    Public Property JILANESTK() As Long
        Get
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("JILANESTK")
            JILANESTK = _AttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("JILANESTK")
            _AttributesRow.Value(intPos) = value
            _AttributesRow.Store()
        End Set
    End Property


    Public Property PSRC_E2ID() As Long
        Get
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("PSRC_E2ID")
            PSRC_E2ID = _AttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("PSRC_E2ID")
            _AttributesRow.Value(intPos) = value
            _AttributesRow.Store()
        End Set
    End Property


    Public Property OLD_EDGEID() As Long
        Get
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("OLD_EDGEID")
            OLD_EDGEID = _AttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("OLD_EDGEID")
            _AttributesRow.Value(intPos) = value
            _AttributesRow.Store()
        End Set
    End Property

    Public Property NEW1() As Long
        Get
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("NEW")
            NEW1 = _AttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("NEW")
            _AttributesRow.Value(intPos) = value
            _AttributesRow.Store()
        End Set
    End Property
    Public ReadOnly Property projRteID() As Long
        Get
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("PSRCEDGEID")
            projRteID = _AttributesRow.Value(intPos)
        End Get
    End Property
    Public ReadOnly Property projDBS() As Long
        Get
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("projDBS")
            projDBS = _AttributesRow.Value(intPos)
        End Get
    End Property
    Public ReadOnly Property projID() As String
        Get
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("projID")
            projID = _AttributesRow.Value(intPos)
        End Get
    End Property
    Public ReadOnly Property version() As Integer
        Get
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("version")
            version = _AttributesRow.Value(intPos)
        End Get
    End Property
    Public ReadOnly Property InServiceDate() As Short
        Get
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("InServiceDate")
            InServiceDate = _AttributesRow.Value(intPos)
        End Get
    End Property
    Public ReadOnly Property OutServiceDate() As Short
        Get
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("OutServiceDate")
            OutServiceDate = _AttributesRow.Value(intPos)
        End Get
    End Property
    Public ReadOnly Property CompletionDate() As Short
        Get
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("CompletionDate")
            CompletionDate = _AttributesRow.Value(intPos)
        End Get
    End Property

    Public ReadOnly Property Modes() As String
        Get
            Dim intPos As Integer
            intPos = _AttributesRow.Fields.FindField("Modes")
            Modes = _AttributesRow.Value(intPos)
        End Get
    End Property
End Class


