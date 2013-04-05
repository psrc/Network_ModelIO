Imports ESRI.ArcGIS.Geodatabase
Public Class clsBaseTblAttributes
    Private _baseAttributesRow As IRow
    Public Sub New(ByVal row As Irow)
        _baseAttributesRow = row
    End Sub
    Public ReadOnly Property OID() As Long
        Get
            'Dim intPos As Integer
            'intPos = modeAttributeRow.Fields.FindField("PSRCEDGEID")
            OID = _baseAttributesRow.Value(0)
        End Get
    End Property

    Public Property IJLANESGPAM() As Long
        Get
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("IJLANESGPAM")
            IJLANESGPAM = _baseAttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("IJLANESGPAM")
            _baseAttributesRow.Value(intPos) = value
            _baseAttributesRow.Store()
        End Set
    End Property






    Public Property IJLANESGPMD() As Long
        Get
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("IJLANESGPMD")
            IJLANESGPMD = _baseAttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("IJLANESGPMD")
            _baseAttributesRow.Value(intPos) = value
            _baseAttributesRow.Store()
        End Set
    End Property


    Public Property IJLANESGPPM() As Long
        Get
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("IJLANESGPPM")
            IJLANESGPPM = _baseAttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("IJLANESGPPM")
            _baseAttributesRow.Value(intPos) = value
            _baseAttributesRow.Store()
        End Set
    End Property


    Public Property IJLANESGPEV() As Long
        Get
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("IJLANESGPEV")
            IJLANESGPEV = _baseAttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("IJLANESGPEV")
            _baseAttributesRow.Value(intPos) = value
            _baseAttributesRow.Store()
        End Set
    End Property



    Public Property IJLANESGPNI() As Long
        Get
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("IJLANESGPNI")
            IJLANESGPNI = _baseAttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("IJLANESGPNI")
            _baseAttributesRow.Value(intPos) = value
            _baseAttributesRow.Store()
        End Set
    End Property



    Public Property JILANESGPAM() As Long
        Get
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("JILANESGPAM")
            JILANESGPAM = _baseAttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("JILANESGPAM")
            _baseAttributesRow.Value(intPos) = value
            _baseAttributesRow.Store()
        End Set
    End Property


    Public Property JILANESGPMD() As Long
        Get
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("JILANESGPMD")
            JILANESGPMD = _baseAttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("JILANESGPMD")
            _baseAttributesRow.Value(intPos) = value
            _baseAttributesRow.Store()
        End Set
    End Property


    Public Property JILANESGPPM() As Long
        Get
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("JILANESGPPM")
            JILANESGPPM = _baseAttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("JILANESGPPM")
            _baseAttributesRow.Value(intPos) = value
            _baseAttributesRow.Store()
        End Set
    End Property

    Public Property JILANESGPEV() As Long
        Get
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("JILANESGPEV")
            JILANESGPEV = _baseAttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intpos As Integer
            intpos = _baseAttributesRow.Fields.FindField("JILANESGPEV")
            _baseAttributesRow.Value(intpos) = value
            _baseAttributesRow.Store()
        End Set
    End Property


    Public Property JILANESGPNI() As Long
        Get
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("JILANESGPNI")
            JILANESGPNI = _baseAttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("JILANESGPNI")
            _baseAttributesRow.Value(intPos) = value
            _baseAttributesRow.Store()
        End Set
    End Property


    Public Property IJLANESGPADJUST() As Double
        Get
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("IJLANESGPADJUST")
            IJLANESGPADJUST = _baseAttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Double)
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("IJLANESGPADJUST")
            _baseAttributesRow.Value(intPos) = value
            _baseAttributesRow.Store()
        End Set
    End Property



    Public Property JILANESGPADJUST() As Double
        Get
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("JILANESGPADJUST")
            JILANESGPADJUST = _baseAttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Double)
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("JILANESGPADJUST")
            _baseAttributesRow.Value(intPos) = value
            _baseAttributesRow.Store()
        End Set
    End Property



    Public Property IJLANESHOVAM() As Long
        Get
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("IJLANESHOVAM")
            IJLANESHOVAM = _baseAttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("IJLANESHOVAM")
            _baseAttributesRow.Value(intPos) = value
            _baseAttributesRow.Store()
        End Set
    End Property


    Public Property IJLANESHOVMD() As Long
        Get
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("IJLANESHOVMD")
            IJLANESHOVMD = _baseAttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("IJLANESHOVMD")
            _baseAttributesRow.Value(intPos) = value
            _baseAttributesRow.Store()
        End Set
    End Property


    Public Property IJLANESHOVPM() As Long
        Get
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("IJLANESHOVPM")
            IJLANESHOVPM = _baseAttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("IJLANESHOVPM")
            _baseAttributesRow.Value(intPos) = value
            _baseAttributesRow.Store()
        End Set
    End Property




    Public Property IJLANESHOVEV() As Long
        Get
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("IJLANESHOVEV")
            IJLANESHOVEV = _baseAttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("IJLANESHOVEV")
            _baseAttributesRow.Value(intPos) = value
            _baseAttributesRow.Store()
        End Set
    End Property


    Public Property IJLANESHOVNI() As Long
        Get
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("IJLANESHOVNI")
            IJLANESHOVNI = _baseAttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("IJLANESHOVNI")
            _baseAttributesRow.Value(intPos) = value
            _baseAttributesRow.Store()
        End Set
    End Property

    Public Property JILANESHOVAM() As Long
        Get
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("JILANESHOVAM")
            JILANESHOVAM = _baseAttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("JILANESHOVAM")
            _baseAttributesRow.Value(intPos) = value
            _baseAttributesRow.Store()
        End Set
    End Property


    Public Property JILANESHOVMD() As Long
        Get
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("JILANESHOVMD")
            JILANESHOVMD = _baseAttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("JILANESHOVMD")
            _baseAttributesRow.Value(intPos) = value
            _baseAttributesRow.Store()
        End Set
    End Property


    Public Property JILANESHOVPM() As Long
        Get
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("JILANESHOVPM")
            JILANESHOVPM = _baseAttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("JILANESHOVPM")
            _baseAttributesRow.Value(intPos) = value
            _baseAttributesRow.Store()
        End Set
    End Property


    Public Property JILANESHOVEV() As Long
        Get
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("JILANESHOVEV")
            JILANESHOVEV = _baseAttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("JILANESHOVEV")
            _baseAttributesRow.Value(intPos) = value
            _baseAttributesRow.Store()
        End Set
    End Property


    Public Property JILANESHOVNI() As Long
        Get
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("JILANESHOVNI")
            JILANESHOVNI = _baseAttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("JILANESHOVNI")
            _baseAttributesRow.Value(intPos) = value
            _baseAttributesRow.Store()
        End Set
    End Property


    Public Property IJSPEEDLIMIT() As Long
        Get
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("IJSPEEDLIMIT")
            IJSPEEDLIMIT = _baseAttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("IJSPEEDLIMIT")
            _baseAttributesRow.Value(intPos) = value
            _baseAttributesRow.Store()
        End Set
    End Property


    Public Property JISPEEDLIMIT() As Long
        Get
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("JISPEEDLIMIT")
            JISPEEDLIMIT = _baseAttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("JISPEEDLIMIT")
            _baseAttributesRow.Value(intPos) = value
            _baseAttributesRow.Store()
        End Set
    End Property

    Public Property IJFFS() As Double
        Get
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("IJFFS")
            IJFFS = _baseAttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Double)
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("IJFFS")
            _baseAttributesRow.Value(intPos) = value
            _baseAttributesRow.Store()
        End Set
    End Property



    Public Property JIFFS() As Long
        Get
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("JIFFS")
            JIFFS = _baseAttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("JIFFS")
            _baseAttributesRow.Value(intPos) = value
            _baseAttributesRow.Store()
        End Set
    End Property


    Public Property IJVDFUNC() As Long
        Get
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("IJVDFUNC")
            IJVDFUNC = _baseAttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("IJVDFUNC")
            _baseAttributesRow.Value(intPos) = value
            _baseAttributesRow.Store()
        End Set
    End Property



    Public Property JIVDFUNC() As Long
        Get
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("JIVDFUNC")
            JIVDFUNC = _baseAttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("JIVDFUNC")
            _baseAttributesRow.Value(intPos) = value
            _baseAttributesRow.Store()
        End Set
    End Property


    Public Property IJLANECAPGP() As Long
        Get
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("IJLANECAPGP")
            IJLANECAPGP = _baseAttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("IJLANECAPGP")
            _baseAttributesRow.Value(intPos) = value
            _baseAttributesRow.Store()
        End Set
    End Property

    Public Property IJLANECAPHOV() As Long
        Get
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("IJLANECAPHOV")
            IJLANECAPHOV = _baseAttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("IJLANECAPHOV")
            _baseAttributesRow.Value(intPos) = value
            _baseAttributesRow.Store()
        End Set
    End Property


    Public Property JILANECAPGP() As Long
        Get
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("JILANECAPGP")
            JILANECAPGP = _baseAttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("JILANECAPGP")
            _baseAttributesRow.Value(intPos) = value
            _baseAttributesRow.Store()
        End Set
    End Property


    Public Property JILANECAPHOV() As Long
        Get
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("JILANECAPHOV")
            JILANECAPHOV = _baseAttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("JILANECAPHOV")
            _baseAttributesRow.Value(intPos) = value
            _baseAttributesRow.Store()
        End Set
    End Property


    Public Property IJSIDEWALKS() As Long
        Get
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("IJSIDEWALKS")
            IJSIDEWALKS = _baseAttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("IJSIDEWALKS")
            _baseAttributesRow.Value(intPos) = value
            _baseAttributesRow.Store()
        End Set
    End Property


    Public Property JISIDEWALKS() As Long
        Get
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("JISIDEWALKS")
            JISIDEWALKS = _baseAttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("JISIDEWALKS")
            _baseAttributesRow.Value(intPos) = value
            _baseAttributesRow.Store()
        End Set
    End Property


    Public Property IJBIKELANES() As Long
        Get
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("IJBIKELANES")
            IJBIKELANES = _baseAttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("IJBIKELANES")
            _baseAttributesRow.Value(intPos) = value
            _baseAttributesRow.Store()
        End Set
    End Property



    Public Property JIBIKELANES() As Long
        Get
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("JIBIKELANES")
            JIBIKELANES = _baseAttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("JIBIKELANES")
            _baseAttributesRow.Value(intPos) = value
            _baseAttributesRow.Store()
        End Set
    End Property


    Public Property IJLANESTR() As Long
        Get
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("IJLANESTR")
            IJLANESTR = _baseAttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("IJLANESTR")
            _baseAttributesRow.Value(intPos) = value
            _baseAttributesRow.Store()
        End Set
    End Property


    Public Property JILANESTR() As Long
        Get
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("JILANESTR")
            JILANESTR = _baseAttributesRow.Value(intPos)

        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("JILANESTR")
            _baseAttributesRow.Value(intPos) = value
            _baseAttributesRow.Store()
        End Set
    End Property


    Public Property IJLANESTK() As Long
        Get
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("IJLANESTK")
            IJLANESTK = _baseAttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("IJLANESTK")
            _baseAttributesRow.Value(intPos) = value
            _baseAttributesRow.Store()
        End Set
    End Property



    Public Property JILANESTK() As Long
        Get
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("JILANESTK")
            JILANESTK = _baseAttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("JILANESTK")
            _baseAttributesRow.Value(intPos) = value
            _baseAttributesRow.Store()
        End Set
    End Property


    Public Property PSRC_E2ID() As Long
        Get
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("PSRC_E2ID")
            PSRC_E2ID = _baseAttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("PSRC_E2ID")
            _baseAttributesRow.Value(intPos) = value
            _baseAttributesRow.Store()
        End Set
    End Property


    Public Property OLD_EDGEID() As Long
        Get
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("OLD_EDGEID")
            OLD_EDGEID = _baseAttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("OLD_EDGEID")
            _baseAttributesRow.Value(intPos) = value
            _baseAttributesRow.Store()
        End Set
    End Property

    Public Property NEW1() As Long
        Get
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("NEW")
            NEW1 = _baseAttributesRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = _baseAttributesRow.Fields.FindField("NEW")
            _baseAttributesRow.Value(intPos) = value
            _baseAttributesRow.Store()
        End Set
    End Property
End Class
