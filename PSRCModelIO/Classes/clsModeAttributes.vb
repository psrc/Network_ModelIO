Imports ESRI.ArcGIS.Geodatabase
Public Class clsModeAttributes
    Inherits clsBaseTblAttributes


    Private modeAttributeRow As IRow
    Public Sub New(ByVal row As IRow)
        MyBase.New(row)
        modeAttributeRow = row
    End Sub


    Public Property PSRCEDGEID() As Long
        Get
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("PSRCEDGEID")
            PSRCEDGEID = modeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("PSRCEDGEID")
            modeAttributeRow.Value(intPos) = value
            modeAttributeRow.Store()
        End Set
    End Property





    Shadows ReadOnly Property ToString() As String
        Get
            Return "ModeAttributes."
        End Get

    End Property
    Public Property AttributeByString(ByVal Name As String) As Long


        Get
            Name = UCase(Name)
            Select Case Name
                Case "IJBIKELANES"
                    Return Me.IJBIKELANES
                Case "IJFFS"
                    Return Me.IJFFS
                Case "IJLANECAPGP"
                    Return Me.IJLANECAPGP
                Case "IJLANECAPHOV"
                    Return Me.IJLANECAPHOV
                Case "IJLANESGPADJUST"
                    Return Me.IJLANESGPADJUST
                Case "IJLANESGPAM"
                    Return Me.IJLANESGPAM
                Case "IJLANESGPEV"
                    Return Me.IJLANESGPEV
                Case "IJLANESGPMD"
                    Return Me.IJLANESGPMD
                Case "IJLANESGPNI"
                    Return Me.IJLANESGPNI
                Case "IJLANESGPPM"
                    Return Me.IJLANESGPPM
                Case "IJLANESHOVAM"
                    Return Me.IJLANESHOVAM
                Case "IJLANESHOVEV"
                    Return Me.IJLANESHOVEV
                Case "IJLANESHOVMD"
                    Return Me.IJLANESHOVMD
                Case "IJLANESHOVNI"
                    Return Me.IJLANESHOVNI
                Case "IJLANESHOVNI"
                    Return Me.IJLANESHOVNI
                Case "IJLANESHOVPM"
                    Return Me.IJLANESHOVPM
                Case "IJLANESTK"
                    Return Me.IJLANESTK
                Case "IJLANESTR"
                    Return Me.IJLANESTR
                Case "IJSIDEWALKS"
                    Return Me.IJSIDEWALKS
                Case "IJSPEEDLIMIT"
                    Return Me.IJSPEEDLIMIT
                Case "IJVDFUNC"
                    Return Me.IJVDFUNC
                Case "JIBIKELANES"
                    Return Me.JIBIKELANES
                Case "JIFFS"
                    Return Me.JIFFS
                Case "JILANECAPGP"
                    Return Me.JILANECAPGP
                Case "JILANECAPHOV"
                    Return Me.JILANECAPHOV
                Case "JILANESGPADJUST"
                    Return Me.JILANESGPADJUST
                Case "JILANESGPAM"
                    Return Me.JILANESGPAM
                Case "JILANESGPEV"
                    Return Me.JILANESGPEV
                Case "JILANESGPMD"
                    Return Me.JILANESGPMD
                Case "JILANESGPNI"
                    Return Me.JILANESGPNI
                Case "JILANESGPPM"
                    Return Me.JILANESGPPM
                Case "JILANESHOVAM"
                    Return Me.JILANESHOVAM
                Case "JILANESHOVEV"
                    Return Me.JILANESHOVEV
                Case "JILANESHOVMD"
                    Return Me.JILANESHOVMD
                Case "JILANESHOVNI"
                    Return Me.JILANESHOVNI
                Case "JILANESHOVPM"
                    Return Me.JILANESHOVPM
                Case "JILANESTK"
                    Return Me.JILANESTK
                Case "JILANESTR"
                    Return Me.JILANESTR
                Case "JILANESTR"
                    Return Me.JILANESTR
                Case "JISIDEWALKS"
                    Return Me.JISIDEWALKS
                Case "JISPEEDLIMIT"
                    Return Me.JISPEEDLIMIT
                Case "JIVDFUNC"
                    Return Me.JIVDFUNC


            End Select








        End Get
        Set(ByVal value As Long)

            Select Case Name
                Case "IJBIKELANES"
                    Me.IJBIKELANES = value
                Case "IJFFS"
                    Me.IJFFS = value
                Case "IJLANECAPGP"
                    Me.IJLANECAPGP = value
                Case "IJLANECAPHOV"
                    Me.IJLANECAPHOV = value
                Case "IJLANESGPADJUST"
                    Me.IJLANESGPADJUST = value
                Case "IJLANESGPAM"
                    Me.IJLANESGPAM = value
                Case "IJLANESGPEV"
                    Me.IJLANESGPEV = value
                Case "IJLANESGPMD"
                    Me.IJLANESGPMD = value
                Case "IJLANESGPNI"
                    Me.IJLANESGPNI = value
                Case "IJLANESGPPM"
                    Me.IJLANESGPPM = value
                Case "IJLANESHOVAM"
                    Me.IJLANESHOVAM = value
                Case "IJLANESHOVEV"
                    Me.IJLANESHOVEV = value
                Case "IJLANESHOVMD"
                    Me.IJLANESHOVMD = value
                Case "IJLANESHOVNI"
                    Me.IJLANESHOVNI = value
                Case "IJLANESHOVNI"
                    Me.IJLANESHOVNI = value
                Case "IJLANESHOVPM"
                    Me.IJLANESHOVPM = value
                Case "IJLANESTK"
                    Me.IJLANESTK = value
                Case "IJLANESTR"
                    Me.IJLANESTR = value
                Case "IJSIDEWALKS"
                    Me.IJSIDEWALKS = value
                Case "IJSPEEDLIMIT"
                    Me.IJSPEEDLIMIT = value
                Case "IJVDFUNC"
                    Me.IJVDFUNC = value
                Case "JIBIKELANES"
                    Me.JIBIKELANES = value
                Case "JIFFS"
                    Me.JIFFS = value
                Case "JILANECAPGP"
                    Me.JILANECAPGP = value
                Case "JILANECAPHOV"
                    Me.JILANECAPHOV = value
                Case "JILANESGPADJUST"
                    Me.JILANESGPADJUST = value
                Case "JILANESGPAM"
                    Me.JILANESGPAM = value
                Case "JILANESGPEV"
                    Me.JILANESGPEV = value
                Case "JILANESGPMD"
                    Me.JILANESGPMD = value
                Case "JILANESGPNI"
                    Me.JILANESGPNI = value
                Case "JILANESGPPM"
                    Me.JILANESGPPM = value
                Case "JILANESHOVAM"
                    Me.JILANESHOVAM = value
                Case "JILANESHOVEV"
                    Me.JILANESHOVEV = value
                Case "JILANESHOVMD"
                    Me.JILANESHOVMD = value
                Case "JILANESHOVNI"
                    Me.JILANESHOVNI = value
                Case "JILANESHOVPM"
                    Me.JILANESHOVPM = value
                Case "JILANESTK"
                    Me.JILANESTK = value
                Case "JILANESTR"
                    Me.JILANESTR = value
                Case "JILANESTR"
                    Me.JILANESTR = value
                Case "JISIDEWALKS"
                    Me.JISIDEWALKS = value
                Case "JISPEEDLIMIT"
                    Me.JISPEEDLIMIT = value
                Case "JIVDFUNC"
                    Me.JIVDFUNC = value
            End Select

        End Set
    End Property

    Public ReadOnly Property FindField(ByVal FieldName As String) As Long
        Get
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField(FieldName)
            FindField = modeAttributeRow.Value(intPos)
        End Get
    End Property





End Class

