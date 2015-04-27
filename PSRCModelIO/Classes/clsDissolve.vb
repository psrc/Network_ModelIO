
Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI


Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.SystemUI
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Carto
Imports System.Windows.Forms
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.DataSourcesGDB
'Imports ESRI.ArcGIS.esriSystem


Imports ESRI.ArcGIS.GeoDatabaseUI
Imports Scripting

Imports ESRI.ArcGIS.DataSourcesFile
Public Class clsDissolve
    Private _app As IApplication
    Private _doc As IMxDocument
    Private _activeView As IActiveView
    Private _map As IMap

    Public Sub New(ByVal app As Iapplication)
        _app = app
        InitiateMainVariables()
    End Sub
    Public Sub DissolveEdges()
        Dissolve_Keep_TTP()
        'DissolveByFT()


    End Sub
    Private Sub InitiateMainVariables()
        _doc = _app.Document
        _activeView = _doc.ActiveView
        _map = _activeView.FocusMap

    End Sub
    Private Sub Dissolve_Keep_TTP()
        '[022506] hyu: revised function to replace [dissolveTransEdges2_0]

        'Retains junctions and edges used by transit, tolls and turns
        'performs "pseudonode" thinning of the intermediate edge shp file
        'modified to handle toll, turn coincidence with thinning by Melinda 5-25-05
        'retrofitted (minimally) by Jeff with our logging stuff
        'added transit thinning by Melinda 6-3-05 pan

        'on error GoTo ErrChk

        WriteLogLine("")
        WriteLogLine("======================================")
        WriteLogLine("dissolveTransEdges2 started " & Now())
        WriteLogLine("======================================")
        WriteLogLine("")
        If (strLayerPrefix = "SDE") Then
            Dim pwf As IWorkspaceFactory
            pwf = New SdeWorkspaceFactory
        End If

        Dim pCounter As Long
        Dim pFLayerE As IFeatureLayer, pFLayerJ As IFeatureLayer
        pFLayerE = New FeatureLayer
        pFLayerE.FeatureClass = m_edgeShp 'Intermediate Edge
        WriteLogLine(m_edgeShp.AliasName + CStr(m_edgeShp.FeatureCount(Nothing)))
        pFLayerJ = New FeatureLayer
        pFLayerJ.FeatureClass = m_junctShp 'Intermediate Junction

        Dim pTblMode As ITable, pTolls As ITable
        Dim pTurn As IFeatureLayer
        Dim pVehicleCounts As IFeatureLayer
        Dim pTblPrjEdgeAtt As ITable, pEvtLine As ITable

        pTblPrjEdgeAtt = get_TableClass(m_layers(24))
        Try


            pTblMode = get_TableClass(m_layers(2)) 'modeAttributes
            pTolls = get_TableClass(m_layers(9)) 'modeTolls
            pTurn = get_FeatureLayer2(m_layers(12)) 'TurnMovements
            pEvtLine = get_TableClass(m_layers(4)) 'use if future event project
            pVehicleCounts = get_FeatureLayer2(m_layers(25))
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        '[041607]hyu: now dealing with transit points feature layer instead of segment table
        '  Dim tblTSeg As ITable
        '  Set tblTSeg = get_TableClass(m_layers(14)) 'tbltransitsegments
        Dim pTrPoints As IFeatureLayer
        pTrPoints = get_FeatureLayer2(m_layers(23))

        Dim pFC As IFeatureCursor, pFC2 As IFeatureCursor
        Dim pFCj As IFeatureCursor
        Dim pQF As IQueryFilter, pQFt As IQueryFilter, pQF2 As IQueryFilter
        Dim pQFj As IQueryFilter

        Dim pGeometry As IGeometry, pNewGeometry As IGeometry
        Dim pGeometryBag As IGeometryCollection
        pGeometryBag = New GeometryBag
        Dim pTopo As ITopologicalOperator
        Dim pPolyline As IPolyline

        Dim pFeat2 As IFeature
        Dim pNextFeat As IFeature, pNewFeat As IFeature
        Dim pJunctFeat As IFeature
        Dim pRow As IRow, pNextRow As IRow
        Dim pTC As ICursor
        Dim pFlds As IFields

        Dim psrcId As String
        Dim i As Long
        Dim fcountI, fcountIJ, mergeCountIJ, mergeCountJI As Long, noMergeCount As Long, attMismatchCount As Long

        Dim match As Boolean
        match = True
        Dim edgesdelete As Boolean
        edgesdelete = False

        Dim newI As Long, newJ As Long

        '  Dim pWorkspaceEdit As IWorkspaceEdit, pWorkspaceEdit2 As IWorkspaceEdit
        '  Set pWorkspaceEdit = g_FWS
        '  pWorkspaceEdit.StartEditing False
        '  pWorkspaceEdit.StartEditOperation

        '  'this one just loops thru junct layer- should be faster
        '  Set pFCj = m_junctShp.Search(Nothing, False)
        '  Set pJunctFeat = pFCj.NextFeature

        Dim lSFld As Integer, dnode As Long, idIndex As Integer
        Dim lIndex As Integer, upIndex As Integer
        Dim idxN As Integer, idxJ As Integer
        Dim idxMode As Integer, idxLT As Integer, idxFT As Integer
        Dim idxJunctID As Integer, idxEdgeID As Integer
        Dim ifld As Integer
        idxJunctID = m_junctShp.FindField("PSRCJunctID")
        '
        '[122305]pan--revised field names for current SDE layer match
        '
        With m_edgeShp
            idxEdgeID = .FindField(g_PSRCEdgeID)
            idIndex = .FindField("PSRC_E2ID")
            upIndex = .FindField("Updated1")
            idxN = .FindField(g_INode)
            idxJ = .FindField(g_JNode)
            idxMode = .FindField(g_Modes)
            idxLT = .FindField(g_LinkType)
            idxFT = .FindField(g_FacType)
        End With

        Dim m As Integer
        Dim same As String, opposite As String
        Dim dir2 As String, dirNext As String
        Dim index2 As Integer, indexNext As Integer
        Dim dir As Boolean
        Dim lType(4) As String
        lType(0) = "GP"
        lType(1) = "HOV"
        lType(2) = "TR"
        lType(3) = "TK"
        Dim timePd(5) As String
        timePd(0) = "AM"
        timePd(1) = "MD"
        timePd(2) = "PM"
        timePd(3) = "EV"
        timePd(4) = "NI"

        mergeCountIJ = 0
        mergeCountJI = 0
        noMergeCount = 0
        attMismatchCount = 0
        pCounter = 1

        '[022406] hyu: get those nodes/edges that can/cannot be thinned into dictionaries.
        Dim dctNoThinNodes As Dictionary(Of Object, Object), dctNoThinEdges As Dictionary(Of Object, Object), dctThinEdges As Dictionary(Of Object, Object)
        'dctNoThinNodes consists of nodes cannot be thinned
        'dctThinEdges consists of edges that are possible be thinned.

        dctNoThinNodes = New Dictionary(Of Object, Object)
        dctThinEdges = New Dictionary(Of Object, Object)
        dctNoThinEdges = New Dictionary(Of Object, Object)
        WriteLogLine(Now() & " start collecting nodes and edges ")

        '  [050306] hyu: transitSegment
        '  getNodes tblTSeg, g_PSRCEdgeID, dctNoThinEdges

        '[041607]hyu: dealing w/ transit points instead of segments table
        'getNodes tblTSeg, g_INode, dctNoThinNodes, "SUBSTRING(CAST(LineID AS CHARACTER), 2, 2) = '" & Right(CStr(inserviceyear), 2) & "' AND SegOrder=1"
        'getNodes tblTSeg, g_JNode, dctNoThinNodes, "SUBSTRING(CAST(LineID AS CHARACTER), 2, 2) = '" & Right(CStr(inserviceyear), 2) & "' AND change>0"
        '[051407] jaf: use SQL with SDE but not with Access
        If strLayerPrefix = "SDE" Then
            getNodes(pTrPoints, g_PSRCJctID, dctNoThinNodes, "SUBSTRING(CAST(LineID AS CHARACTER), 2, 2) = '" & Microsoft.VisualBasic.Right(CStr(inserviceyear), 2) & "' or SUBSTRING(CAST(LineID AS CHARACTER), 2, 2) = '" & Microsoft.VisualBasic.Right(CStr(inserviceyear + 1), 2) & "'")
            'getNodes(pTrPoints, g_PSRCJctID, dctNoThinNodes, "SUBSTRING(CAST(LineID AS CHARACTER), 2, 2) = '" & Microsoft.VisualBasic.Right(CStr(inserviceyear), 2) & "' or SUBSTRING(CAST(LineID AS CHARACTER), 2, 2) = '" & Microsoft.VisualBasic.Right(CStr(inserviceyear + 1), 2) & "' or SUBSTRING(CAST(LineID AS CHARACTER), 2, 2) = '" & Microsoft.VisualBasic.Right(CStr(inserviceyear + 1), 2) & "'")
            'getNodes(pTrPoints, g_PSRCJctID, dctNoThinNodes, "SUBSTRING(CAST(LineID AS CHARACTER), 2, 2) = '" & Microsoft.VisualBasic.Right(CStr(inserviceyear), 2) & "' or SUBSTRING(CAST(LineID AS CHARACTER), 2, 2) = '" & Microsoft.VisualBasic.Right(CStr(inserviceyear + 1), 2) & "'" & "' or SUBSTRING(CAST(LineID AS CHARACTER), 2, 2) = '" & Microsoft.VisualBasic.Right(CStr(inserviceyear + 3), 2) & "'" & "' or SUBSTRING(CAST(LineID AS CHARACTER), 2, 2) = '" & Microsoft.VisualBasic.Right(CStr(inserviceyear + 3), 2) & "'")
            'getPRNodes pFLayerJ.FeatureClass, "Scen_Node", "P_RStalls", dctNoThinNodes
        Else
            getNodes(pTrPoints, g_PSRCJctID, dctNoThinNodes, "MID(Cstr(LineID), 2, 2) = '" & Microsoft.VisualBasic.Right(CStr(inserviceyear), 2) & "'")
        End If

        WriteLogLine(Now() & " finish collecting junctions being transit points: " & dctNoThinNodes.Count)

        getNodes(pTblMode, g_PSRCEdgeID, dctThinEdges, , True)
        WriteLogLine(Now() & " finish collecting edges can be thinned: " & dctThinEdges.Count)

        getNodes(pTolls, g_PSRCEdgeID, dctNoThinEdges)
        WriteLogLine(Now() & " finish collecting edges can NOT be thinned: " & dctNoThinEdges.Count)

        '[060210] SEC- Need to preserve Turn Edges & "Retain" Junctions
        getNodes(pFLayerJ.FeatureClass, "PSRCJUNCTID", dctNoThinNodes, "JunctionType = 10")

        getNodes(pTurn, "FrEdgeID", dctNoThinEdges)

        getNodes(pTurn, "ToEdgeID", dctNoThinEdges)


        getNodes(pTurn, g_PSRCJctID, dctNoThinNodes)
        'Vehicle Counts:
        getNodes(pVehicleCounts, "PSRCEDGEID", dctNoThinEdges)

        WriteLogLine(Now() & " finish collecting junction can NOT be thinned (including transit): " & dctNoThinNodes.Count)

        '[082409] SEC: Added Park and Rides
        getNodes(m_edgeShp, g_PSRCEdgeID, dctNoThinEdges, g_FacType & "=16 or " & g_FacType & "=17") '[042706] hyu: exclude centroid connector + park and rides from dissolving
        WriteLogLine(Now() & " finish collecting edges can NOT be thinned : " & dctNoThinEdges.Count)

        '[022406] hyu: stats how many edges joining to a junction
        Dim dctJoinEdges As Dictionary(Of Object, Object)
        Dim dctJcts As Dictionary(Of Object, Object)
        Dim dctEdgeID As Dictionary(Of Object, Object)
        Dim lMaxEdgeOID As Long
        dctJoinEdges = New Dictionary(Of Object, Object)
        dctJcts = New Dictionary(Of Object, Object)
        dctEdgeID = New Dictionary(Of Object, Object)

        '  Set dctEdgeID = New Dictionary
        'calcEdgesAtJoint m_edgeShp, g_INode, g_JNode, dctJoinEdges, dctJcts, dctEdgeID
        calcEdgesAtJoint(m_edgeShp, m_junctShp, g_INode, g_JNode, dctJoinEdges, dctJcts, dctEdgeID, lMaxEdgeOID)

        WriteLogLine(Now() & " finish collecting edges at junctions that will be considered dissolving")
        WriteLogLine("junctions count=" & dctJcts.Count)
        WriteLogLine("edges count=" & dctEdgeID.Count)

        Dim pWorkspaceEdit As IWorkspaceEdit, pWorkspaceEdit2 As IWorkspaceEdit
        pWorkspaceEdit = g_FWS
        pWorkspaceEdit.StartEditing(False)
        pWorkspaceEdit.StartEditOperation()

        Dim pFeat As IFeature
        Dim sNode As String, sEdgeID() As String
        '  Do Until pJunctFeat Is Nothing
        Dim sNodes()
        Dim lJctCt As Long
        Dim pSFilt As ISpatialFilter
        Dim lDissolvedEdges As Long, lSelectedEdges As Long, lCreatedEdges As Long

        lSelectedEdges = dctEdgeID.Count
        lDissolvedEdges = 0
        lCreatedEdges = 0

        pSFilt = New SpatialFilter
        sNodes = dctJcts.Keys.ToArray

        pQFj = New QueryFilter
        pQF = New QueryFilter
        pQF2 = New QueryFilter

        'Try


        For lJctCt = 0 To dctJcts.Count - 1
            GC.Collect()
            Try


                If pCounter = 500 Then
                    WriteLogLine("pCounter reaches 500")
                    '[011806] pan pCounter appears to serve no purpose
                    pWorkspaceEdit.StopEditOperation()
                    pWorkspaceEdit.StopEditing(True)
                    pCounter = 1
                    WriteLogLine("saved edits" & "  " & Now)
                    pWorkspaceEdit.StartEditing(False)
                    pWorkspaceEdit.StartEditOperation()
                End If
            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
            pCounter = pCounter + 1
            '    dnode = pJunctFeat.value(idxJunctID)
            '    sNode = CStr(dnode)
            sNode = sNodes(lJctCt)
            dnode = Val(sNode)
            match = True
            edgesdelete = False
            '    If dnode = 1023 Then MsgBox "here"

            If Not dctNoThinNodes.ContainsKey(sNode) Then
                If dctJoinEdges.ContainsKey(sNode) Then
                    sEdgeID = Split(dctJoinEdges.Item(sNode), ",")

                    If UBound(sEdgeID) - LBound(sEdgeID) <> 1 Then
                        noMergeCount = noMergeCount + 1
                    Else '2 edges connected at this joint
                        If dctThinEdges.ContainsKey(sEdgeID(0)) And dctThinEdges.ContainsKey(sEdgeID(1)) And (Not (dctNoThinEdges.ContainsKey(sEdgeID(0)) Or dctNoThinEdges.ContainsKey(sEdgeID(1)))) Then

                            '[061907] jaf: if the two edge ID's are equal we have a problem...
                            If sEdgeID(0) = sEdgeID(1) Then
                                WriteLogLine("Dissolve found EdgeID=" & sEdgeID(0) & " CONNECTED TO ITSELF")
                            Else
                                '[061907] jaf: the two edges are NOT the same, OK to proceed
                                pFeat2 = dctEdgeID.Item(sEdgeID(0))
                                pNextFeat = dctEdgeID.Item(sEdgeID(1))

                                If Not (pFeat2 Is Nothing Or pNextFeat Is Nothing) Then
                                    WriteLogLine("potential dnode " & CStr(dnode) & "  " & Now)
                                    'now check if junction really real
                                    '                    Set pFeat2 = m_edgeShp.GetFeature(Val(sEdgeID(0)))
                                    '                    Set pNextFeat = m_edgeShp.GetFeature(Val(sEdgeID(1)))
                                    dir = False
                                    If pFeat2.Value(idxJ) = dnode Then
                                        If pNextFeat.Value(idxJ) = dnode Then 'then deleting both jnodes
                                            newI = pFeat2.Value(idxN)
                                            newJ = pNextFeat.Value(idxN)
                                        Else 'then deleting jnode of first edge and inode of second
                                            newJ = pNextFeat.Value(idxJ)
                                            newI = pFeat2.Value(idxN)
                                            dir = True
                                        End If
                                    Else
                                        newJ = pFeat2.Value(idxJ)
                                        If pNextFeat.Value(idxJ) = dnode Then 'then deleting inode of first edge and jnode of second
                                            newI = pNextFeat.Value(idxN)
                                            dir = True
                                        Else 'then deleting both inodes
                                            newI = pNextFeat.Value(idxJ)
                                        End If
                                    End If

                                    'compare attributes to detect if we have "real" pseudonode
                                    'pQF2.WhereClause = g_PSRCEdgeID & "=" & pFeat2.value(idxEdgeID) & " OR " & g_PSRCEdgeID & "=" & pNextFeat.value(idxEdgeID)

                                    'Set pQF2 = New QueryFilter
                                    'pQF2.WhereClause = g_PSRCEdgeID & "=" & sEdgeID(0) & " OR " & g_PSRCEdgeID & "=" & sEdgeID(1)
                                    'Set pRow = pTC.NextRow
                                    'Set pNextRow = pTC.NextRow
                                    pQF2.WhereClause = g_PSRCEdgeID & "=" & sEdgeID(0)
                                    If IsDBNull(pFeat2.Value(pFeat2.Fields.FindField("Updated1"))) Then
                                        pTC = pTblMode.Search(pQF2, False)
                                        pRow = pTC.NextRow
                                    Else
                                        pRow = getAttributesRow(pFeat2, Nothing, pTblPrjEdgeAtt, pEvtLine)

                                    End If

                                    'End If


                                    pQF2.WhereClause = g_PSRCEdgeID & "=" & sEdgeID(1)
                                    If IsDBNull(pNextFeat.Value(pFeat2.Fields.FindField("Updated1"))) Then
                                        pTC = pTblMode.Search(pQF2, False)
                                        pNextRow = pTC.NextRow
                                    Else


                                        pNextRow = getAttributesRow(pNextFeat, Nothing, pTblPrjEdgeAtt, pEvtLine)


                                    End If


                                    '                    If dnode = 144794 Then MsgBox "here"
                                    If pRow.Fields.FieldCount <> pNextRow.Fields.FieldCount Then
                                        'MsgBox("not equal")
                                    End If


                                    If dir = True Then
                                        For i = 1 To pRow.Fields.FieldCount - 1
                                            'some of the attribute tables might have one less filed than mode attributes
                                            If i <= pNextRow.Fields.FieldCount - 1 Then
                                                '[081208] sec: added fields dateCreated, DateLastUpdated, and LastEditor as fields that should not be compared
                                                If UCase(pRow.Fields.Field(i).Name) <> "PSRCEDGEID" And _
                                                    UCase(pRow.Fields.Field(i).Name) <> "IJFFS" And UCase(pRow.Fields.Field(i).Name) <> "JIFFS" And UCase(pRow.Fields.Field(i).Name) <> "DATELASTUPDATED" And UCase(pRow.Fields.Field(i).Name) <> "DATECREATED" And UCase(pRow.Fields.Field(i).Name) <> "LASTEDITOR" Then

                                                    'SEC- see if either field is null, consider this still a match for now

                                                    If IsDBNull(pRow.Value(i)) Or IsDBNull(pNextRow.Value(i)) Then


                                                    ElseIf (pRow.Value(i) <> pNextRow.Value(i)) Then
                                                        match = False
                                                        WriteLogLine(pRow.Fields.Field(i).Name + " diff")
                                                    End If
                                                    'End If
                                                End If
                                            End If
                                        Next i
                                    Else 'need to compare IJ to JI to match direction
                                        Dim j As Long, l As Long, t As Long
                                        For j = 0 To 1
                                            If j = 0 Then
                                                dir2 = "IJ"
                                                dirNext = "JI"
                                            Else
                                                dir2 = "JI"
                                                dirNext = "IJ"
                                            End If

                                            'begin attribute comparison by directions
                                            index2 = pRow.Fields.FindField(dir2 + "SpeedLimit")
                                            indexNext = pNextRow.Fields.FindField(dirNext + "SpeedLimit")
                                            If (pRow.Value(index2) <> pNextRow.Value(indexNext)) Then
                                                match = False
                                                WriteLogLine("dir speed diff")
                                            End If
                                            'not comparing FFS any more 06/16/05
                                            'index2 = pRow.Fields.FindField(dir2 + "FFS")
                                            ' indexnext = pNextRow.Fields.FindField(dirnext + "FFS")
                                            ' If (pRow.value(index2) <> pNextRow.value(indexnext)) Then
                                            '    match = False
                                            'WriteLogLine "dir ffs diff"
                                            ' End If
                                            index2 = pRow.Fields.FindField(dir2 + "VDFunc")
                                            indexNext = pNextRow.Fields.FindField(dirNext + "VDFunc")
                                            If (pRow.Value(index2) <> pNextRow.Value(indexNext)) Then
                                                match = False
                                                WriteLogLine("dir vdf diff")
                                            End If

                                            index2 = pRow.Fields.FindField(dir2 + "SideWalks")
                                            indexNext = pNextRow.Fields.FindField(dirNext + "SideWalks")
                                            If (pRow.Value(index2) <> pNextRow.Value(indexNext)) Then
                                                match = False
                                                WriteLogLine("dir side diff")
                                            End If
                                            index2 = pRow.Fields.FindField(dir2 + "BikeLanes")
                                            indexNext = pNextRow.Fields.FindField(dirNext + "BikeLanes")
                                            If (pRow.Value(index2) <> pNextRow.Value(indexNext)) Then
                                                match = False
                                                WriteLogLine("dir bike diff")
                                            End If
                                            For l = 0 To 3
                                                If l > 1 Then
                                                    index2 = pRow.Fields.FindField(dir2 + "Lanes" + lType(l))
                                                    indexNext = pNextRow.Fields.FindField(dirNext + "Lanes" + lType(l))
                                                    If (pRow.Value(index2) <> pNextRow.Value(indexNext)) Then
                                                        match = False
                                                        WriteLogLine("dir lanes diff")
                                                    End If
                                                Else
                                                    index2 = pRow.Fields.FindField(dir2 + "LaneCap" + lType(l))
                                                    indexNext = pNextRow.Fields.FindField(dirNext + "LaneCap" + lType(l))
                                                    If (pRow.Value(index2) <> pNextRow.Value(indexNext)) Then
                                                        match = False
                                                        WriteLogLine("dir lanecap diff")
                                                    End If
                                                    For t = 0 To 4
                                                        index2 = pRow.Fields.FindField(dir2 + "Lanes" + lType(l) + timePd(t))
                                                        indexNext = pNextRow.Fields.FindField(dirNext + "Lanes" + lType(l) + timePd(t))

                                                        If (pRow.Value(index2) <> pNextRow.Value(indexNext)) Then
                                                            match = False
                                                            WriteLogLine("lanes diff")
                                                        End If
                                                    Next t
                                                End If
                                            Next l
                                        Next j
                                        pRow = Nothing
                                        pNextRow = Nothing
                                    End If
                                    'SEC- see if either field is null, consider this still a match for now
                                    If IsDBNull(pFeat2.Value(upIndex)) Or IsDBNull(pNextFeat.Value(upIndex)) Then
                                    ElseIf pFeat2.Value(upIndex) <> pNextFeat.Value(upIndex) Then
                                        match = False
                                        WriteLogLine("update diff")
                                    End If
                                    If IsDBNull(pFeat2.Value(idxMode)) Or IsDBNull(pNextFeat.Value(idxMode)) Then
                                    ElseIf pFeat2.Value(idxMode) <> pNextFeat.Value(idxMode) Then
                                        match = False
                                        WriteLogLine("modes diff")
                                    End If

                                    If IsDBNull(pFeat2.Value(idxLT)) Or IsDBNull(pNextFeat.Value(idxLT)) Then
                                    ElseIf pFeat2.Value(idxLT) <> pNextFeat.Value(idxLT) Then
                                        match = False
                                        WriteLogLine("linktype diff")
                                    End If

                                    i = pFeat2.Fields.FindField("FacilityTy")
                                    If IsDBNull(pFeat2.Value(idxFT)) Or IsDBNull(pNextFeat.Value(idxFT)) Then
                                    ElseIf pFeat2.Value(idxFT) <> pNextFeat.Value(idxFT) Then
                                        match = False
                                        WriteLogLine("fac diff")
                                    End If


                                    If (match = False) Then
                                        attMismatchCount = attMismatchCount + 1
                                    Else
                                        WriteLogLine("found match " & Now)

                                        '[073007] hyu: new appraoch to merge edges:
                                        '[073007] hyu: consider the geometry orientation when merging two edges. Replace the old approach of using iTopologicalOperator.ConstructUnion.
                                        pPolyline = MergeEdges(pNextFeat, pFeat2)

                                        '                    Set pGeometryBag = New GeometryBag
                                        '                    Set pGeometry = pNextFeat.ShapeCopy
                                        '                    pGeometryBag.AddGeometry pGeometry
                                        '                    Set pGeometry = pFeat2.ShapeCopy
                                        '                    pGeometryBag.AddGeometry pGeometry
                                        '
                                        'update the merge count for the log
                                        If dir Then
                                            mergeCountIJ = mergeCountIJ + 1
                                        Else
                                            mergeCountJI = mergeCountJI + 1
                                        End If

                                        'merge the edge features
                                        '                    Set pPolyline = New Polyline
                                        '                    Set pTopo = pPolyline
                                        '                    pTopo.ConstructUnion pGeometryBag
                                        pNewFeat = m_edgeShp.CreateFeature
                                        pNewFeat.Shape = pPolyline
                                        '                        pnewfeat.Store
                                        pFlds = pNewFeat.Fields
                                        For i = 1 To pFlds.FieldCount - 1
                                            lIndex = pFeat2.Fields.FindField(pFlds.Field(i).Name)
                                            If (lIndex <> -1) Then
                                                '[032006] hyu: change the if condition
                                                '                                If UCase(pFlds.field(i).name) = "SCEN_LINK" Then
                                                '                                    MsgBox 1
                                                '                                End If
                                                If (pFlds.Field(i).Type <> esriFieldType.esriFieldTypeOID And pFlds.Field(i).Type <> esriFieldType.esriFieldTypeGeometry _
                                                    And UCase(pFlds.Field(i).Name) <> UCase("Shape")) Then
                                                    'If (pFlds.field(i).Type <> esriFieldTypeOID And Not pFlds.field(i).Type = esriFieldTypeGeometry Or pFlds.field(i).name <> "Shape") Then
                                                    'inode and jnode need to be set differently
                                                    If UCase((pFlds.Field(i).Name)) = "INODE" Then
                                                        pNewFeat.Value(i) = newI
                                                    ElseIf UCase((pFlds.Field(i).Name)) = "JNODE" Then
                                                        pNewFeat.Value(i) = newJ
                                                        '[122905] pan--SHAPE.len is internal field not editable
                                                        'ElseIf (pFlds.field(i).name = "SHAPE.len") Then
                                                        '  pnewfeat.value(i) = pPolyline.length
                                                        'Else
                                                        '  pnewfeat.value(i) = pFeat2.value(lindex)
                                                    Else
                                                        If pNewFeat.Fields.Field(i).Editable Then pNewFeat.Value(i) = pFeat2.Value(i)
                                                    End If
                                                End If

                                            End If
                                        Next i

                                        If IsDBNull(pNewFeat.Value(pNewFeat.Fields.FindField("SCENARIOID"))) Then
                                            'MsgBox "Null ScenarioID in new feature",,"NonFatal Error"
                                        End If

                                        'pan Edit session incompatible with Store
                                        'pan 12-16-05 re-inserted store and edge delete
                                        pNewFeat.Value(pNewFeat.Fields.FindField("Dissolve")) = 1
                                        pNewFeat.Store()
                                        lCreatedEdges = lCreatedEdges + 1
                                        edgesdelete = True

                                        updateEdgesAtJoint(dctJoinEdges, pFeat2, pNextFeat, pNewFeat)
                                        '                        dctEdgeID.Item(sEdgeID(0)) = 1
                                        '                        dctEdgeID.Item(sEdgeID(1)) = 1
                                        If pFeat2.OID <= lMaxEdgeOID Then lDissolvedEdges = lDissolvedEdges + 1
                                        If pNextFeat.OID <= lMaxEdgeOID Then lDissolvedEdges = lDissolvedEdges + 1

                                        dctEdgeID.Item(sEdgeID(0)) = pNewFeat
                                        dctEdgeID.Remove(sEdgeID(1))

                                        '   WriteLogLine "delete edges oid=" & pFeat2.OID & ", " & pNextFeat.OID
                                        pFeat2.Delete()
                                        pNextFeat.Delete()

                                        'find junction node and delete it

                                        '[122905] pan--I changed check to gt because want to delete this junction
                                        'If idxJunctID < 0 Then
                                        pJunctFeat = dctJcts.Item(sNode)

                                        '                        pQFj.WhereClause = g_PSRCJctID & "=" & sNode
                                        '                        Set pFCj = m_junctShp.Search(pQFj, False)
                                        '                        Set pJunctFeat = pFCj.NextFeature

                                        If Not pJunctFeat Is Nothing Then
                                            If fVerboseLog Then
                                                If idxJunctID > 0 Then
                                                    WriteLogLine("Deleting intrmed. feature junct OID=" & pJunctFeat.OID)
                                                Else
                                                    WriteLogLine("Deleting intrmed. feature junctID=" & pJunctFeat.Value(idxJunctID))
                                                End If
                                            End If
                                            pJunctFeat.Delete()
                                        End If

                                        'MsgBox pSSet.count

                                    End If '(match = False)
                                End If 'Not (pFeat2 Is Nothing Or pNextFeat Is Nothing)
                            End If 'edge connect to itself: sEdgeID(0) = sEdgeID(1)
                        End If  'match = true:  dctThinEdges.Exists(sEdgeID(0)) And dctThinEdges.Exists(sEdgeID(1)) ...
                    End If 'UBound(sEdgeID) - LBound(sEdgeID) <> 1
                End If 'dctJoinEdges.Exists(sNode)

            End If 'Not dctNoThinNodes.Exists(sNode)

            pFCj = Nothing
            pTC = Nothing
            pFC = Nothing
            pFC2 = Nothing
        Next lJctCt
        'Catch ex As Exception
        'MessageBox.Show(ex.ToString)
        'End Try
        pWorkspaceEdit.StopEditOperation()
        pWorkspaceEdit.StopEditing(True)

        _map.ClearSelection()
        Dim PMx As IMxDocument

        _activeView.Refresh()

        pFLayerE = Nothing
        pFLayerJ = Nothing
        pJunctFeat = Nothing
        pFeat2 = Nothing
        pNextFeat = Nothing
        pNewFeat = Nothing
        pQFj = Nothing
        pFCj = Nothing
        pRow = Nothing
        pNextRow = Nothing
        pTblMode = Nothing
        pTolls = Nothing
        pTurn = Nothing
        '   Set tblTSeg = Nothing
        pWorkspaceEdit = Nothing

        dctNoThinEdges.Clear()
        dctThinEdges.Clear()
        dctNoThinNodes.Clear()
        dctJcts.Clear()

        dctNoThinEdges = Nothing
        dctThinEdges = Nothing
        dctNoThinNodes = Nothing
        dctJcts = Nothing

        WriteLogLine("FINISHED dissolveTransEdges2 at " & Now())
        WriteLogLine("Junctions selected= " & dctJoinEdges.Count)
        WriteLogLine("Same dir thinned= " & mergeCountIJ & "; Opp dir thinned= " & mergeCountJI)
        WriteLogLine("Junctions thinned= " & mergeCountIJ + mergeCountJI)
        WriteLogLine("Junctions geometrically eligible to thin but not thinned due to attribute mismatch= " & attMismatchCount)
        WriteLogLine("Junctions geometrically ineligible to thin= " & noMergeCount)
        WriteLogLine("Junctions final= " & m_junctShp.FeatureCount(Nothing))
        '  t = 0
        '  For l = 0 To dctEdgeID.count - 1
        '    If UCase(TypeName(dctEdgeID.Item(l))) = "INTEGER" Then t = t + 1
        '    'If dctEdgeID.Items(l) = 1 Then t = t + 1
        '  Next l

        WriteLogLine("Edges selected= " & lSelectedEdges)
        WriteLogLine("Edges merged= " & lDissolvedEdges)

        '  Set pQF = New QueryFilter
        '  pQF.WhereClause = "Dissolve=1"
        '  WriteLogLine "Edges created= " & m_edgeShp.featurecount(pQF)
        '  WriteLogLine "Edges final= " & m_edgeShp.featurecount(Nothing)

        WriteLogLine("Edges created= " & lCreatedEdges)
        WriteLogLine("Edges final= " & dctEdgeID.Count)
        pQF = Nothing

        WriteLogLine(m_edgeShp.AliasName + " " + CStr(m_edgeShp.FeatureCount(Nothing)))
        'Note: needs to be a round two version of this based on atrributes that don't match but are not significant
        'need to get these attributes!
        Exit Sub

ErrChk:
        CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.dissolveTransEdges2")
        MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.dissolveTransEdges2")
        'Print #1, Err.Description, vbInformation, "dissolveTransEdges2"

        '  If pWorkspaceEdit.IsBeingEdited Then
        pWorkspaceEdit.StopEditOperation()
        pWorkspaceEdit.StopEditing(True)
        '  End If


        pFLayerE = Nothing
        pFLayerJ = Nothing
        pJunctFeat = Nothing
        pFeat2 = Nothing
        pNextFeat = Nothing
        pNewFeat = Nothing
        pQFj = Nothing
        pFCj = Nothing
        pRow = Nothing
        pNextRow = Nothing
        pTblMode = Nothing
        pTolls = Nothing
        pTurn = Nothing
        '   Set tblTSeg = Nothing
        pWorkspaceEdit = Nothing

    End Sub
    Public Sub DissolveByFT()
        '[022506] hyu: revised function to replace [dissolveTransEdges2_0]

        'Retains junctions and edges used by transit, tolls and turns
        'performs "pseudonode" thinning of the intermediate edge shp file
        'modified to handle toll, turn coincidence with thinning by Melinda 5-25-05
        'retrofitted (minimally) by Jeff with our logging stuff
        'added transit thinning by Melinda 6-3-05 pan

        'on error GoTo ErrChk

        WriteLogLine("")
        WriteLogLine("======================================")
        WriteLogLine("dissolveTransEdges2 started " & Now())
        WriteLogLine("======================================")
        WriteLogLine("")
        If (strLayerPrefix = "SDE") Then
            Dim pwf As IWorkspaceFactory
            pwf = New SdeWorkspaceFactory
        End If

        Dim pCounter As Long
        Dim pFLayerE As IFeatureLayer, pFLayerJ As IFeatureLayer
        pFLayerE = New FeatureLayer
        pFLayerE.FeatureClass = m_edgeShp 'Intermediate Edge
        WriteLogLine(m_edgeShp.AliasName + CStr(m_edgeShp.FeatureCount(Nothing)))
        pFLayerJ = New FeatureLayer
        pFLayerJ.FeatureClass = m_junctShp 'Intermediate Junction

        Dim pTblMode As ITable, pTolls As ITable
        Dim pTurn As IFeatureLayer
        Dim pTblPrjEdgeAtt As ITable, pEvtLine As ITable

        pTblPrjEdgeAtt = get_TableClass(m_layers(24))
        Try


            pTblMode = get_TableClass(m_layers(2)) 'modeAttributes
            pTolls = get_TableClass(m_layers(9)) 'modeTolls
            pTurn = get_FeatureLayer2(m_layers(12)) 'TurnMovements
            pEvtLine = get_TableClass(m_layers(4)) 'use if future event project
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        '[041607]hyu: now dealing with transit points feature layer instead of segment table
        '  Dim tblTSeg As ITable
        '  Set tblTSeg = get_TableClass(m_layers(14)) 'tbltransitsegments
        Dim pTrPoints As IFeatureLayer
        pTrPoints = get_FeatureLayer2(m_layers(23))

        Dim pFC As IFeatureCursor, pFC2 As IFeatureCursor
        Dim pFCj As IFeatureCursor
        Dim pQF As IQueryFilter, pQFt As IQueryFilter, pQF2 As IQueryFilter
        Dim pQFj As IQueryFilter

        Dim pGeometry As IGeometry, pNewGeometry As IGeometry
        Dim pGeometryBag As IGeometryCollection
        pGeometryBag = New GeometryBag
        Dim pTopo As ITopologicalOperator
        Dim pPolyline As IPolyline

        Dim pFeat2 As IFeature
        Dim pNextFeat As IFeature, pNewFeat As IFeature
        Dim pJunctFeat As IFeature
        Dim pRow As IRow, pNextRow As IRow
        Dim pTC As ICursor
        Dim pFlds As IFields

        Dim psrcId As String
        Dim i As Long
        Dim fcountI, fcountIJ, mergeCountIJ, mergeCountJI As Long, noMergeCount As Long, attMismatchCount As Long

        Dim match As Boolean
        match = True
        Dim edgesdelete As Boolean
        edgesdelete = False

        Dim newI As Long, newJ As Long

        '  Dim pWorkspaceEdit As IWorkspaceEdit, pWorkspaceEdit2 As IWorkspaceEdit
        '  Set pWorkspaceEdit = g_FWS
        '  pWorkspaceEdit.StartEditing False
        '  pWorkspaceEdit.StartEditOperation

        '  'this one just loops thru junct layer- should be faster
        '  Set pFCj = m_junctShp.Search(Nothing, False)
        '  Set pJunctFeat = pFCj.NextFeature

        Dim lSFld As Integer, dnode As Long, idIndex As Integer
        Dim lIndex As Integer, upIndex As Integer
        Dim idxN As Integer, idxJ As Integer
        Dim idxMode As Integer, idxLT As Integer, idxFT As Integer
        Dim idxJunctID As Integer, idxEdgeID As Integer
        Dim ifld As Integer
        idxJunctID = m_junctShp.FindField("PSRCJunctID")
        '
        '[122305]pan--revised field names for current SDE layer match
        '
        With m_edgeShp
            idxEdgeID = .FindField(g_PSRCEdgeID)
            idIndex = .FindField("PSRC_E2ID")
            upIndex = .FindField("Updated1")
            idxN = .FindField(g_INode)
            idxJ = .FindField(g_JNode)
            idxMode = .FindField(g_Modes)
            idxLT = .FindField(g_LinkType)
            idxFT = .FindField(g_FacType)
        End With

        Dim m As Integer
        Dim same As String, opposite As String
        Dim dir2 As String, dirNext As String
        Dim index2 As Integer, indexNext As Integer
        Dim dir As Boolean
        Dim lType(4) As String
        lType(0) = "GP"
        lType(1) = "HOV"
        lType(2) = "TR"
        lType(3) = "TK"
        Dim timePd(5) As String
        timePd(0) = "AM"
        timePd(1) = "MD"
        timePd(2) = "PM"
        timePd(3) = "EV"
        timePd(4) = "NI"

        mergeCountIJ = 0
        mergeCountJI = 0
        noMergeCount = 0
        attMismatchCount = 0
        pCounter = 1

        '[022406] hyu: get those nodes/edges that can/cannot be thinned into dictionaries.
        Dim dctNoThinNodes As Dictionary(Of Object, Object), dctNoThinEdges As Dictionary(Of Object, Object), dctThinEdges As Dictionary(Of Object, Object)
        'dctNoThinNodes consists of nodes cannot be thinned
        'dctThinEdges consists of edges that are possible be thinned.

        dctNoThinNodes = New Dictionary(Of Object, Object)
        dctThinEdges = New Dictionary(Of Object, Object)
        dctNoThinEdges = New Dictionary(Of Object, Object)
        WriteLogLine(Now() & " start collecting nodes and edges ")

        '  [050306] hyu: transitSegment
        '  getNodes tblTSeg, g_PSRCEdgeID, dctNoThinEdges

        '[041607]hyu: dealing w/ transit points instead of segments table
        'getNodes tblTSeg, g_INode, dctNoThinNodes, "SUBSTRING(CAST(LineID AS CHARACTER), 2, 2) = '" & Right(CStr(inserviceyear), 2) & "' AND SegOrder=1"
        'getNodes tblTSeg, g_JNode, dctNoThinNodes, "SUBSTRING(CAST(LineID AS CHARACTER), 2, 2) = '" & Right(CStr(inserviceyear), 2) & "' AND change>0"
        '[051407] jaf: use SQL with SDE but not with Access
        If strLayerPrefix = "SDE" Then
            'getNodes(pTrPoints, g_PSRCJctID, dctNoThinNodes, "SUBSTRING(CAST(LineID AS CHARACTER), 2, 2) = '" & Microsoft.VisualBasic.Right(CStr(inserviceyear), 2) & "'")
            'getPRNodes pFLayerJ.FeatureClass, "Scen_Node", "P_RStalls", dctNoThinNodes
        Else
            getNodes(pTrPoints, g_PSRCJctID, dctNoThinNodes, "MID(Cstr(LineID), 2, 2) = '" & Microsoft.VisualBasic.Right(CStr(inserviceyear), 2) & "'")
        End If

        WriteLogLine(Now() & " finish collecting junctions being transit points: " & dctNoThinNodes.Count)

        getNodes(pTblMode, g_PSRCEdgeID, dctThinEdges, , True)
        WriteLogLine(Now() & " finish collecting edges can be thinned: " & dctThinEdges.Count)

        getNodes(pTolls, g_PSRCEdgeID, dctNoThinEdges)
        WriteLogLine(Now() & " finish collecting edges can NOT be thinned: " & dctNoThinEdges.Count)

        '[060210] SEC- Need to preserve Turn Edges & "Retain" Junctions
        getNodes(pFLayerJ.FeatureClass, "PSRCJUNCTID", dctNoThinNodes, "JunctionType = 10")
        getNodes(pTurn, "FrEdgeID", dctNoThinEdges)

        getNodes(pTurn, "ToEdgeID", dctNoThinEdges)

        getNodes(pTurn, g_PSRCJctID, dctNoThinNodes)

        WriteLogLine(Now() & " finish collecting junction can NOT be thinned (including transit): " & dctNoThinNodes.Count)

        '[082409] SEC: Added Park and Rides
        getNodes(m_edgeShp, g_PSRCEdgeID, dctNoThinEdges, g_FacType & "=16 or " & g_FacType & "=17") '[042706] hyu: exclude centroid connector + park and rides from dissolving
        WriteLogLine(Now() & " finish collecting edges can NOT be thinned : " & dctNoThinEdges.Count)

        '[022406] hyu: stats how many edges joining to a junction
        Dim dctJoinEdges As Dictionary(Of Object, Object)
        Dim dctJcts As Dictionary(Of Object, Object)
        Dim dctEdgeID As Dictionary(Of Object, Object)
        Dim lMaxEdgeOID As Long
        dctJoinEdges = New Dictionary(Of Object, Object)
        dctJcts = New Dictionary(Of Object, Object)
        dctEdgeID = New Dictionary(Of Object, Object)

        '  Set dctEdgeID = New Dictionary(Of Object, Object)
        'calcEdgesAtJoint m_edgeShp, g_INode, g_JNode, dctJoinEdges, dctJcts, dctEdgeID
        calcEdgesAtJoint(m_edgeShp, m_junctShp, g_INode, g_JNode, dctJoinEdges, dctJcts, dctEdgeID, lMaxEdgeOID)

        WriteLogLine(Now() & " finish collecting edges at junctions that will be considered dissolving")
        WriteLogLine("junctions count=" & dctJcts.Count)
        WriteLogLine("edges count=" & dctEdgeID.Count)

        Dim pWorkspaceEdit As IWorkspaceEdit, pWorkspaceEdit2 As IWorkspaceEdit
        pWorkspaceEdit = g_FWS
        pWorkspaceEdit.StartEditing(False)
        pWorkspaceEdit.StartEditOperation()

        Dim pFeat As IFeature
        Dim sNode As String, sEdgeID() As String
        '  Do Until pJunctFeat Is Nothing
        Dim sNodes()
        Dim lJctCt As Long
        Dim pSFilt As ISpatialFilter
        Dim lDissolvedEdges As Long, lSelectedEdges As Long, lCreatedEdges As Long

        lSelectedEdges = dctEdgeID.Count
        lDissolvedEdges = 0
        lCreatedEdges = 0

        pSFilt = New SpatialFilter
        sNodes = dctJcts.Keys.ToArray
        pQFj = New QueryFilter
        pQF = New QueryFilter
        pQF2 = New QueryFilter

        Try


            For lJctCt = 0 To dctJcts.Count - 1
                GC.Collect()
                If pCounter = 500 Then
                    WriteLogLine("pCounter reaches 500")
                    '[011806] pan pCounter appears to serve no purpose
                    pWorkspaceEdit.StopEditOperation()
                    pWorkspaceEdit.StopEditing(True)
                    pCounter = 1
                    WriteLogLine("saved edits" & "  " & Now)
                    pWorkspaceEdit.StartEditing(False)
                    pWorkspaceEdit.StartEditOperation()
                End If
                pCounter = pCounter + 1
                '    dnode = pJunctFeat.value(idxJunctID)
                '    sNode = CStr(dnode)
                sNode = sNodes(lJctCt)
                dnode = Val(sNode)
                match = True
                edgesdelete = False
                '    If dnode = 1023 Then MsgBox "here"

                If Not dctNoThinNodes.ContainsKey(sNode) Then
                    If dctJoinEdges.ContainsKey(sNode) Then
                        sEdgeID = Split(dctJoinEdges.Item(sNode), ",")

                        If UBound(sEdgeID) - LBound(sEdgeID) <> 1 Then
                            noMergeCount = noMergeCount + 1
                        Else '2 edges connected at this joint
                            If dctThinEdges.ContainsKey(sEdgeID(0)) And dctThinEdges.ContainsKey(sEdgeID(1)) And (Not (dctNoThinEdges.ContainsKey(sEdgeID(0)) Or dctNoThinEdges.ContainsKey(sEdgeID(1)))) Then

                                '[061907] jaf: if the two edge ID's are equal we have a problem...
                                If sEdgeID(0) = sEdgeID(1) Then
                                    WriteLogLine("Dissolve found EdgeID=" & sEdgeID(0) & " CONNECTED TO ITSELF")
                                Else
                                    '[061907] jaf: the two edges are NOT the same, OK to proceed
                                    pFeat2 = dctEdgeID.Item(sEdgeID(0))
                                    pNextFeat = dctEdgeID.Item(sEdgeID(1))

                                    If Not (pFeat2 Is Nothing Or pNextFeat Is Nothing) Then
                                        WriteLogLine("potential dnode " & CStr(dnode) & "  " & Now)
                                        'now check if junction really real
                                        '                    Set pFeat2 = m_edgeShp.GetFeature(Val(sEdgeID(0)))
                                        '                    Set pNextFeat = m_edgeShp.GetFeature(Val(sEdgeID(1)))
                                        dir = False
                                        If pFeat2.Value(idxJ) = dnode Then
                                            If pNextFeat.Value(idxJ) = dnode Then 'then deleting both jnodes
                                                newI = pFeat2.Value(idxN)
                                                newJ = pNextFeat.Value(idxN)
                                            Else 'then deleting jnode of first edge and inode of second
                                                newJ = pNextFeat.Value(idxJ)
                                                newI = pFeat2.Value(idxN)
                                                dir = True
                                            End If
                                        Else
                                            newJ = pFeat2.Value(idxJ)
                                            If pNextFeat.Value(idxJ) = dnode Then 'then deleting inode of first edge and jnode of second
                                                newI = pNextFeat.Value(idxN)
                                                dir = True
                                            Else 'then deleting both inodes
                                                newI = pNextFeat.Value(idxJ)
                                            End If
                                        End If

                                        'compare attributes to detect if we have "real" pseudonode
                                        'pQF2.WhereClause = g_PSRCEdgeID & "=" & pFeat2.value(idxEdgeID) & " OR " & g_PSRCEdgeID & "=" & pNextFeat.value(idxEdgeID)

                                        'Set pQF2 = New QueryFilter
                                        'pQF2.WhereClause = g_PSRCEdgeID & "=" & sEdgeID(0) & " OR " & g_PSRCEdgeID & "=" & sEdgeID(1)
                                        'Set pRow = pTC.NextRow
                                        'Set pNextRow = pTC.NextRow
                                        pQF2.WhereClause = g_PSRCEdgeID & "=" & sEdgeID(0)
                                        If IsDBNull(pFeat2.Value(pFeat2.Fields.FindField("Updated1"))) Then
                                            pTC = pTblMode.Search(pQF2, False)
                                            pRow = pTC.NextRow
                                        Else
                                            pRow = getAttributesRow(pFeat2, Nothing, pTblPrjEdgeAtt, pEvtLine)

                                        End If

                                        'End If


                                        pQF2.WhereClause = g_PSRCEdgeID & "=" & sEdgeID(1)
                                        If IsDBNull(pNextFeat.Value(pFeat2.Fields.FindField("Updated1"))) Then
                                            pTC = pTblMode.Search(pQF2, False)
                                            pNextRow = pTC.NextRow
                                        Else


                                            pNextRow = getAttributesRow(pNextFeat, Nothing, pTblPrjEdgeAtt, pEvtLine)


                                        End If


                                        '                    If dnode = 144794 Then MsgBox "here"
                                        If pRow.Fields.FieldCount <> pNextRow.Fields.FieldCount Then
                                            'MsgBox("not equal")
                                        End If


                                        If dir = True Then
                                            For i = 1 To pRow.Fields.FieldCount - 1
                                                'some of the attribute tables might have one less filed than mode attributes
                                                If i <= pNextRow.Fields.FieldCount - 1 Then
                                                    '[081208] sec: added fields dateCreated, DateLastUpdated, and LastEditor as fields that should not be compared
                                                    If UCase(pRow.Fields.Field(i).Name) <> "PSRCEDGEID" And _
                                                        UCase(pRow.Fields.Field(i).Name) <> "IJFFS" And UCase(pRow.Fields.Field(i).Name) <> "JIFFS" And UCase(pRow.Fields.Field(i).Name) <> "DATELASTUPDATED" And UCase(pRow.Fields.Field(i).Name) <> "DATECREATED" And UCase(pRow.Fields.Field(i).Name) <> "LASTEDITOR" Then

                                                        'SEC- see if either field is null, consider this still a match for now

                                                        If IsDBNull(pRow.Value(i)) Or IsDBNull(pNextRow.Value(i)) Then


                                                        ElseIf (pRow.Value(i) <> pNextRow.Value(i)) Then
                                                            'match = False
                                                            WriteLogLine(pRow.Fields.Field(i).Name + " diff")
                                                        End If
                                                        'End If
                                                    End If
                                                End If
                                            Next i
                                        Else 'need to compare IJ to JI to match direction
                                            Dim j As Long, l As Long, t As Long
                                            For j = 0 To 1
                                                If j = 0 Then
                                                    dir2 = "IJ"
                                                    dirNext = "JI"
                                                Else
                                                    dir2 = "JI"
                                                    dirNext = "IJ"
                                                End If

                                                'begin attribute comparison by directions
                                                index2 = pRow.Fields.FindField(dir2 + "SpeedLimit")
                                                indexNext = pNextRow.Fields.FindField(dirNext + "SpeedLimit")
                                                If (pRow.Value(index2) <> pNextRow.Value(indexNext)) Then
                                                    'match = False
                                                    WriteLogLine("dir speed diff")
                                                End If
                                                'not comparing FFS any more 06/16/05
                                                'index2 = pRow.Fields.FindField(dir2 + "FFS")
                                                ' indexnext = pNextRow.Fields.FindField(dirnext + "FFS")
                                                ' If (pRow.value(index2) <> pNextRow.value(indexnext)) Then
                                                '    match = False
                                                'WriteLogLine "dir ffs diff"
                                                ' End If
                                                index2 = pRow.Fields.FindField(dir2 + "VDFunc")
                                                indexNext = pNextRow.Fields.FindField(dirNext + "VDFunc")
                                                If (pRow.Value(index2) <> pNextRow.Value(indexNext)) Then
                                                    'match = False
                                                    WriteLogLine("dir vdf diff")
                                                End If

                                                index2 = pRow.Fields.FindField(dir2 + "SideWalks")
                                                indexNext = pNextRow.Fields.FindField(dirNext + "SideWalks")
                                                If (pRow.Value(index2) <> pNextRow.Value(indexNext)) Then
                                                    'match = False
                                                    WriteLogLine("dir side diff")
                                                End If
                                                index2 = pRow.Fields.FindField(dir2 + "BikeLanes")
                                                indexNext = pNextRow.Fields.FindField(dirNext + "BikeLanes")
                                                If (pRow.Value(index2) <> pNextRow.Value(indexNext)) Then
                                                    'match = False
                                                    WriteLogLine("dir bike diff")
                                                End If
                                                For l = 0 To 3
                                                    If l > 1 Then
                                                        index2 = pRow.Fields.FindField(dir2 + "Lanes" + lType(l))
                                                        indexNext = pNextRow.Fields.FindField(dirNext + "Lanes" + lType(l))
                                                        If (pRow.Value(index2) <> pNextRow.Value(indexNext)) Then
                                                            'match = False
                                                            WriteLogLine("dir lanes diff")
                                                        End If
                                                    Else
                                                        index2 = pRow.Fields.FindField(dir2 + "LaneCap" + lType(l))
                                                        indexNext = pNextRow.Fields.FindField(dirNext + "LaneCap" + lType(l))
                                                        If (pRow.Value(index2) <> pNextRow.Value(indexNext)) Then
                                                            'match = False
                                                            WriteLogLine("dir lanecap diff")
                                                        End If
                                                        For t = 0 To 4
                                                            index2 = pRow.Fields.FindField(dir2 + "Lanes" + lType(l) + timePd(t))
                                                            indexNext = pNextRow.Fields.FindField(dirNext + "Lanes" + lType(l) + timePd(t))

                                                            If (pRow.Value(index2) <> pNextRow.Value(indexNext)) Then
                                                                'match = False
                                                                WriteLogLine("lanes diff")
                                                            End If
                                                        Next t
                                                    End If
                                                Next l
                                            Next j
                                            pRow = Nothing
                                            pNextRow = Nothing
                                        End If
                                        'SEC- see if either field is null, consider this still a match for now
                                        If IsDBNull(pFeat2.Value(upIndex)) Or IsDBNull(pNextFeat.Value(upIndex)) Then
                                        ElseIf pFeat2.Value(upIndex) <> pNextFeat.Value(upIndex) Then
                                            'match = False
                                            WriteLogLine("update diff")
                                        End If
                                        If IsDBNull(pFeat2.Value(idxMode)) Or IsDBNull(pNextFeat.Value(idxMode)) Then
                                        ElseIf pFeat2.Value(idxMode) <> pNextFeat.Value(idxMode) Then
                                            'match = False
                                            WriteLogLine("modes diff")
                                        End If

                                        If IsDBNull(pFeat2.Value(idxLT)) Or IsDBNull(pNextFeat.Value(idxLT)) Then
                                        ElseIf pFeat2.Value(idxLT) <> pNextFeat.Value(idxLT) Then
                                            'match = False
                                            WriteLogLine("linktype diff")
                                        End If

                                        i = pFeat2.Fields.FindField("FacilityTy")
                                        If IsDBNull(pFeat2.Value(idxFT)) Or IsDBNull(pNextFeat.Value(idxFT)) Then
                                        ElseIf pFeat2.Value(idxFT) <> pNextFeat.Value(idxFT) Then
                                            'match = False
                                            WriteLogLine("fac diff")
                                        End If


                                        If (match = False) Then
                                            attMismatchCount = attMismatchCount + 1
                                        Else
                                            WriteLogLine("found match " & Now)

                                            '[073007] hyu: new appraoch to merge edges:
                                            '[073007] hyu: consider the geometry orientation when merging two edges. Replace the old approach of using iTopologicalOperator.ConstructUnion.
                                            pPolyline = MergeEdges(pNextFeat, pFeat2)

                                            '                    Set pGeometryBag = New GeometryBag
                                            '                    Set pGeometry = pNextFeat.ShapeCopy
                                            '                    pGeometryBag.AddGeometry pGeometry
                                            '                    Set pGeometry = pFeat2.ShapeCopy
                                            '                    pGeometryBag.AddGeometry pGeometry
                                            '
                                            'update the merge count for the log
                                            If dir Then
                                                mergeCountIJ = mergeCountIJ + 1
                                            Else
                                                mergeCountJI = mergeCountJI + 1
                                            End If

                                            'merge the edge features
                                            '                    Set pPolyline = New Polyline
                                            '                    Set pTopo = pPolyline
                                            '                    pTopo.ConstructUnion pGeometryBag
                                            pNewFeat = m_edgeShp.CreateFeature
                                            pNewFeat.Shape = pPolyline
                                            '                        pnewfeat.Store
                                            pFlds = pNewFeat.Fields
                                            For i = 1 To pFlds.FieldCount - 1
                                                lIndex = pFeat2.Fields.FindField(pFlds.Field(i).Name)
                                                If (lIndex <> -1) Then
                                                    '[032006] hyu: change the if condition
                                                    '                                If UCase(pFlds.field(i).name) = "SCEN_LINK" Then
                                                    '                                    MsgBox 1
                                                    '                                End If
                                                    If (pFlds.Field(i).Type <> esriFieldType.esriFieldTypeOID And pFlds.Field(i).Type <> esriFieldType.esriFieldTypeGeometry _
                                                        And UCase(pFlds.Field(i).Name) <> UCase("Shape")) Then
                                                        'If (pFlds.field(i).Type <> esriFieldTypeOID And Not pFlds.field(i).Type = esriFieldTypeGeometry Or pFlds.field(i).name <> "Shape") Then
                                                        'inode and jnode need to be set differently
                                                        If UCase((pFlds.Field(i).Name)) = "INODE" Then
                                                            pNewFeat.Value(i) = newI
                                                        ElseIf UCase((pFlds.Field(i).Name)) = "JNODE" Then
                                                            pNewFeat.Value(i) = newJ
                                                            '[122905] pan--SHAPE.len is internal field not editable
                                                            'ElseIf (pFlds.field(i).name = "SHAPE.len") Then
                                                            '  pnewfeat.value(i) = pPolyline.length
                                                            'Else
                                                            '  pnewfeat.value(i) = pFeat2.value(lindex)
                                                        Else
                                                            If pNewFeat.Fields.Field(i).Editable Then pNewFeat.Value(i) = pFeat2.Value(i)
                                                        End If
                                                    End If

                                                End If
                                            Next i

                                            If IsDBNull(pNewFeat.Value(pNewFeat.Fields.FindField("SCENARIOID"))) Then
                                                'MsgBox "Null ScenarioID in new feature",,"NonFatal Error"
                                            End If

                                            'pan Edit session incompatible with Store
                                            'pan 12-16-05 re-inserted store and edge delete
                                            pNewFeat.Value(pNewFeat.Fields.FindField("Dissolve")) = 1
                                            pNewFeat.Store()
                                            lCreatedEdges = lCreatedEdges + 1
                                            edgesdelete = True

                                            updateEdgesAtJoint(dctJoinEdges, pFeat2, pNextFeat, pNewFeat)
                                            '                        dctEdgeID.Item(sEdgeID(0)) = 1
                                            '                        dctEdgeID.Item(sEdgeID(1)) = 1
                                            If pFeat2.OID <= lMaxEdgeOID Then lDissolvedEdges = lDissolvedEdges + 1
                                            If pNextFeat.OID <= lMaxEdgeOID Then lDissolvedEdges = lDissolvedEdges + 1

                                            dctEdgeID.Item(sEdgeID(0)) = pNewFeat
                                            dctEdgeID.Remove(sEdgeID(1))

                                            '   WriteLogLine "delete edges oid=" & pFeat2.OID & ", " & pNextFeat.OID
                                            pFeat2.Delete()
                                            pNextFeat.Delete()

                                            'find junction node and delete it

                                            '[122905] pan--I changed check to gt because want to delete this junction
                                            'If idxJunctID < 0 Then
                                            pJunctFeat = dctJcts.Item(sNode)

                                            '                        pQFj.WhereClause = g_PSRCJctID & "=" & sNode
                                            '                        Set pFCj = m_junctShp.Search(pQFj, False)
                                            '                        Set pJunctFeat = pFCj.NextFeature

                                            If Not pJunctFeat Is Nothing Then
                                                If fVerboseLog Then
                                                    If idxJunctID > 0 Then
                                                        WriteLogLine("Deleting intrmed. feature junct OID=" & pJunctFeat.OID)
                                                    Else
                                                        WriteLogLine("Deleting intrmed. feature junctID=" & pJunctFeat.Value(idxJunctID))
                                                    End If
                                                End If
                                                pJunctFeat.Delete()
                                            End If

                                            'MsgBox pSSet.count

                                        End If '(match = False)
                                    End If 'Not (pFeat2 Is Nothing Or pNextFeat Is Nothing)
                                End If 'edge connect to itself: sEdgeID(0) = sEdgeID(1)
                            End If  'match = true:  dctThinEdges.Exists(sEdgeID(0)) And dctThinEdges.Exists(sEdgeID(1)) ...
                        End If 'UBound(sEdgeID) - LBound(sEdgeID) <> 1
                    End If 'dctJoinEdges.Exists(sNode)

                End If 'Not dctNoThinNodes.Exists(sNode)

                pFCj = Nothing
                pTC = Nothing
                pFC = Nothing
                pFC2 = Nothing
            Next lJctCt
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        pWorkspaceEdit.StopEditOperation()
        pWorkspaceEdit.StopEditing(True)

        _map.ClearSelection()
        Dim PMx As IMxDocument

        _activeView.Refresh()

        pFLayerE = Nothing
        pFLayerJ = Nothing
        pJunctFeat = Nothing
        pFeat2 = Nothing
        pNextFeat = Nothing
        pNewFeat = Nothing
        pQFj = Nothing
        pFCj = Nothing
        pRow = Nothing
        pNextRow = Nothing
        pTblMode = Nothing
        pTolls = Nothing
        pTurn = Nothing
        '   Set tblTSeg = Nothing
        pWorkspaceEdit = Nothing

        dctNoThinEdges.Clear()
        dctThinEdges.Clear()
        dctNoThinNodes.Clear()
        dctJcts.Clear()

        dctNoThinEdges = Nothing
        dctThinEdges = Nothing
        dctNoThinNodes = Nothing
        dctJcts = Nothing

        WriteLogLine("FINISHED dissolveTransEdges2 at " & Now())
        WriteLogLine("Junctions selected= " & dctJoinEdges.Count)
        WriteLogLine("Same dir thinned= " & mergeCountIJ & "; Opp dir thinned= " & mergeCountJI)
        WriteLogLine("Junctions thinned= " & mergeCountIJ + mergeCountJI)
        WriteLogLine("Junctions geometrically eligible to thin but not thinned due to attribute mismatch= " & attMismatchCount)
        WriteLogLine("Junctions geometrically ineligible to thin= " & noMergeCount)
        WriteLogLine("Junctions final= " & m_junctShp.FeatureCount(Nothing))
        '  t = 0
        '  For l = 0 To dctEdgeID.count - 1
        '    If UCase(TypeName(dctEdgeID.Item(l))) = "INTEGER" Then t = t + 1
        '    'If dctEdgeID.Items(l) = 1 Then t = t + 1
        '  Next l

        WriteLogLine("Edges selected= " & lSelectedEdges)
        WriteLogLine("Edges merged= " & lDissolvedEdges)

        '  Set pQF = New QueryFilter
        '  pQF.WhereClause = "Dissolve=1"
        '  WriteLogLine "Edges created= " & m_edgeShp.featurecount(pQF)
        '  WriteLogLine "Edges final= " & m_edgeShp.featurecount(Nothing)

        WriteLogLine("Edges created= " & lCreatedEdges)
        WriteLogLine("Edges final= " & dctEdgeID.Count)
        pQF = Nothing

        WriteLogLine(m_edgeShp.AliasName + " " + CStr(m_edgeShp.FeatureCount(Nothing)))
        'Note: needs to be a round two version of this based on atrributes that don't match but are not significant
        'need to get these attributes!
        Exit Sub

ErrChk:
        CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.dissolveTransEdges2")
        MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.dissolveTransEdges2")
        'Print #1, Err.Description, vbInformation, "dissolveTransEdges2"

        '  If pWorkspaceEdit.IsBeingEdited Then
        pWorkspaceEdit.StopEditOperation()
        pWorkspaceEdit.StopEditing(True)
        '  End If


        pFLayerE = Nothing
        pFLayerJ = Nothing
        pJunctFeat = Nothing
        pFeat2 = Nothing
        pNextFeat = Nothing
        pNewFeat = Nothing
        pQFj = Nothing
        pFCj = Nothing
        pRow = Nothing
        pNextRow = Nothing
        pTblMode = Nothing
        pTolls = Nothing
        pTurn = Nothing
        '   Set tblTSeg = Nothing
        pWorkspaceEdit = Nothing

    End Sub

    Private Sub Dissolve_Keep_TTP2()
        '[022506] hyu: revised function to replace [dissolveTransEdges2_0]

        'Retains junctions and edges used by transit, tolls and turns
        'performs "pseudonode" thinning of the intermediate edge shp file
        'modified to handle toll, turn coincidence with thinning by Melinda 5-25-05
        'retrofitted (minimally) by Jeff with our logging stuff
        'added transit thinning by Melinda 6-3-05 pan

        'on error GoTo ErrChk

        WriteLogLine("")
        WriteLogLine("======================================")
        WriteLogLine("dissolveTransEdges2 started " & Now())
        WriteLogLine("======================================")
        WriteLogLine("")
        If (strLayerPrefix = "SDE") Then
            Dim pwf As IWorkspaceFactory
            pwf = New SdeWorkspaceFactory
        End If

        Dim pCounter As Long
        Dim pFLayerE As IFeatureLayer, pFLayerJ As IFeatureLayer
        pFLayerE = New FeatureLayer
        pFLayerE.FeatureClass = m_edgeShp 'Intermediate Edge
        WriteLogLine(m_edgeShp.AliasName + CStr(m_edgeShp.FeatureCount(Nothing)))
        pFLayerJ = New FeatureLayer
        pFLayerJ.FeatureClass = m_junctShp 'Intermediate Junction

        Dim pTblMode As ITable, pTolls As ITable
        Dim pTurn As IFeatureLayer
        Dim pTblPrjEdgeAtt As ITable, pEvtLine As ITable

        pTblPrjEdgeAtt = get_TableClass(m_layers(24))
        Try


            pTblMode = get_TableClass(m_layers(2)) 'modeAttributes
            pTolls = get_TableClass(m_layers(9)) 'modeTolls
            pTurn = get_FeatureLayer2(m_layers(12)) 'TurnMovements
            pEvtLine = get_TableClass(m_layers(4)) 'use if future event project
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        '[041607]hyu: now dealing with transit points feature layer instead of segment table
        '  Dim tblTSeg As ITable
        '  Set tblTSeg = get_TableClass(m_layers(14)) 'tbltransitsegments
        Dim pTrPoints As IFeatureLayer
        pTrPoints = get_FeatureLayer2(m_layers(23))

        Dim pFC As IFeatureCursor, pFC2 As IFeatureCursor
        Dim pFCj As IFeatureCursor
        Dim pQF As IQueryFilter, pQFt As IQueryFilter, pQF2 As IQueryFilter
        Dim pQFj As IQueryFilter

        Dim pGeometry As IGeometry, pNewGeometry As IGeometry
        Dim pGeometryBag As IGeometryCollection
        pGeometryBag = New GeometryBag
        Dim pTopo As ITopologicalOperator
        Dim pPolyline As IPolyline

        Dim pFeat2 As IFeature
        Dim pNextFeat As IFeature, pNewFeat As IFeature
        Dim pFeat2Atts, pNextFeatAtts, pNewFeatAtts As clsScenarioEdgeAtts
        Dim pJunctFeat As IFeature

        Dim pRow As IRow, pNextRow As IRow
        Dim pTC As ICursor
        Dim pFlds As IFields

        Dim psrcId As String
        Dim i As Long
        Dim fcountI, fcountIJ, mergeCountIJ, mergeCountJI As Long, noMergeCount As Long, attMismatchCount As Long

        Dim match As Boolean
        match = True
        Dim edgesdelete As Boolean
        edgesdelete = False

        Dim newI As Long, newJ As Long

        '  Dim pWorkspaceEdit As IWorkspaceEdit, pWorkspaceEdit2 As IWorkspaceEdit
        '  Set pWorkspaceEdit = g_FWS
        '  pWorkspaceEdit.StartEditing False
        '  pWorkspaceEdit.StartEditOperation

        '  'this one just loops thru junct layer- should be faster
        '  Set pFCj = m_junctShp.Search(Nothing, False)
        '  Set pJunctFeat = pFCj.NextFeature

        Dim lSFld As Integer, dnode As Long, idIndex As Integer
        Dim lIndex As Integer, upIndex As Integer
        Dim idxN As Integer, idxJ As Integer
        Dim idxMode As Integer, idxLT As Integer, idxFT As Integer
        Dim idxJunctID As Integer, idxEdgeID As Integer
        Dim ifld As Integer
        idxJunctID = m_junctShp.FindField("PSRCJunctID")
        '
        '[122305]pan--revised field names for current SDE layer match
        '
        With m_edgeShp
            idxEdgeID = .FindField(g_PSRCEdgeID)
            idIndex = .FindField("PSRC_E2ID")
            upIndex = .FindField("Updated1")
            idxN = .FindField(g_INode)
            idxJ = .FindField(g_JNode)
            idxMode = .FindField(g_Modes)
            idxLT = .FindField(g_LinkType)
            idxFT = .FindField(g_FacType)
        End With

        Dim m As Integer
        Dim same As String, opposite As String
        Dim dir2 As String, dirNext As String
        Dim index2 As Integer, indexNext As Integer
        Dim dir As Boolean
        Dim lType(4) As String
        lType(0) = "GP"
        lType(1) = "HOV"
        lType(2) = "TR"
        lType(3) = "TK"
        Dim timePd(5) As String
        timePd(0) = "AM"
        timePd(1) = "MD"
        timePd(2) = "PM"
        timePd(3) = "EV"
        timePd(4) = "NI"

        mergeCountIJ = 0
        mergeCountJI = 0
        noMergeCount = 0
        attMismatchCount = 0
        pCounter = 1

        '[022406] hyu: get those nodes/edges that can/cannot be thinned into dictionaries.
        Dim dctNoThinNodes As Dictionary(Of Object, Object), dctNoThinEdges As Dictionary(Of Object, Object), dctThinEdges As Dictionary(Of Object, Object)
        'dctNoThinNodes consists of nodes cannot be thinned
        'dctThinEdges consists of edges that are possible be thinned.

        dctNoThinNodes = New Dictionary(Of Object, Object)
        dctThinEdges = New Dictionary(Of Object, Object)
        dctNoThinEdges = New Dictionary(Of Object, Object)
        WriteLogLine(Now() & " start collecting nodes and edges ")

        '  [050306] hyu: transitSegment
        '  getNodes tblTSeg, g_PSRCEdgeID, dctNoThinEdges

        '[041607]hyu: dealing w/ transit points instead of segments table
        'getNodes tblTSeg, g_INode, dctNoThinNodes, "SUBSTRING(CAST(LineID AS CHARACTER), 2, 2) = '" & Right(CStr(inserviceyear), 2) & "' AND SegOrder=1"
        'getNodes tblTSeg, g_JNode, dctNoThinNodes, "SUBSTRING(CAST(LineID AS CHARACTER), 2, 2) = '" & Right(CStr(inserviceyear), 2) & "' AND change>0"
        '[051407] jaf: use SQL with SDE but not with Access
        If strLayerPrefix = "SDE" Then
            getNodes(pTrPoints, g_PSRCJctID, dctNoThinNodes, "SUBSTRING(CAST(LineID AS CHARACTER), 2, 2) = '" & Microsoft.VisualBasic.Right(CStr(inserviceyear), 2) & "'")
            'getPRNodes pFLayerJ.FeatureClass, "Scen_Node", "P_RStalls", dctNoThinNodes
        Else
            getNodes(pTrPoints, g_PSRCJctID, dctNoThinNodes, "MID(Cstr(LineID), 2, 2) = '" & Microsoft.VisualBasic.Right(CStr(inserviceyear), 2) & "'")
        End If

        WriteLogLine(Now() & " finish collecting junctions being transit points: " & dctNoThinNodes.Count)

        getNodes(pTblMode, g_PSRCEdgeID, dctThinEdges, , True)
        WriteLogLine(Now() & " finish collecting edges can be thinned: " & dctThinEdges.Count)

        getNodes(pTolls, g_PSRCEdgeID, dctNoThinEdges)
        WriteLogLine(Now() & " finish collecting edges can NOT be thinned: " & dctNoThinEdges.Count)

        '[060210] SEC- Need to preserve Turn Edges
        getNodes(pTurn, "FrEdgeID", dctNoThinEdges)

        getNodes(pTurn, "ToEdgeID", dctNoThinEdges)

        getNodes(pTurn, g_PSRCJctID, dctNoThinNodes)
        WriteLogLine(Now() & " finish collecting junction can NOT be thinned (including transit): " & dctNoThinNodes.Count)

        '[082409] SEC: Added Park and Rides
        getNodes(m_edgeShp, g_PSRCEdgeID, dctNoThinEdges, g_FacType & "=16 or " & g_FacType & "=17") '[042706] hyu: exclude centroid connector + park and rides from dissolving
        WriteLogLine(Now() & " finish collecting edges can NOT be thinned : " & dctNoThinEdges.Count)

        '[022406] hyu: stats how many edges joining to a junction
        Dim dctJoinEdges As Dictionary(Of Object, Object)
        Dim dctJcts As Dictionary(Of Object, Object)
        Dim dctEdgeID As Dictionary(Of Object, Object)
        Dim lMaxEdgeOID As Long
        dctJoinEdges = New Dictionary(Of Object, Object)
        dctJcts = New Dictionary(Of Object, Object)
        dctEdgeID = New Dictionary(Of Object, Object)

        '  Set dctEdgeID = New Dictionary
        'calcEdgesAtJoint m_edgeShp, g_INode, g_JNode, dctJoinEdges, dctJcts, dctEdgeID
        calcEdgesAtJoint(m_edgeShp, m_junctShp, g_INode, g_JNode, dctJoinEdges, dctJcts, dctEdgeID, lMaxEdgeOID)

        WriteLogLine(Now() & " finish collecting edges at junctions that will be considered dissolving")
        WriteLogLine("junctions count=" & dctJcts.Count)
        WriteLogLine("edges count=" & dctEdgeID.Count)

        Dim pWorkspaceEdit As IWorkspaceEdit, pWorkspaceEdit2 As IWorkspaceEdit
        pWorkspaceEdit = g_FWS
        pWorkspaceEdit.StartEditing(False)
        pWorkspaceEdit.StartEditOperation()

        Dim pFeat As IFeature
        Dim sNode As String, sEdgeID() As String
        '  Do Until pJunctFeat Is Nothing
        Dim sNodes()

        Dim lJctCt As Long
        Dim pSFilt As ISpatialFilter
        Dim lDissolvedEdges As Long, lSelectedEdges As Long, lCreatedEdges As Long

        lSelectedEdges = dctEdgeID.Count
        lDissolvedEdges = 0
        lCreatedEdges = 0

        pSFilt = New SpatialFilter
        sNodes = dctJcts.Keys.ToArray
        pQFj = New QueryFilter
        pQF = New QueryFilter
        pQF2 = New QueryFilter

        For lJctCt = 0 To dctJcts.Count - 1
            If pCounter = 500 Then
                WriteLogLine("pCounter reaches 500")
                '[011806] pan pCounter appears to serve no purpose
                pWorkspaceEdit.StopEditOperation()
                pWorkspaceEdit.StopEditing(True)
                pCounter = 1
                WriteLogLine("saved edits" & "  " & Now)
                pWorkspaceEdit.StartEditing(False)
                pWorkspaceEdit.StartEditOperation()
            End If
            pCounter = pCounter + 1
            '    dnode = pJunctFeat.value(idxJunctID)
            '    sNode = CStr(dnode)
            sNode = sNodes(lJctCt)
            dnode = Val(sNode)
            match = True
            edgesdelete = False
            '    If dnode = 1023 Then MsgBox "here"

            If Not dctNoThinNodes.ContainsKey(sNode) Then
                If dctJoinEdges.ContainsKey(sNode) Then
                    sEdgeID = Split(dctJoinEdges.Item(sNode), ",")

                    If UBound(sEdgeID) - LBound(sEdgeID) <> 1 Then
                        noMergeCount = noMergeCount + 1
                    Else '2 edges connected at this joint
                        If dctThinEdges.ContainsKey(sEdgeID(0)) And dctThinEdges.ContainsKey(sEdgeID(1)) And (Not (dctNoThinEdges.ContainsKey(sEdgeID(0)) Or dctNoThinEdges.ContainsKey(sEdgeID(1)))) Then

                            '[061907] jaf: if the two edge ID's are equal we have a problem...
                            If sEdgeID(0) = sEdgeID(1) Then
                                WriteLogLine("Dissolve found EdgeID=" & sEdgeID(0) & " CONNECTED TO ITSELF")
                            Else
                                '[061907] jaf: the two edges are NOT the same, OK to proceed
                                pFeat2 = dctEdgeID.Item(sEdgeID(0))
                                pNextFeat = dctEdgeID.Item(sEdgeID(1))


                                If Not (pFeat2 Is Nothing Or pNextFeat Is Nothing) Then
                                    pFeat2Atts = New clsScenarioEdgeAtts(pFeat2)
                                    pNextFeatAtts = New clsScenarioEdgeAtts(pNextFeatAtts)
                                    WriteLogLine("potential dnode " & CStr(dnode) & "  " & Now)
                                    'now check if junction really real
                                    '                    Set pFeat2 = m_edgeShp.GetFeature(Val(sEdgeID(0)))
                                    '                    Set pNextFeat = m_edgeShp.GetFeature(Val(sEdgeID(1)))
                                    dir = False
                                    'If pFeat2.Value(idxJ) = dnode Then
                                    If pFeat2Atts.JNode = dnode Then
                                        If pNextFeatAtts.JNode = dnode Then 'then deleting both jnodes
                                            newI = pFeat2Atts.JNode
                                            newJ = pNextFeatAtts.INode
                                        Else 'then deleting jnode of first edge and inode of second
                                            newJ = pNextFeatAtts.JNode
                                            newI = pFeat2Atts.INode
                                            dir = True
                                        End If
                                    Else
                                        newJ = pFeat2Atts.JNode
                                        If pNextFeatAtts.JNode = dnode Then 'then deleting inode of first edge and jnode of second
                                            newI = pNextFeatAtts.INode
                                            dir = True
                                        Else 'then deleting both inodes
                                            newI = pNextFeatAtts.JNode
                                        End If
                                    End If

                                    'compare attributes to detect if we have "real" pseudonode
                                    'pQF2.WhereClause = g_PSRCEdgeID & "=" & pFeat2.value(idxEdgeID) & " OR " & g_PSRCEdgeID & "=" & pNextFeat.value(idxEdgeID)

                                    'Set pQF2 = New QueryFilter
                                    'pQF2.WhereClause = g_PSRCEdgeID & "=" & sEdgeID(0) & " OR " & g_PSRCEdgeID & "=" & sEdgeID(1)
                                    'Set pRow = pTC.NextRow
                                    'Set pNextRow = pTC.NextRow
                                    pQF2.WhereClause = g_PSRCEdgeID & "=" & sEdgeID(0)
                                    'If IsDBNull(pFeat2.Value(pFeat2.Fields.FindField("Updated1"))) Then
                                    If pFeat2Atts.Updated1 = "Null" Then
                                        pTC = pTblMode.Search(pQF2, False)
                                        pRow = pTC.NextRow
                                    Else
                                        pRow = getAttributesRow(pFeat2, Nothing, pTblPrjEdgeAtt, pEvtLine)

                                    End If

                                End If


                                pQF2.WhereClause = g_PSRCEdgeID & "=" & sEdgeID(1)
                                'If IsDBNull(pNextFeat.Value(pFeat2.Fields.FindField("Updated1"))) Then
                                If pNextFeatAtts.Updated1 = "Null" Then
                                    pTC = pTblMode.Search(pQF2, False)
                                    pNextRow = pTC.NextRow
                                Else
                                    Dim pTableName As String


                                    pNextRow = GetAttributesRow2(pNextFeat, Nothing, pTblPrjEdgeAtt, pEvtLine, pTableName)




                                    '                    If dnode = 144794 Then MsgBox "here"
                                    If pRow.Fields.FieldCount <> pNextRow.Fields.FieldCount Then
                                        MsgBox("not equal")
                                    End If


                                    If dir = True Then
                                        For i = 1 To pRow.Fields.FieldCount - 1
                                            '[081208] sec: added fields dateCreated, DateLastUpdated, and LastEditor as fields that should not be compared
                                            If UCase(pRow.Fields.Field(i).Name) <> "PSRCEDGEID" And _
                                                UCase(pRow.Fields.Field(i).Name) <> "IJFFS" And UCase(pRow.Fields.Field(i).Name) <> "JIFFS" And UCase(pRow.Fields.Field(i).Name) <> "DATELASTUPDATED" And UCase(pRow.Fields.Field(i).Name) <> "DATECREATED" And UCase(pRow.Fields.Field(i).Name) <> "LASTEDITOR" Then

                                                'SEC- see if either field is null, consider this still a match for now
                                                If IsDBNull(pRow.Value(i)) Or IsDBNull(pNextRow.Value(i)) Then

                                                ElseIf (pRow.Value(i) <> pNextRow.Value(i)) Then
                                                    match = False
                                                    WriteLogLine(pRow.Fields.Field(i).Name + " diff")
                                                End If
                                                'End If
                                            End If
                                        Next i
                                    Else 'need to compare IJ to JI to match direction
                                        Dim j As Long, l As Long, t As Long
                                        For j = 0 To 1
                                            If j = 0 Then
                                                dir2 = "IJ"
                                                dirNext = "JI"
                                            Else
                                                dir2 = "JI"
                                                dirNext = "IJ"
                                            End If

                                            'begin attribute comparison by directions
                                            index2 = pRow.Fields.FindField(dir2 + "SpeedLimit")
                                            indexNext = pNextRow.Fields.FindField(dirNext + "SpeedLimit")
                                            If (pRow.Value(index2) <> pNextRow.Value(indexNext)) Then
                                                match = False
                                                WriteLogLine("dir speed diff")
                                            End If
                                            'not comparing FFS any more 06/16/05
                                            'index2 = pRow.Fields.FindField(dir2 + "FFS")
                                            ' indexnext = pNextRow.Fields.FindField(dirnext + "FFS")
                                            ' If (pRow.value(index2) <> pNextRow.value(indexnext)) Then
                                            '    match = False
                                            'WriteLogLine "dir ffs diff"
                                            ' End If
                                            index2 = pRow.Fields.FindField(dir2 + "VDFunc")
                                            indexNext = pNextRow.Fields.FindField(dirNext + "VDFunc")
                                            If (pRow.Value(index2) <> pNextRow.Value(indexNext)) Then
                                                match = False
                                                WriteLogLine("dir vdf diff")
                                            End If

                                            index2 = pRow.Fields.FindField(dir2 + "SideWalks")
                                            indexNext = pNextRow.Fields.FindField(dirNext + "SideWalks")
                                            If (pRow.Value(index2) <> pNextRow.Value(indexNext)) Then
                                                match = False
                                                WriteLogLine("dir side diff")
                                            End If
                                            index2 = pRow.Fields.FindField(dir2 + "BikeLanes")
                                            indexNext = pNextRow.Fields.FindField(dirNext + "BikeLanes")
                                            If (pRow.Value(index2) <> pNextRow.Value(indexNext)) Then
                                                match = False
                                                WriteLogLine("dir bike diff")
                                            End If
                                            For l = 0 To 3
                                                If l > 1 Then
                                                    index2 = pRow.Fields.FindField(dir2 + "Lanes" + lType(l))
                                                    indexNext = pNextRow.Fields.FindField(dirNext + "Lanes" + lType(l))
                                                    If (pRow.Value(index2) <> pNextRow.Value(indexNext)) Then
                                                        match = False
                                                        WriteLogLine("dir lanes diff")
                                                    End If
                                                Else
                                                    index2 = pRow.Fields.FindField(dir2 + "LaneCap" + lType(l))
                                                    indexNext = pNextRow.Fields.FindField(dirNext + "LaneCap" + lType(l))
                                                    If (pRow.Value(index2) <> pNextRow.Value(indexNext)) Then
                                                        match = False
                                                        WriteLogLine("dir lanecap diff")
                                                    End If
                                                    For t = 0 To 4
                                                        index2 = pRow.Fields.FindField(dir2 + "Lanes" + lType(l) + timePd(t))
                                                        indexNext = pNextRow.Fields.FindField(dirNext + "Lanes" + lType(l) + timePd(t))

                                                        If (pRow.Value(index2) <> pNextRow.Value(indexNext)) Then
                                                            match = False
                                                            WriteLogLine("lanes diff")
                                                        End If
                                                    Next t
                                                End If
                                            Next l
                                        Next j
                                        pRow = Nothing
                                        pNextRow = Nothing
                                    End If
                                    'SEC- see if either field is null, consider this still a match for now
                                    If IsDBNull(pFeat2.Value(upIndex)) Or IsDBNull(pNextFeat.Value(upIndex)) Then
                                    ElseIf pFeat2.Value(upIndex) <> pNextFeat.Value(upIndex) Then
                                        match = False
                                        WriteLogLine("update diff")
                                    End If

                                    If pFeat2.Value(idxMode) <> pNextFeat.Value(idxMode) Then
                                        match = False
                                        WriteLogLine("modes diff")
                                    End If

                                    If pFeat2.Value(idxLT) <> pNextFeat.Value(idxLT) Then
                                        match = False
                                        WriteLogLine("linktype diff")
                                    End If
                                    i = pFeat2.Fields.FindField("FacilityTy")
                                    If pFeat2.Value(idxFT) <> pNextFeat.Value(idxFT) Then
                                        match = False
                                        WriteLogLine("fac diff")
                                    End If

                                    If (match = False) Then
                                        attMismatchCount = attMismatchCount + 1
                                    Else
                                        WriteLogLine("found match " & Now)

                                        '[073007] hyu: new appraoch to merge edges:
                                        '[073007] hyu: consider the geometry orientation when merging two edges. Replace the old approach of using iTopologicalOperator.ConstructUnion.
                                        pPolyline = MergeEdges(pNextFeat, pFeat2)

                                        '                    Set pGeometryBag = New GeometryBag
                                        '                    Set pGeometry = pNextFeat.ShapeCopy
                                        '                    pGeometryBag.AddGeometry pGeometry
                                        '                    Set pGeometry = pFeat2.ShapeCopy
                                        '                    pGeometryBag.AddGeometry pGeometry
                                        '
                                        'update the merge count for the log
                                        If dir Then
                                            mergeCountIJ = mergeCountIJ + 1
                                        Else
                                            mergeCountJI = mergeCountJI + 1
                                        End If

                                        'merge the edge features
                                        '                    Set pPolyline = New Polyline
                                        '                    Set pTopo = pPolyline
                                        '                    pTopo.ConstructUnion pGeometryBag
                                        pNewFeat = m_edgeShp.CreateFeature
                                        pNewFeat.Shape = pPolyline
                                        '                        pnewfeat.Store
                                        pFlds = pNewFeat.Fields
                                        For i = 1 To pFlds.FieldCount - 1
                                            lIndex = pFeat2.Fields.FindField(pFlds.Field(i).Name)
                                            If (lIndex <> -1) Then
                                                '[032006] hyu: change the if condition
                                                '                                If UCase(pFlds.field(i).name) = "SCEN_LINK" Then
                                                '                                    MsgBox 1
                                                '                                End If
                                                If (pFlds.Field(i).Type <> esriFieldType.esriFieldTypeOID And pFlds.Field(i).Type <> esriFieldType.esriFieldTypeGeometry _
                                                    And UCase(pFlds.Field(i).Name) <> UCase("Shape")) Then
                                                    'If (pFlds.field(i).Type <> esriFieldTypeOID And Not pFlds.field(i).Type = esriFieldTypeGeometry Or pFlds.field(i).name <> "Shape") Then
                                                    'inode and jnode need to be set differently
                                                    If UCase((pFlds.Field(i).Name)) = "INODE" Then
                                                        pNewFeat.Value(i) = newI
                                                    ElseIf UCase((pFlds.Field(i).Name)) = "JNODE" Then
                                                        pNewFeat.Value(i) = newJ
                                                        '[122905] pan--SHAPE.len is internal field not editable
                                                        'ElseIf (pFlds.field(i).name = "SHAPE.len") Then
                                                        '  pnewfeat.value(i) = pPolyline.length
                                                        'Else
                                                        '  pnewfeat.value(i) = pFeat2.value(lindex)
                                                    Else
                                                        If pNewFeat.Fields.Field(i).Editable Then pNewFeat.Value(i) = pFeat2.Value(i)
                                                    End If
                                                End If
                                            End If
                                        Next i

                                        If IsDBNull(pNewFeat.Value(pNewFeat.Fields.FindField("SCENARIOID"))) Then
                                            'MsgBox "Null ScenarioID in new feature",,"NonFatal Error"
                                        End If

                                        'pan Edit session incompatible with Store
                                        'pan 12-16-05 re-inserted store and edge delete
                                        pNewFeat.Value(pNewFeat.Fields.FindField("Dissolve")) = 1
                                        pNewFeat.Store()
                                        lCreatedEdges = lCreatedEdges + 1
                                        edgesdelete = True

                                        updateEdgesAtJoint(dctJoinEdges, pFeat2, pNextFeat, pNewFeat)
                                        '                        dctEdgeID.Item(sEdgeID(0)) = 1
                                        '                        dctEdgeID.Item(sEdgeID(1)) = 1
                                        If pFeat2.OID <= lMaxEdgeOID Then lDissolvedEdges = lDissolvedEdges + 1
                                        If pNextFeat.OID <= lMaxEdgeOID Then lDissolvedEdges = lDissolvedEdges + 1

                                        dctEdgeID.Item(sEdgeID(0)) = pNewFeat
                                        dctEdgeID.Remove(sEdgeID(1))

                                        '   WriteLogLine "delete edges oid=" & pFeat2.OID & ", " & pNextFeat.OID
                                        pFeat2.Delete()
                                        pNextFeat.Delete()

                                        'find junction node and delete it

                                        '[122905] pan--I changed check to gt because want to delete this junction
                                        'If idxJunctID < 0 Then
                                        pJunctFeat = dctJcts.Item(sNode)

                                        '                        pQFj.WhereClause = g_PSRCJctID & "=" & sNode
                                        '                        Set pFCj = m_junctShp.Search(pQFj, False)
                                        '                        Set pJunctFeat = pFCj.NextFeature

                                        If Not pJunctFeat Is Nothing Then
                                            If fVerboseLog Then
                                                If idxJunctID > 0 Then
                                                    WriteLogLine("Deleting intrmed. feature junct OID=" & pJunctFeat.OID)
                                                Else
                                                    WriteLogLine("Deleting intrmed. feature junctID=" & pJunctFeat.Value(idxJunctID))
                                                End If
                                            End If
                                            pJunctFeat.Delete()
                                        End If

                                        'MsgBox pSSet.count

                                    End If '(match = False)
                                End If 'Not (pFeat2 Is Nothing Or pNextFeat Is Nothing)
                            End If 'edge connect to itself: sEdgeID(0) = sEdgeID(1)
                        End If  'match = true:  dctThinEdges.Exists(sEdgeID(0)) And dctThinEdges.Exists(sEdgeID(1)) ...
                    End If 'UBound(sEdgeID) - LBound(sEdgeID) <> 1
                End If 'dctJoinEdges.Exists(sNode)
            End If 'Not dctNoThinNodes.Exists(sNode)

            pFCj = Nothing
            pTC = Nothing
            pFC = Nothing
            pFC2 = Nothing
        Next lJctCt

        pWorkspaceEdit.StopEditOperation()
        pWorkspaceEdit.StopEditing(True)

        _map.ClearSelection()
        Dim PMx As IMxDocument

        _activeView.Refresh()

        pFLayerE = Nothing
        pFLayerJ = Nothing
        pJunctFeat = Nothing
        pFeat2 = Nothing
        pNextFeat = Nothing
        pNewFeat = Nothing
        pQFj = Nothing
        pFCj = Nothing
        pRow = Nothing
        pNextRow = Nothing
        pTblMode = Nothing
        pTolls = Nothing
        pTurn = Nothing
        '   Set tblTSeg = Nothing
        pWorkspaceEdit = Nothing

        dctNoThinEdges.Clear()
        dctThinEdges.Clear()
        dctNoThinNodes.Clear()
        dctJcts.Clear()

        dctNoThinEdges = Nothing
        dctThinEdges = Nothing
        dctNoThinNodes = Nothing
        dctJcts = Nothing

        WriteLogLine("FINISHED dissolveTransEdges2 at " & Now())
        WriteLogLine("Junctions selected= " & dctJoinEdges.Count)
        WriteLogLine("Same dir thinned= " & mergeCountIJ & "; Opp dir thinned= " & mergeCountJI)
        WriteLogLine("Junctions thinned= " & mergeCountIJ + mergeCountJI)
        WriteLogLine("Junctions geometrically eligible to thin but not thinned due to attribute mismatch= " & attMismatchCount)
        WriteLogLine("Junctions geometrically ineligible to thin= " & noMergeCount)
        WriteLogLine("Junctions final= " & m_junctShp.FeatureCount(Nothing))
        '  t = 0
        '  For l = 0 To dctEdgeID.count - 1
        '    If UCase(TypeName(dctEdgeID.Item(l))) = "INTEGER" Then t = t + 1
        '    'If dctEdgeID.Items(l) = 1 Then t = t + 1
        '  Next l

        WriteLogLine("Edges selected= " & lSelectedEdges)
        WriteLogLine("Edges merged= " & lDissolvedEdges)

        '  Set pQF = New QueryFilter
        '  pQF.WhereClause = "Dissolve=1"
        '  WriteLogLine "Edges created= " & m_edgeShp.featurecount(pQF)
        '  WriteLogLine "Edges final= " & m_edgeShp.featurecount(Nothing)

        WriteLogLine("Edges created= " & lCreatedEdges)
        WriteLogLine("Edges final= " & dctEdgeID.Count)
        pQF = Nothing

        WriteLogLine(m_edgeShp.AliasName + " " + CStr(m_edgeShp.FeatureCount(Nothing)))
        'Note: needs to be a round two version of this based on atrributes that don't match but are not significant
        'need to get these attributes!
        Exit Sub

ErrChk:
        CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.dissolveTransEdges2")
        MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.dissolveTransEdges2")
        'Print #1, Err.Description, vbInformation, "dissolveTransEdges2"

        '  If pWorkspaceEdit.IsBeingEdited Then
        pWorkspaceEdit.StopEditOperation()
        pWorkspaceEdit.StopEditing(True)
        '  End If


        pFLayerE = Nothing
        pFLayerJ = Nothing
        pJunctFeat = Nothing
        pFeat2 = Nothing
        pNextFeat = Nothing
        pNewFeat = Nothing
        pQFj = Nothing
        pFCj = Nothing
        pRow = Nothing
        pNextRow = Nothing
        pTblMode = Nothing
        pTolls = Nothing
        pTurn = Nothing
        '   Set tblTSeg = Nothing
        pWorkspaceEdit = Nothing

    End Sub

    Private Sub Dissolve_Keep_TTP3()
        '[022506] hyu: revised function to replace [dissolveTransEdges2_0]

        'Retains junctions and edges used by transit, tolls and turns
        'performs "pseudonode" thinning of the intermediate edge shp file
        'modified to handle toll, turn coincidence with thinning by Melinda 5-25-05
        'retrofitted (minimally) by Jeff with our logging stuff
        'added transit thinning by Melinda 6-3-05 pan

        'on error GoTo ErrChk

        WriteLogLine("")
        WriteLogLine("======================================")
        WriteLogLine("dissolveTransEdges2 started " & Now())
        WriteLogLine("======================================")
        WriteLogLine("")
        If (strLayerPrefix = "SDE") Then
            Dim pwf As IWorkspaceFactory
            pwf = New SdeWorkspaceFactory
        End If

        Dim pCounter As Long
        Dim pFLayerE As IFeatureLayer, pFLayerJ As IFeatureLayer
        pFLayerE = New FeatureLayer
        pFLayerE.FeatureClass = m_edgeShp 'Intermediate Edge
        WriteLogLine(m_edgeShp.AliasName + CStr(m_edgeShp.FeatureCount(Nothing)))
        pFLayerJ = New FeatureLayer
        pFLayerJ.FeatureClass = m_junctShp 'Intermediate Junction

        Dim pTblMode As ITable, pTolls As ITable
        Dim pTurn As IFeatureLayer
        Dim pTblPrjEdgeAtt As ITable, pEvtLine As ITable

        pTblPrjEdgeAtt = get_TableClass(m_layers(24))
        Try


            pTblMode = get_TableClass(m_layers(2)) 'modeAttributes
            pTolls = get_TableClass(m_layers(9)) 'modeTolls
            pTurn = get_FeatureLayer2(m_layers(12)) 'TurnMovements
            pEvtLine = get_TableClass(m_layers(4)) 'use if future event project
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        '[041607]hyu: now dealing with transit points feature layer instead of segment table
        '  Dim tblTSeg As ITable
        '  Set tblTSeg = get_TableClass(m_layers(14)) 'tbltransitsegments
        Dim pTrPoints As IFeatureLayer
        pTrPoints = get_FeatureLayer2(m_layers(23))

        Dim pFC As IFeatureCursor, pFC2 As IFeatureCursor
        Dim pFCj As IFeatureCursor
        Dim pQF As IQueryFilter, pQFt As IQueryFilter, pQF2 As IQueryFilter
        Dim pQFj As IQueryFilter

        Dim pGeometry As IGeometry, pNewGeometry As IGeometry
        Dim pGeometryBag As IGeometryCollection
        pGeometryBag = New GeometryBag
        Dim pTopo As ITopologicalOperator
        Dim pPolyline As IPolyline


        Dim pFeat2Atts, pNextFeatAtts, pNewFeatAtts As clsScenarioEdgeAtts

        Dim pFeat2 As IFeature
        Dim pNextFeat As IFeature, pNewFeat As IFeature
        Dim pJunctFeat As IFeature
        Dim pRow As IRow, pNextRow As IRow
        Dim pTC As ICursor
        Dim pFlds As IFields

        Dim psrcId As String
        Dim i As Long
        Dim fcountI, fcountIJ, mergeCountIJ, mergeCountJI As Long, noMergeCount As Long, attMismatchCount As Long

        Dim match As Boolean
        match = True
        Dim edgesdelete As Boolean
        edgesdelete = False

        Dim newI As Long, newJ As Long

        '  Dim pWorkspaceEdit As IWorkspaceEdit, pWorkspaceEdit2 As IWorkspaceEdit
        '  Set pWorkspaceEdit = g_FWS
        '  pWorkspaceEdit.StartEditing False
        '  pWorkspaceEdit.StartEditOperation

        '  'this one just loops thru junct layer- should be faster
        '  Set pFCj = m_junctShp.Search(Nothing, False)
        '  Set pJunctFeat = pFCj.NextFeature

        Dim lSFld As Integer, dnode As Long, idIndex As Integer
        Dim lIndex As Integer, upIndex As Integer
        Dim idxN As Integer, idxJ As Integer
        Dim idxMode As Integer, idxLT As Integer, idxFT As Integer
        Dim idxJunctID As Integer, idxEdgeID As Integer
        Dim ifld As Integer
        idxJunctID = m_junctShp.FindField("PSRCJunctID")
        '
        '[122305]pan--revised field names for current SDE layer match
        '
        With m_edgeShp
            idxEdgeID = .FindField(g_PSRCEdgeID)
            idIndex = .FindField("PSRC_E2ID")
            upIndex = .FindField("Updated1")
            idxN = .FindField(g_INode)
            idxJ = .FindField(g_JNode)
            idxMode = .FindField(g_Modes)
            idxLT = .FindField(g_LinkType)
            idxFT = .FindField(g_FacType)
        End With

        Dim m As Integer
        Dim same As String, opposite As String
        Dim dir2 As String, dirNext As String
        Dim index2 As Integer, indexNext As Integer
        Dim dir As Boolean
        Dim lType(4) As String
        lType(0) = "GP"
        lType(1) = "HOV"
        lType(2) = "TR"
        lType(3) = "TK"
        Dim timePd(5) As String
        timePd(0) = "AM"
        timePd(1) = "MD"
        timePd(2) = "PM"
        timePd(3) = "EV"
        timePd(4) = "NI"

        mergeCountIJ = 0
        mergeCountJI = 0
        noMergeCount = 0
        attMismatchCount = 0
        pCounter = 1

        '[022406] hyu: get those nodes/edges that can/cannot be thinned into dictionaries.
        Dim dctNoThinNodes As Dictionary(Of Object, Object), dctNoThinEdges As Dictionary(Of Object, Object), dctThinEdges As Dictionary(Of Object, Object)
        'dctNoThinNodes consists of nodes cannot be thinned
        'dctThinEdges consists of edges that are possible be thinned.

        dctNoThinNodes = New Dictionary(Of Object, Object)
        dctThinEdges = New Dictionary(Of Object, Object)
        dctNoThinEdges = New Dictionary(Of Object, Object)
        WriteLogLine(Now() & " start collecting nodes and edges ")

        '  [050306] hyu: transitSegment
        '  getNodes tblTSeg, g_PSRCEdgeID, dctNoThinEdges

        '[041607]hyu: dealing w/ transit points instead of segments table
        'getNodes tblTSeg, g_INode, dctNoThinNodes, "SUBSTRING(CAST(LineID AS CHARACTER), 2, 2) = '" & Right(CStr(inserviceyear), 2) & "' AND SegOrder=1"
        'getNodes tblTSeg, g_JNode, dctNoThinNodes, "SUBSTRING(CAST(LineID AS CHARACTER), 2, 2) = '" & Right(CStr(inserviceyear), 2) & "' AND change>0"
        '[051407] jaf: use SQL with SDE but not with Access
        If strLayerPrefix = "SDE" Then
            getNodes(pTrPoints, g_PSRCJctID, dctNoThinNodes, "SUBSTRING(CAST(LineID AS CHARACTER), 2, 2) = '" & Microsoft.VisualBasic.Right(CStr(inserviceyear), 2) & "'")
            'getPRNodes pFLayerJ.FeatureClass, "Scen_Node", "P_RStalls", dctNoThinNodes
        Else
            getNodes(pTrPoints, g_PSRCJctID, dctNoThinNodes, "MID(Cstr(LineID), 2, 2) = '" & Microsoft.VisualBasic.Right(CStr(inserviceyear), 2) & "'")
        End If

        WriteLogLine(Now() & " finish collecting junctions being transit points: " & dctNoThinNodes.Count)

        getNodes(pTblMode, g_PSRCEdgeID, dctThinEdges, , True)
        WriteLogLine(Now() & " finish collecting edges can be thinned: " & dctThinEdges.Count)

        getNodes(pTolls, g_PSRCEdgeID, dctNoThinEdges)
        WriteLogLine(Now() & " finish collecting edges can NOT be thinned: " & dctNoThinEdges.Count)

        '[060210] SEC- Need to preserve Turn Edges & "Retain" Junctions
        getNodes(pFLayerJ.FeatureClass, "PSRCJUNCTID", dctNoThinNodes, "JunctionType = 10")
        getNodes(pTurn, "FrEdgeID", dctNoThinEdges)

        getNodes(pTurn, "ToEdgeID", dctNoThinEdges)

        getNodes(pTurn, g_PSRCJctID, dctNoThinNodes)

        WriteLogLine(Now() & " finish collecting junction can NOT be thinned (including transit): " & dctNoThinNodes.Count)

        '[082409] SEC: Added Park and Rides
        getNodes(m_edgeShp, g_PSRCEdgeID, dctNoThinEdges, g_FacType & "=16 or " & g_FacType & "=17") '[042706] hyu: exclude centroid connector + park and rides from dissolving
        WriteLogLine(Now() & " finish collecting edges can NOT be thinned : " & dctNoThinEdges.Count)

        '[022406] hyu: stats how many edges joining to a junction
        Dim dctJoinEdges As Dictionary(Of Object, Object)
        Dim dctJcts As Dictionary(Of Object, Object)
        Dim dctEdgeID As Dictionary(Of Object, Object)
        Dim lMaxEdgeOID As Long
        dctJoinEdges = New Dictionary(Of Object, Object)
        dctJcts = New Dictionary(Of Object, Object)
        dctEdgeID = New Dictionary(Of Object, Object)

        '  Set dctEdgeID = New Dictionary(Of Object, Object)
        'calcEdgesAtJoint m_edgeShp, g_INode, g_JNode, dctJoinEdges, dctJcts, dctEdgeID
        calcEdgesAtJoint(m_edgeShp, m_junctShp, g_INode, g_JNode, dctJoinEdges, dctJcts, dctEdgeID, lMaxEdgeOID)

        WriteLogLine(Now() & " finish collecting edges at junctions that will be considered dissolving")
        WriteLogLine("junctions count=" & dctJcts.Count)
        WriteLogLine("edges count=" & dctEdgeID.Count)

        Dim pWorkspaceEdit As IWorkspaceEdit, pWorkspaceEdit2 As IWorkspaceEdit
        pWorkspaceEdit = g_FWS
        pWorkspaceEdit.StartEditing(False)
        pWorkspaceEdit.StartEditOperation()

        Dim pFeat As IFeature
        Dim sNode As String, sEdgeID() As String
        '  Do Until pJunctFeat Is Nothing
        Dim sNodes()
        Dim lJctCt As Long
        Dim pSFilt As ISpatialFilter
        Dim lDissolvedEdges As Long, lSelectedEdges As Long, lCreatedEdges As Long

        lSelectedEdges = dctEdgeID.Count
        lDissolvedEdges = 0
        lCreatedEdges = 0

        pSFilt = New SpatialFilter
        sNodes = dctJcts.Keys.ToArray
        pQFj = New QueryFilter
        pQF = New QueryFilter
        pQF2 = New QueryFilter

        For lJctCt = 0 To dctJcts.Count - 1
            If pCounter = 500 Then
                WriteLogLine("pCounter reaches 500")
                '[011806] pan pCounter appears to serve no purpose
                pWorkspaceEdit.StopEditOperation()
                pWorkspaceEdit.StopEditing(True)
                pCounter = 1
                WriteLogLine("saved edits" & "  " & Now)
                pWorkspaceEdit.StartEditing(False)
                pWorkspaceEdit.StartEditOperation()
            End If
            pCounter = pCounter + 1
            '    dnode = pJunctFeat.value(idxJunctID)
            '    sNode = CStr(dnode)
            sNode = sNodes(lJctCt)
            dnode = Val(sNode)
            match = True
            edgesdelete = False
            '    If dnode = 1023 Then MsgBox "here"

            If Not dctNoThinNodes.ContainsKey(sNode) Then
                If dctJoinEdges.ContainsKey(sNode) Then
                    sEdgeID = Split(dctJoinEdges.Item(sNode), ",")

                    If UBound(sEdgeID) - LBound(sEdgeID) <> 1 Then
                        noMergeCount = noMergeCount + 1
                    Else '2 edges connected at this joint
                        If dctThinEdges.ContainsKey(sEdgeID(0)) And dctThinEdges.ContainsKey(sEdgeID(1)) And (Not (dctNoThinEdges.ContainsKey(sEdgeID(0)) Or dctNoThinEdges.ContainsKey(sEdgeID(1)))) Then

                            '[061907] jaf: if the two edge ID's are equal we have a problem...
                            If sEdgeID(0) = sEdgeID(1) Then
                                WriteLogLine("Dissolve found EdgeID=" & sEdgeID(0) & " CONNECTED TO ITSELF")
                            Else
                                '[061907] jaf: the two edges are NOT the same, OK to proceed

                                pFeat2 = dctEdgeID.Item(sEdgeID(0))
                                pNextFeat = dctEdgeID.Item(sEdgeID(1))

                                If Not (pFeat2 Is Nothing Or pNextFeat Is Nothing) Then
                                    pFeat2Atts = New clsScenarioEdgeAtts(pFeat2)
                                    pNextFeatAtts = New clsScenarioEdgeAtts(pNextFeatAtts)
                                    WriteLogLine("potential dnode " & CStr(dnode) & "  " & Now)
                                    'now check if junction really real
                                    '                    Set pFeat2 = m_edgeShp.GetFeature(Val(sEdgeID(0)))
                                    '                    Set pNextFeat = m_edgeShp.GetFeature(Val(sEdgeID(1)))
                                    dir = False
                                    'If pFeat2.Value(idxJ) = dnode Then
                                    If pFeat2Atts.JNode = dnode Then
                                        If pNextFeatAtts.JNode = dnode Then 'then deleting both jnodes
                                            newI = pFeat2Atts.INode
                                            newJ = pNextFeatAtts.INode
                                        Else 'then deleting jnode of first edge and inode of second
                                            newJ = pNextFeatAtts.JNode
                                            newI = pFeat2Atts.INode
                                            dir = True
                                        End If
                                    Else
                                        newJ = pFeat2Atts.JNode
                                        If pNextFeatAtts.JNode = dnode Then 'then deleting inode of first edge and jnode of second
                                            newI = pNextFeatAtts.INode
                                            dir = True
                                        Else 'then deleting both inodes
                                            newI = pNextFeatAtts.JNode
                                        End If
                                    End If

                                    'compare attributes to detect if we have "real" pseudonode
                                    'pQF2.WhereClause = g_PSRCEdgeID & "=" & pFeat2.value(idxEdgeID) & " OR " & g_PSRCEdgeID & "=" & pNextFeat.value(idxEdgeID)

                                    'Set pQF2 = New QueryFilter
                                    'pQF2.WhereClause = g_PSRCEdgeID & "=" & sEdgeID(0) & " OR " & g_PSRCEdgeID & "=" & sEdgeID(1)
                                    'Set pRow = pTC.NextRow
                                    'Set pNextRow = pTC.NextRow
                                    pQF2.WhereClause = g_PSRCEdgeID & "=" & sEdgeID(0)
                                    If pFeat2Atts.Updated1 = "Null" Then
                                        pTC = pTblMode.Search(pQF2, False)
                                        pRow = pTC.NextRow
                                    Else
                                        pRow = getAttributesRow(pFeat2, Nothing, pTblPrjEdgeAtt, pEvtLine)

                                    End If

                                    'End If


                                    pQF2.WhereClause = g_PSRCEdgeID & "=" & sEdgeID(1)
                                    If IsDBNull(pNextFeat.Value(pFeat2.Fields.FindField("Updated1"))) Then
                                        pTC = pTblMode.Search(pQF2, False)
                                        pNextRow = pTC.NextRow
                                    Else


                                        pNextRow = getAttributesRow(pNextFeat, Nothing, pTblPrjEdgeAtt, pEvtLine)


                                    End If


                                    '                    If dnode = 144794 Then MsgBox "here"
                                    If pRow.Fields.FieldCount <> pNextRow.Fields.FieldCount Then
                                        MsgBox("not equal")
                                    End If


                                    If dir = True Then
                                        For i = 1 To pRow.Fields.FieldCount - 1
                                            '[081208] sec: added fields dateCreated, DateLastUpdated, and LastEditor as fields that should not be compared
                                            If UCase(pRow.Fields.Field(i).Name) <> "PSRCEDGEID" And _
                                                UCase(pRow.Fields.Field(i).Name) <> "IJFFS" And UCase(pRow.Fields.Field(i).Name) <> "JIFFS" And UCase(pRow.Fields.Field(i).Name) <> "DATELASTUPDATED" And UCase(pRow.Fields.Field(i).Name) <> "DATECREATED" And UCase(pRow.Fields.Field(i).Name) <> "LASTEDITOR" Then

                                                'SEC- see if either field is null, consider this still a match for now
                                                If IsDBNull(pRow.Value(i)) Or IsDBNull(pNextRow.Value(i)) Then

                                                ElseIf (pRow.Value(i) <> pNextRow.Value(i)) Then
                                                    match = False
                                                    WriteLogLine(pRow.Fields.Field(i).Name + " diff")
                                                End If
                                                'End If
                                            End If
                                        Next i
                                    Else 'need to compare IJ to JI to match direction
                                        Dim j As Long, l As Long, t As Long
                                        For j = 0 To 1
                                            If j = 0 Then
                                                dir2 = "IJ"
                                                dirNext = "JI"
                                            Else
                                                dir2 = "JI"
                                                dirNext = "IJ"
                                            End If

                                            'begin attribute comparison by directions
                                            index2 = pRow.Fields.FindField(dir2 + "SpeedLimit")
                                            indexNext = pNextRow.Fields.FindField(dirNext + "SpeedLimit")
                                            If (pRow.Value(index2) <> pNextRow.Value(indexNext)) Then
                                                match = False
                                                WriteLogLine("dir speed diff")
                                            End If
                                            'not comparing FFS any more 06/16/05
                                            'index2 = pRow.Fields.FindField(dir2 + "FFS")
                                            ' indexnext = pNextRow.Fields.FindField(dirnext + "FFS")
                                            ' If (pRow.value(index2) <> pNextRow.value(indexnext)) Then
                                            '    match = False
                                            'WriteLogLine "dir ffs diff"
                                            ' End If
                                            index2 = pRow.Fields.FindField(dir2 + "VDFunc")
                                            indexNext = pNextRow.Fields.FindField(dirNext + "VDFunc")
                                            If (pRow.Value(index2) <> pNextRow.Value(indexNext)) Then
                                                match = False
                                                WriteLogLine("dir vdf diff")
                                            End If

                                            index2 = pRow.Fields.FindField(dir2 + "SideWalks")
                                            indexNext = pNextRow.Fields.FindField(dirNext + "SideWalks")
                                            If (pRow.Value(index2) <> pNextRow.Value(indexNext)) Then
                                                match = False
                                                WriteLogLine("dir side diff")
                                            End If
                                            index2 = pRow.Fields.FindField(dir2 + "BikeLanes")
                                            indexNext = pNextRow.Fields.FindField(dirNext + "BikeLanes")
                                            If (pRow.Value(index2) <> pNextRow.Value(indexNext)) Then
                                                match = False
                                                WriteLogLine("dir bike diff")
                                            End If
                                            For l = 0 To 3
                                                If l > 1 Then
                                                    index2 = pRow.Fields.FindField(dir2 + "Lanes" + lType(l))
                                                    indexNext = pNextRow.Fields.FindField(dirNext + "Lanes" + lType(l))
                                                    If (pRow.Value(index2) <> pNextRow.Value(indexNext)) Then
                                                        match = False
                                                        WriteLogLine("dir lanes diff")
                                                    End If
                                                Else
                                                    index2 = pRow.Fields.FindField(dir2 + "LaneCap" + lType(l))
                                                    indexNext = pNextRow.Fields.FindField(dirNext + "LaneCap" + lType(l))
                                                    If (pRow.Value(index2) <> pNextRow.Value(indexNext)) Then
                                                        match = False
                                                        WriteLogLine("dir lanecap diff")
                                                    End If
                                                    For t = 0 To 4
                                                        index2 = pRow.Fields.FindField(dir2 + "Lanes" + lType(l) + timePd(t))
                                                        indexNext = pNextRow.Fields.FindField(dirNext + "Lanes" + lType(l) + timePd(t))

                                                        If (pRow.Value(index2) <> pNextRow.Value(indexNext)) Then
                                                            match = False
                                                            WriteLogLine("lanes diff")
                                                        End If
                                                    Next t
                                                End If
                                            Next l
                                        Next j
                                        pRow = Nothing
                                        pNextRow = Nothing
                                    End If
                                    'SEC- see if either field is null, consider this still a match for now
                                    'If IsDBNull(pFeat2.Value(upIndex)) Or IsDBNull(pNextFeat.Value(upIndex)) Then
                                    If pFeat2Atts.Updated1 <> pNextFeatAtts.Updated1 Then
                                        match = False
                                        WriteLogLine("update diff")
                                    End If

                                    If pFeat2Atts.Modes <> pNextFeatAtts.Modes Then
                                        match = False
                                        WriteLogLine("modes diff")
                                    End If

                                    If pFeat2Atts.LinkType <> pNextFeatAtts.LinkType Then
                                        match = False
                                        WriteLogLine("linktype diff")
                                    End If
                                    i = pFeat2.Fields.FindField("FacilityTy")
                                    If pFeat2Atts.FacilityType <> pNextFeatAtts.FacilityType Then
                                        match = False
                                        WriteLogLine("fac diff")
                                    End If

                                    If (match = False) Then
                                        attMismatchCount = attMismatchCount + 1
                                    Else
                                        WriteLogLine("found match " & Now)

                                        '[073007] hyu: new appraoch to merge edges:
                                        '[073007] hyu: consider the geometry orientation when merging two edges. Replace the old approach of using iTopologicalOperator.ConstructUnion.
                                        pPolyline = MergeEdges(pNextFeat, pFeat2)

                                        '                    Set pGeometryBag = New GeometryBag
                                        '                    Set pGeometry = pNextFeat.ShapeCopy
                                        '                    pGeometryBag.AddGeometry pGeometry
                                        '                    Set pGeometry = pFeat2.ShapeCopy
                                        '                    pGeometryBag.AddGeometry pGeometry
                                        '
                                        'update the merge count for the log
                                        If dir Then
                                            mergeCountIJ = mergeCountIJ + 1
                                        Else
                                            mergeCountJI = mergeCountJI + 1
                                        End If

                                        'merge the edge features
                                        '                    Set pPolyline = New Polyline
                                        '                    Set pTopo = pPolyline
                                        '                    pTopo.ConstructUnion pGeometryBag
                                        pNewFeat = m_edgeShp.CreateFeature
                                        pNewFeat.Shape = pPolyline
                                        '                        pnewfeat.Store
                                        pFlds = pNewFeat.Fields
                                        For i = 1 To pFlds.FieldCount - 1
                                            lIndex = pFeat2.Fields.FindField(pFlds.Field(i).Name)
                                            If (lIndex <> -1) Then
                                                '[032006] hyu: change the if condition
                                                '                                If UCase(pFlds.field(i).name) = "SCEN_LINK" Then
                                                '                                    MsgBox 1
                                                '                                End If
                                                If (pFlds.Field(i).Type <> esriFieldType.esriFieldTypeOID And pFlds.Field(i).Type <> esriFieldType.esriFieldTypeGeometry _
                                                    And UCase(pFlds.Field(i).Name) <> UCase("Shape")) Then
                                                    'If (pFlds.field(i).Type <> esriFieldTypeOID And Not pFlds.field(i).Type = esriFieldTypeGeometry Or pFlds.field(i).name <> "Shape") Then
                                                    'inode and jnode need to be set differently
                                                    If UCase((pFlds.Field(i).Name)) = "INODE" Then
                                                        pNewFeat.Value(i) = newI
                                                    ElseIf UCase((pFlds.Field(i).Name)) = "JNODE" Then
                                                        pNewFeat.Value(i) = newJ
                                                        '[122905] pan--SHAPE.len is internal field not editable
                                                        'ElseIf (pFlds.field(i).name = "SHAPE.len") Then
                                                        '  pnewfeat.value(i) = pPolyline.length
                                                        'Else
                                                        '  pnewfeat.value(i) = pFeat2.value(lindex)
                                                    Else
                                                        If pNewFeat.Fields.Field(i).Editable Then pNewFeat.Value(i) = pFeat2.Value(i)
                                                    End If
                                                End If
                                            End If
                                        Next i

                                        If IsDBNull(pNewFeat.Value(pNewFeat.Fields.FindField("SCENARIOID"))) Then
                                            'MsgBox "Null ScenarioID in new feature",,"NonFatal Error"
                                        End If

                                        'pan Edit session incompatible with Store
                                        'pan 12-16-05 re-inserted store and edge delete
                                        pNewFeat.Value(pNewFeat.Fields.FindField("Dissolve")) = 1
                                        pNewFeat.Store()
                                        lCreatedEdges = lCreatedEdges + 1
                                        edgesdelete = True

                                        updateEdgesAtJoint(dctJoinEdges, pFeat2, pNextFeat, pNewFeat)
                                        '                        dctEdgeID.Item(sEdgeID(0)) = 1
                                        '                        dctEdgeID.Item(sEdgeID(1)) = 1
                                        If pFeat2.OID <= lMaxEdgeOID Then lDissolvedEdges = lDissolvedEdges + 1
                                        If pNextFeat.OID <= lMaxEdgeOID Then lDissolvedEdges = lDissolvedEdges + 1

                                        dctEdgeID.Item(sEdgeID(0)) = pNewFeat
                                        dctEdgeID.Remove(sEdgeID(1))

                                        '   WriteLogLine "delete edges oid=" & pFeat2.OID & ", " & pNextFeat.OID
                                        pFeat2.Delete()
                                        pNextFeat.Delete()

                                        'find junction node and delete it

                                        '[122905] pan--I changed check to gt because want to delete this junction
                                        'If idxJunctID < 0 Then
                                        pJunctFeat = dctJcts.Item(sNode)

                                        '                        pQFj.WhereClause = g_PSRCJctID & "=" & sNode
                                        '                        Set pFCj = m_junctShp.Search(pQFj, False)
                                        '                        Set pJunctFeat = pFCj.NextFeature

                                        If Not pJunctFeat Is Nothing Then
                                            If fVerboseLog Then
                                                If idxJunctID > 0 Then
                                                    WriteLogLine("Deleting intrmed. feature junct OID=" & pJunctFeat.OID)
                                                Else
                                                    WriteLogLine("Deleting intrmed. feature junctID=" & pJunctFeat.Value(idxJunctID))
                                                End If
                                            End If
                                            pJunctFeat.Delete()
                                        End If

                                        'MsgBox pSSet.count

                                    End If '(match = False)
                                End If 'Not (pFeat2 Is Nothing Or pNextFeat Is Nothing)
                            End If 'edge connect to itself: sEdgeID(0) = sEdgeID(1)
                        End If  'match = true:  dctThinEdges.Exists(sEdgeID(0)) And dctThinEdges.Exists(sEdgeID(1)) ...
                    End If 'UBound(sEdgeID) - LBound(sEdgeID) <> 1
                End If 'dctJoinEdges.Exists(sNode)
            End If 'Not dctNoThinNodes.Exists(sNode)

            pFCj = Nothing
            pTC = Nothing
            pFC = Nothing
            pFC2 = Nothing
        Next lJctCt

        pWorkspaceEdit.StopEditOperation()
        pWorkspaceEdit.StopEditing(True)

        _map.ClearSelection()
        Dim PMx As IMxDocument

        _activeView.Refresh()

        pFLayerE = Nothing
        pFLayerJ = Nothing
        pJunctFeat = Nothing
        pFeat2 = Nothing
        pNextFeat = Nothing
        pNewFeat = Nothing
        pQFj = Nothing
        pFCj = Nothing
        pRow = Nothing
        pNextRow = Nothing
        pTblMode = Nothing
        pTolls = Nothing
        pTurn = Nothing
        '   Set tblTSeg = Nothing
        pWorkspaceEdit = Nothing

        dctNoThinEdges.Clear()
        dctThinEdges.Clear()
        dctNoThinNodes.Clear()
        dctJcts.Clear()

        dctNoThinEdges = Nothing
        dctThinEdges = Nothing
        dctNoThinNodes = Nothing
        dctJcts = Nothing

        WriteLogLine("FINISHED dissolveTransEdges2 at " & Now())
        WriteLogLine("Junctions selected= " & dctJoinEdges.Count)
        WriteLogLine("Same dir thinned= " & mergeCountIJ & "; Opp dir thinned= " & mergeCountJI)
        WriteLogLine("Junctions thinned= " & mergeCountIJ + mergeCountJI)
        WriteLogLine("Junctions geometrically eligible to thin but not thinned due to attribute mismatch= " & attMismatchCount)
        WriteLogLine("Junctions geometrically ineligible to thin= " & noMergeCount)
        WriteLogLine("Junctions final= " & m_junctShp.FeatureCount(Nothing))
        '  t = 0
        '  For l = 0 To dctEdgeID.count - 1
        '    If UCase(TypeName(dctEdgeID.Item(l))) = "INTEGER" Then t = t + 1
        '    'If dctEdgeID.Items(l) = 1 Then t = t + 1
        '  Next l

        WriteLogLine("Edges selected= " & lSelectedEdges)
        WriteLogLine("Edges merged= " & lDissolvedEdges)

        '  Set pQF = New QueryFilter
        '  pQF.WhereClause = "Dissolve=1"
        '  WriteLogLine "Edges created= " & m_edgeShp.featurecount(pQF)
        '  WriteLogLine "Edges final= " & m_edgeShp.featurecount(Nothing)

        WriteLogLine("Edges created= " & lCreatedEdges)
        WriteLogLine("Edges final= " & dctEdgeID.Count)
        pQF = Nothing

        WriteLogLine(m_edgeShp.AliasName + " " + CStr(m_edgeShp.FeatureCount(Nothing)))
        'Note: needs to be a round two version of this based on atrributes that don't match but are not significant
        'need to get these attributes!
        Exit Sub

ErrChk:
        CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.dissolveTransEdges2")
        MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.dissolveTransEdges2")
        'Print #1, Err.Description, vbInformation, "dissolveTransEdges2"

        '  If pWorkspaceEdit.IsBeingEdited Then
        pWorkspaceEdit.StopEditOperation()
        pWorkspaceEdit.StopEditing(True)
        '  End If


        pFLayerE = Nothing
        pFLayerJ = Nothing
        pJunctFeat = Nothing
        pFeat2 = Nothing
        pNextFeat = Nothing
        pNewFeat = Nothing
        pQFj = Nothing
        pFCj = Nothing
        pRow = Nothing
        pNextRow = Nothing
        pTblMode = Nothing
        pTolls = Nothing
        pTurn = Nothing
        '   Set tblTSeg = Nothing
        pWorkspaceEdit = Nothing

    End Sub

    Public Function get_FeatureLayer2(ByVal featClsname As String) As IFeatureLayer
        'returns the feature class stated in featClsname
        'on error GoTo eh




        Dim pEnumLy As IEnumLayer


        'm_Map = pActiveView.FocusMap
        pEnumLy = _map.Layers

        Dim pLy As ILayer, pFLy As IFeatureLayer
        pLy = pEnumLy.Next
        Do Until pLy Is Nothing

            If pLy.Valid Then
                If TypeOf pLy Is IFeatureLayer Then
                    pFLy = pLy

                    'If UCase(pFLy.FeatureClass.AliasName) = UCase(featClsname) Then
                    If UCase(pFLy.Name) = UCase(featClsname) Then
                        get_FeatureLayer2 = pFLy
                        Exit Do

                    End If
                End If
            End If
            pLy = pEnumLy.Next
        Loop

        If Not get_FeatureLayer2 Is Nothing Then Exit Function

        Dim pFeatLayer As IFeatureLayer
        pFeatLayer = New FeatureLayer
        Dim pfeatcls As IFeatureClass
        pfeatcls = getFeatureClass(featClsname)
        pFeatLayer.FeatureClass = pfeatcls

        get_FeatureLayer2 = pFeatLayer
        If pFeatLayer Is Nothing Then
            MsgBox("Error: did not find the feature class " + featClsname)
        End If

        Exit Function

eh:
        MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.get_FeatureLayer")

    End Function
    Public Function GetAttributesRow2(ByVal pFeat As IFeature, ByVal pMRow As IRow, ByVal pTblLine As ITable, ByVal pEvtLine As ITable, ByVal TableName As String) As IRow
        Dim pRow As IRow
        Dim lSFld As Long, shpindex As Long, upIndex As Long
        Dim tempIndex As Long
        Dim tstring As String
        Dim pQFt As IQueryFilter2
        Dim pTC As ICursor

        upIndex = m_edgeShp.FindField("Updated1")
        lSFld = m_edgeShp.FindField(g_PSRCEdgeID) 'this used to get out of tblmode
        shpindex = m_edgeShp.FindField("shptype")

        '[062007] jaf:  Updated1 field of ScenarioEdge is "Yes" is future year project replaces base year ModeAttributes record
        If (pFeat.Value(upIndex) = "Yes") Then
            If (pFeat.Value(shpindex) <> "Event") Then
                'get future attributes from tblScenarioProjects
                '        tempIndex = pFeat.Fields.FindField("prjRte")
                '        tstring = pFeat.value(tempIndex)
                TableName = "tblScenarioProjects"
                pQFt = New QueryFilter
                'pQFt.WhereClause = "projRteID = " + tstring
                pQFt.WhereClause = g_PSRCEdgeID & "=" & pFeat.Value(lSFld)
                pTC = pTblLine.Search(pQFt, False)
                pRow = pTC.NextRow
                'If (pTblLine.rowcount(pQFt) = 0) Then
                If pRow Is Nothing Then
                    'jaf--modified to give more info
                    'WriteLogLine "DATA ERROR: project " & CStr(pFeat.value(lSFld)) & " is missing in tblLine--GlobalMod.create_NetFile"
                    WriteLogLine("DATA ERROR: project edge " & CStr(pFeat.Value(lSFld)) & " is missing in tblScenarioProject--GlobalMod.create_NetFile")
                    pRow = pMRow
                    '        Else
                    '            Set pRow = pTC.NextRow
                End If
            Else
                'get future attributes from evtLineProjects
                TableName = "evtLineProjects"
                tempIndex = pFeat.Fields.FindField("prjRte")
                tstring = pFeat.Value(tempIndex)
                pQFt = New QueryFilter
                pQFt.WhereClause = "projRteID = " + tstring
                pTC = pEvtLine.Search(pQFt, False)
                pRow = pTC.NextRow
                'If (pEvtLine.rowcount(pQFt) = 0) Then
                If pRow Is Nothing Then
                    'jaf--modified to give more info
                    WriteLogLine("DATA ERROR: project " & CStr(pFeat.Value(lSFld)) & " is missing in evtLine--GlobalMod.create_NetFile")
                    pRow = pMRow
                    '        Else
                    '            Set pRow = pTC.NextRow
                End If
            End If
        Else
            'no future project applied to this edge in this scenario: use ModeAttributes
            pRow = pMRow
            TableName = "ModeAttributes"
        End If 'now set which table to pull the attributes from

        GetAttributesRow2 = pRow


    End Function

End Class
