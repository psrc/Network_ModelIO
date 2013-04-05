
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

Imports ESRI.ArcGIS.GeoDatabaseUI
Imports Scripting


Module hyu_Dissolve
    Sub SelectByLocation(ByVal pFLayer1 As IFeatureLayer, ByVal pFLayer2 As IFeatureLayer, ByVal pFSel1 As IFeatureSelection) 'transrefjunction, transrefedge
        '[022406] hyu: replace old [SelectByLocation0] function to spead up performance
        ' selects all features in layer(0) that intersect any
        ' selected features in layer(1)
        Debug.Print("Select by location start: " & Now())

        pFSel1.Clear()

        Dim pQByLy As IQueryByLayer
        pQByLy = New QueryByLayer
        With pQByLy
            .FromLayer = pFLayer1
            .ByLayer = pFLayer2
            .LayerSelectionMethod = esriLayerSelectionMethod.esriLayerSelectIntersect
            .ResultType = esriSelectionResultEnum.esriSelectionResultAdd
            .UseSelectedFeatures = True
            pFSel1.SelectionSet = .Select
        End With

        pQByLy = Nothing

        Debug.Print("Select by location end: " & Now())

    End Sub

    Public Sub SelectByAtt0(ByVal pEdgeFcls As IFeatureClass, ByVal pJunctLy As IFeatureLayer)

        Dim pFCS As ICursor, lFld As Long
        lFld = pJunctLy.FeatureClass.FindField(g_PSRCJctID)

        Dim pMemRelClsFac As IMemoryRelationshipClassFactory
        Dim pRelCls As IRelationshipClass
        pMemRelClsFac = New MemoryRelationshipClassFactory
        pRelCls = pMemRelClsFac.Open("test", pEdgeFcls, g_INode, pJunctLy.FeatureClass, g_PSRCJctID, "forward", "backward", esriRelCardinality.esriRelCardinalityOneToOne)

        Dim pFilt As IQueryFilter
        pFilt = New QueryFilter
        '    pFilt.WhereClause = "Not """ & pEdgeFCls.AliasName & "." & pEdgeFCls.OIDFieldName & """ is null"
        '    pQFilter.WhereClause = """STATE_NAME"" = 'California'"
        pFilt.WhereClause = "NOT " & pEdgeFcls.AliasName & "." & g_INode & " IS NULL"

        Dim pFSel As IFeatureSelection
        Dim pDispRel As IDisplayRelationshipClass
        pDispRel = pJunctLy
        pDispRel.DisplayRelationshipClass(pRelCls, esriJoinType.esriLeftOuterJoin)

        pFSel = pJunctLy
        pFSel.SelectFeatures(pFilt, esriSelectionResultEnum.esriSelectionResultNew, False)
        Dim p As IQueryDef
        'remove join
        pDispRel.DisplayRelationshipClass(Nothing, esriJoinType.esriLeftInnerJoin)

        pRelCls = pMemRelClsFac.Open("test", pEdgeFcls, g_JNode, pJunctLy.FeatureClass, g_PSRCJctID, "forward", "backward", esriRelCardinality.esriRelCardinalityOneToOne)
        '    pFilt.WhereClause = "Not """ & pEdgeFCls.AliasName & "." & pEdgeFCls.OIDFieldName & """ is null"
        pFilt.WhereClause = "NOT " & pEdgeFcls.AliasName & "." & g_JNode & " IS NULL"
        '    pFilt.WhereClause = pEdgeFCls.AliasName & "." & g_JNode & ">=0"
        pDispRel.DisplayRelationshipClass(pRelCls, esriJoinType.esriLeftOuterJoin)
        pFSel.SelectFeatures(pFilt, esriSelectionResultEnum.esriSelectionResultAdd, False)
        pDispRel.DisplayRelationshipClass(Nothing, esriJoinType.esriLeftInnerJoin)

        pDispRel = Nothing
        pFSel = Nothing
        pFilt = Nothing
        pRelCls = Nothing
        pMemRelClsFac = Nothing

    End Sub

    'Public Sub selectByAtt(pEdgeFcls As IFeatureClass, pJunctLy As IFeatureLayer)
    Public Sub selectJunctByAtt(ByVal pEdgeFcls As IFeatureClass, ByVal pJunctLy As IFeatureLayer, ByVal pToJctFCls As IFeatureClass)
        Dim dctJcts As Dictionary
        dctJcts = New Dictionary
        getNodes(pEdgeFcls, g_INode, dctJcts)
        getNodes(pEdgeFcls, g_JNode, dctJcts)

        Dim pFCls As IFeatureClass, pFCls2 As IFeatureClass
        Dim pFCS As IFeatureCursor
        Dim pFt As IFeature
        Dim pIns As IFeatureCursor
        Dim pFBuf As IFeatureBuffer
        Dim fldJctID As Long
        Dim l As Long
        Dim fld2 As Long

        pFCls2 = pToJctFCls
        pIns = pFCls2.Insert(True)
        pFCls = pJunctLy.FeatureClass
        fldJctID = pFCls.FindField(g_PSRCJctID)
        Debug.Print("Selecting junctions start: " & Now())
        pFCS = pFCls.Search(Nothing, False)
        pFt = pFCS.NextFeature
        Do Until pFt Is Nothing
            If Not IsDBNull(pFt.Value(fldJctID)) Then
                If dctJcts.Exists(CStr(pFt.Value(fldJctID))) Then
                    pFBuf = pFCls2.CreateFeatureBuffer
                    With pFBuf
                        .Shape = pFt.ShapeCopy
                        For l = 0 To pFt.Fields.FieldCount - 1

                            If (Not IsDBNull(pFt.Value(l))) And (pFt.Fields.Field(l).Type <> esriFieldType.esriFieldTypeGeometry) And (pFt.Fields.Field(l).Type <> esriFieldType.esriFieldTypeOID) Then
                                fld2 = pFCls2.FindField(pFt.Fields.Field(l).Name)
                                If pFCls2.Fields.Field(fld2).Editable Then pFBuf.Value(fld2) = pFt.Value(l)
                            End If
                        Next l
                    End With
                    pIns.InsertFeature(pFBuf)
                End If
            End If
            pFt = pFCS.NextFeature
        Loop
        Debug.Print("Selecting junctions end: " & Now())
        pFBuf = Nothing
        pIns = Nothing
        pFCls = Nothing
        pFCls2 = Nothing
        pFCS = Nothing
        dctJcts.RemoveAll()
        dctJcts = Nothing
        '    Dim pFilt As IQueryFilter
        '    Set pFilt = New QueryFilter
        '
        '    Dim pFSel As IFeatureSelection
        '    Set pFSel = pJunctLy
        '
        '    Dim i As Integer, l As Long
        '    Dim sSQL As String
        '    Dim vJcts() As String
        '    'vJcts = dctJcts.Keys
        '    Dim vjct()
        '    vjct = dctJcts.Keys
        'Debug.Print "clear selecting: " & Now()
        '    pFSel.Clear
        'Debug.Print "start selecting: " & Now()
        '
        '    For l = 0 To dctJcts.count - 1
        '
        '        If i = 0 Then
        '            sSQL = g_PSRCJctID & "=" & CStr(vjct(l))
        '
        '        Else
        '            sSQL = sSQL & " OR " & g_PSRCJctID & "=" & CStr(vjct(l))
        '        End If
        '        i = i + 1
        '
        '        If i = 100 Or l = dctJcts.count - 1 Then
        '            pFilt.WhereClause = sSQL
        '            pFSel.SelectFeatures pFilt, esriSelectionResultAdd, False
        '            i = 0
        '            sSQL = ""
        '        End If
        '    Next l
        'Debug.Print "end selecting: " & Now()
        '    Set dctJcts = Nothing
        '    Set pFSel = Nothing
        '    Set pFilt = Nothing

    End Sub

    Public Sub dissolveTransEdges2(ByVal pApp As IApplication)
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
            pTurn = get_FeatureLayer(m_layers(12)) 'TurnMovements
            pEvtLine = get_TableClass(m_layers(4)) 'use if future event project
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        '[041607]hyu: now dealing with transit points feature layer instead of segment table
        '  Dim tblTSeg As ITable
        '  Set tblTSeg = get_TableClass(m_layers(14)) 'tbltransitsegments
        Dim pTrPoints As IFeatureLayer
        pTrPoints = get_FeatureLayer(m_layers(23))

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
        Dim dctNoThinNodes As Dictionary, dctNoThinEdges As Dictionary, dctThinEdges As Dictionary
        'dctNoThinNodes consists of nodes cannot be thinned
        'dctThinEdges consists of edges that are possible be thinned.

        dctNoThinNodes = New Dictionary
        dctThinEdges = New Dictionary
        dctNoThinEdges = New Dictionary
        WriteLogLine(Now() & " start collecting nodes and edges ")

        '  [050306] hyu: transitSegment
        '  getNodes tblTSeg, g_PSRCEdgeID, dctNoThinEdges

        '[041607]hyu: dealing w/ transit points instead of segments table
        'getNodes tblTSeg, g_INode, dctNoThinNodes, "SUBSTRING(CAST(LineID AS CHARACTER), 2, 2) = '" & Right(CStr(inserviceyear), 2) & "' AND SegOrder=1"
        'getNodes tblTSeg, g_JNode, dctNoThinNodes, "SUBSTRING(CAST(LineID AS CHARACTER), 2, 2) = '" & Right(CStr(inserviceyear), 2) & "' AND change>0"
        '[051407] jaf: use SQL with SDE but not with Access
        If strLayerPrefix = "SDE" Then
            getNodes(pTrPoints, g_PSRCJctID, dctNoThinNodes, "SUBSTRING(CAST(LineID AS CHARACTER), 2, 2) = '" & Right(CStr(inserviceyear), 2) & "'")
            'getPRNodes pFLayerJ.FeatureClass, "Scen_Node", "P_RStalls", dctNoThinNodes
        Else
            getNodes(pTrPoints, g_PSRCJctID, dctNoThinNodes, "MID(Cstr(LineID), 2, 2) = '" & Right(CStr(inserviceyear), 2) & "'")
        End If

        WriteLogLine(Now() & " finish collecting junctions being transit points: " & dctNoThinNodes.Count)

        getNodes(pTblMode, g_PSRCEdgeID, dctThinEdges, , True)
        WriteLogLine(Now() & " finish collecting edges can be thinned: " & dctThinEdges.Count)

        getNodes(pTolls, g_PSRCEdgeID, dctNoThinEdges)
        WriteLogLine(Now() & " finish collecting edges can NOT be thinned: " & dctNoThinEdges.Count)

        getNodes(pTurn, g_PSRCJctID, dctNoThinNodes)
        WriteLogLine(Now() & " finish collecting junction can NOT be thinned (including transit): " & dctNoThinNodes.Count)

        '[082409] SEC: Added Park and Rides
        getNodes(m_edgeShp, g_PSRCEdgeID, dctNoThinEdges, g_FacType & "=16 or " & g_FacType & "=17") '[042706] hyu: exclude centroid connector + park and rides from dissolving
        WriteLogLine(Now() & " finish collecting edges can NOT be thinned : " & dctNoThinEdges.Count)

        '[022406] hyu: stats how many edges joining to a junction
        Dim dctJoinEdges As Scripting.Dictionary
        Dim dctJcts As Scripting.Dictionary
        Dim dctEdgeID As Scripting.Dictionary
        Dim lMaxEdgeOID As Long
        dctJoinEdges = New Dictionary
        dctJcts = New Dictionary
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
        sNodes = dctJcts.Keys
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

            If Not dctNoThinNodes.Exists(sNode) Then
                If dctJoinEdges.Exists(sNode) Then
                    sEdgeID = Split(dctJoinEdges.Item(sNode), ",")

                    If UBound(sEdgeID) - LBound(sEdgeID) <> 1 Then
                        noMergeCount = noMergeCount + 1
                    Else '2 edges connected at this joint
                        If dctThinEdges.Exists(sEdgeID(0)) And dctThinEdges.Exists(sEdgeID(1)) And (Not (dctNoThinEdges.Exists(sEdgeID(0)) Or dctNoThinEdges.Exists(sEdgeID(1)))) Then

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
                                    If pFeat2.Value(pFeat2.Fields.FindField("Updated1")) = "Yes" Then
                                        pRow = getAttributesRow(pFeat2, Nothing, pTblPrjEdgeAtt, pEvtLine)
                                    Else
                                        pTC = pTblMode.Search(pQF2, False)
                                        pRow = pTC.NextRow
                                    End If

                                    pQF2.WhereClause = g_PSRCEdgeID & "=" & sEdgeID(1)
                                    If pNextFeat.Value(pNextFeat.Fields.FindField("Updated1")) = "Yes" Then
                                        pNextRow = getAttributesRow(pNextFeat, Nothing, pTblPrjEdgeAtt, pEvtLine)
                                    Else
                                        pTC = pTblMode.Search(pQF2, False)
                                        pNextRow = pTC.NextRow
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
                                                If (pRow.Value(i) <> pNextRow.Value(i)) Then
                                                    match = False
                                                    WriteLogLine(pRow.Fields.Field(i).Name + " diff")
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

                                    If pFeat2.Value(upIndex) <> pNextFeat.Value(upIndex) Then
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

        m_Map.ClearSelection()
        Dim PMx As IMxDocument
        PMx = m_App.Document
        PMx.ActiveView.Refresh()

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

        dctNoThinEdges.RemoveAll()
        dctThinEdges.RemoveAll()
        dctNoThinNodes.RemoveAll()
        dctJcts.RemoveAll()

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


   


    Public Sub getNodes(ByVal pTbl As ITable, ByVal NodeIDFld As String, ByVal dctNodes As Dictionary, Optional ByVal sWhereClause As String = "", Optional ByVal bModeAtt As Boolean = False)
        Dim fld As Long
        fld = pTbl.FindField(NodeIDFld)
        Dim pDS As IDataset
        pDS = pTbl
        Debug.Print("getnodes: " & pDS.Name & " start at " & Now())
        Dim pCs As ICursor, pRow As IRow
        Dim sNodeID As String

        If dctNodes Is Nothing Then dctNodes = New Dictionary
        If bModeAtt Then
            '        Dim pRow As IRow
            pCs = pTbl.Search(Nothing, True)
            pRow = pCs.NextRow
            Do Until pRow Is Nothing
                If Not IsDBNull(pRow.Value(fld)) Then
                    sNodeID = CStr(pRow.Value(fld))
                    If Not dctNodes.Exists(sNodeID) Then dctNodes.Add(sNodeID, sNodeID)
                End If
                pRow = pCs.NextRow
            Loop
            pCs = Nothing
            Debug.Print("getnodes: " & pDS.Name & " end at " & Now())
            Exit Sub
        End If
        '    Set pRow = pCs.NextRow
        '    Do Until pRow Is Nothing
        '
        '        If Not IsDBNull(pRow.value(fld)) Then
        '            sNodeID = CStr(pRow.value(fld))
        '            If Not dctNodes.Exists(sNodeID) Then
        '                dctNodes.Add sNodeID, sNodeID
        '            End If
        '        End If
        '        Set pRow = pCs.NextRow
        '    Loop

        If sWhereClause <> "" Then
            Dim pFilt As IQueryFilter
            pFilt = New QueryFilter
            pFilt.WhereClause = sWhereClause
            pCs = pTbl.Search(pFilt, False)
        Else
            pCs = pTbl.Search(Nothing, False)
        End If

        'Dim pEnumV As IEnumVariantSimple
        Dim pEnumerator As IEnumerator
        Dim pDataStat As IDataStatistics
        pDataStat = New DataStatistics

        : Try
            With pDataStat
                .Cursor = pCs
                .Field = NodeIDFld

                pEnumerator = .UniqueValues
            End With
            pEnumerator.Reset()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
       

        Dim v As Object

        Do While pEnumerator.MoveNext
            v = pEnumerator.Current
            sNodeID = CStr(v)
            If Not dctNodes.Exists(sNodeID) Then dctNodes.Add(sNodeID, sNodeID)

            'v = pEnumerator.MoveNext

        Loop
        Debug.Print("getnodes: " & pDS.Name & " end at " & Now())

        pCs = Nothing
        pEnumerator = Nothing
        pDataStat = Nothing
        pDS = Nothing

    End Sub
    Public Sub getPRNodes(ByVal pFC As IFeatureClass, ByVal NodeIDFld As String, ByVal StallsFld As String, ByVal dctNodes As Dictionary)
        Dim fld As Long
        fld = pFC.FindField(NodeIDFld)
        Dim pDS As IDataset
        pDS = pFC
        Debug.Print("getPRnodes: " & pDS.Name & " start at " & Now())
        Dim pFCursor As IFeatureCursor, pFeat As IFeature
        Dim sNodeID As String
        Dim indexStalls As Integer
        indexStalls = pFC.Fields.FindField(StallsFld)
        Dim pFilter As IQueryFilter
        pFilter = New QueryFilter
        pFilter.WhereClause = StallsFld & " > 0 "

        'If dctNodes Is Nothing Then Set dctNodes = New Dictionary(of Long, Long)
        'If bModeAtt Then
        '        Dim pFeat as ifeature
        pFCursor = pFC.Search(pFilter, True)
        pFeat = pFCursor.NextFeature
        Do Until pFeat Is Nothing
            If Not IsDBNull(pFeat.Value(fld)) Then
                sNodeID = CStr(pFeat.Value(fld))
                If Not dctNodes.Exists(sNodeID) Then dctNodes.Add(sNodeID, sNodeID)
            End If
            pFeat = pFCursor.NextFeature
        Loop
        pFCursor = Nothing
        Debug.Print("getnodes: " & pDS.Name & " end at " & Now())
        Exit Sub
        'End If
        '    Set pRow = pCs.NextRow
        '    Do Until pRow Is Nothing
        '
        '        If Not IsDBNull(pRow.value(fld)) Then
        '            sNodeID = CStr(pRow.value(fld))
        '            If Not dctNodes.Exists(sNodeID) Then
        '                dctNodes.Add sNodeID, sNodeID
        '            End If
        '        End If
        '        Set pRow = pCs.NextRow
        '    Loop

        'If sWhereClause <> "" Then
        'Dim pFilt As IQueryFilter
        'Set pFilt = New QueryFilter
        'pFilt.WhereClause = sWhereClause
        'Set pCs = pTbl.Search(pFilt, False)
        ' Else
        'Set pCs = pTbl.Search(Nothing, False)
        ' End If

        'Dim pEnumV As IEnumVariantSimple
        ' Dim pDataStat As IDataStatistics
        'Set pDataStat = New DataStatistics
        ' With pDataStat
        'Set .Cursor = pCs
        '.field = NodeIDFld

        'Set pEnumV = .UniqueValues
        'End With

        Dim v As Object
        'v = pEnumV.Next
        'Do Until IsEmpty(v)
        'sNodeID = CStr(v)

        ' 'If Not dctNodes.Exists(sNodeID) Then dctNodes.Add sNodeID, sNodeID
        ' v = pEnumV.Next
        'Loop
        'Debug.Print "getnodes: " & pDS.name & " end at " & Now()

        'Set pCs = Nothing
        'Set pEnumV = Nothing
        'Set pDataStat = Nothing
        'Set pDS = Nothing

    End Sub



    'Private Sub calcEdgesAtJoint(pEdges As ITable, INodeFld As String, JNodeFld As String, _
    'dctEdges As Dictionary(of Long, Long), dctJcts As Dictionary(of Long, Long), dctEdgeID As Dictionary(of Long, Long))
    Public Sub calcEdgesAtJoint(ByVal pEdges As IFeatureClass, ByVal pJcts As IFeatureClass, ByVal INodeFld As String, _
        ByVal JNodeFld As String, ByVal dctEdges As Dictionary, ByVal dctJcts As Dictionary, ByVal dctEdgeID As Dictionary, ByVal lMaxEdgeOID As Long)
        Dim pCs As IFeatureCursor, pFt As IFeature
        Dim fldI As Long, fldJ As Long, fldID As Long, fldJctID As Long
        Dim sI As String, sJ As String, sID As String, sJctID As String

        fldI = pEdges.FindField(INodeFld)
        fldJ = pEdges.FindField(JNodeFld)
        fldID = pEdges.FindField(g_PSRCEdgeID)
        fldJctID = pJcts.FindField(g_PSRCJctID)

        If dctEdges Is Nothing Then dctEdges = New Dictionary
        If dctJcts Is Nothing Then dctJcts = New Dictionary
        If dctEdgeID Is Nothing Then dctEdgeID = New Dictionary

        pCs = pJcts.Search(Nothing, False)
        pFt = pCs.NextFeature
        Try

        
            Do Until pFt Is Nothing
                sJctID = CStr(pFt.Value(fldJctID))
                dctJcts.Add(sJctID, pFt)
                pFt = pCs.NextFeature
            Loop
        Catch ex As Exception
            MessageBox.Show(ex.ToString)

        End Try

        pCs = pEdges.Search(Nothing, False)
        pFt = pCs.NextFeature
        lMaxEdgeOID = pFt.OID
        Do Until pFt Is Nothing
            If pFt.OID > lMaxEdgeOID Then lMaxEdgeOID = pFt.OID
            If Not (IsDBNull(pFt.Value(fldI)) Or IsDBNull(pFt.Value(fldJ)) Or IsDBNull(pFt.Value(fldID))) Then
                sI = CStr(pFt.Value(fldI))
                sJ = CStr(pFt.Value(fldJ))
                sID = CStr(pFt.Value(fldID))

                '            If Not dctEdgeID.Exists(sID) Then dctEdgeID.Add sID, 0
                If Not dctEdgeID.Exists(sID) Then dctEdgeID.Add(sID, pFt)

                '            If sI = "152515" Or sJ = "152515" Then MsgBox "here"
                If Not dctEdges.Exists(sI) Then
                    dctEdges.Add(sI, sID) 'CStr(pFt.oid)
                    '                dctJcts.Add sI, 1
                Else
                    dctEdges.Item(sI) = dctEdges.Item(sI) & "," & sID 'CStr(pFt.oid)
                    '                dctJcts.Item(sI) = dctJcts.Item(sI) + 1
                End If

                If Not dctEdges.Exists(sJ) Then
                    dctEdges.Add(sJ, sID) 'CStr(pFt.oid)
                    '                dctJcts.Add sJ, 1
                Else
                    dctEdges.Item(sJ) = dctEdges.Item(sJ) & "," & sID 'CStr(pFt.oid)
                    '                dctJcts.Item(sJ) = dctJcts.Item(sJ) + 1
                End If

            End If

            pFt = pCs.NextFeature
        Loop
        pCs = Nothing
    End Sub


    Private Sub getOID2Dict(ByVal pTbl As ITable, ByVal dctIDs As Dictionary)
        Dim pDS As IDataset
        pDS = pTbl
        Debug.Print("getOID2Dict: " & pDS.Name & " start at " & Now())
        Dim pCs As ICursor, pRow As IRow
        Dim sID As String
        pCs = pTbl.Search(Nothing, False)
        pRow = pCs.NextRow
        Do Until pRow Is Nothing

            sID = CStr(pRow.OID)
            dctIDs.Add(sID, 0)
            pRow = pCs.NextRow
        Loop
        Debug.Print("getOID2Dict: " & pDS.Name & " end at " & Now())

        pCs = Nothing
        '    Set pEnumV = Nothing
        '    Set pDataStat = Nothing
        pDS = Nothing

    End Sub

    Public Sub updateEdgesAtJoint(ByVal dctEdge As Dictionary, ByVal pEdge1 As IFeature, ByVal pEdge2 As IFeature, ByVal pNewEdge As IFeature)
        'pEdge1 and pEdge2 are the deleted edges
        'pNewEdge is the new edge.

        Dim fldI As Long, fldJ As Long, fldID As Long
        fldI = pEdge1.Fields.FindField(g_INode)
        fldJ = pEdge1.Fields.FindField(g_JNode)
        fldID = pEdge1.Fields.FindField(g_PSRCEdgeID)

        removeEdgeAtJoint(dctEdge, CStr(pEdge1.Value(fldI)), CStr(pEdge1.Value(fldID))) ' CStr(pEdge1.oid)
        removeEdgeAtJoint(dctEdge, CStr(pEdge1.Value(fldJ)), CStr(pEdge1.Value(fldID))) 'CStr(pEdge1.oid)
        removeEdgeAtJoint(dctEdge, CStr(pEdge2.Value(fldI)), CStr(pEdge2.Value(fldID))) 'CStr(pEdge2.oid)
        removeEdgeAtJoint(dctEdge, CStr(pEdge2.Value(fldJ)), CStr(pEdge2.Value(fldID))) 'CStr(pEdge2.oid)

        'add the new edge's oid to its i/j node's edge list.
        Dim sKey As String
        Dim sID As String
        sID = CStr(pNewEdge.Value(fldID))
        sKey = CStr(pNewEdge.Value(fldI))

        If Trim(dctEdge.Item(sKey)) = "" Then
            dctEdge.Item(sKey) = sID ' CStr(pNewEdge.oid)
        Else
            dctEdge.Item(sKey) = dctEdge.Item(sKey) & "," & sID 'CStr(pNewEdge.oid)
        End If

        sKey = CStr(pNewEdge.Value(fldJ))
        If Trim(dctEdge.Item(sKey)) = "" Then
            dctEdge.Item(sKey) = sID ' CStr(pNewEdge.oid)
        Else
            dctEdge.Item(sKey) = dctEdge.Item(sKey) & "," & sID 'CStr(pNewEdge.oid)
        End If
    End Sub

    Private Sub removeEdgeAtJoint(ByVal dct As Dictionary, ByVal JunctionID As String, ByVal edgeID As String)
        Dim sEIDs As String
        Dim sSplit() As String
        '    If JunctionID = "152515" Then MsgBox "here"
        sEIDs = dct.Item(JunctionID)
        sSplit = Split(sEIDs, ",")

        Dim i As Integer
        Dim sEIDsNew As String
        sEIDsNew = ""
        For i = LBound(sSplit) To UBound(sSplit)
            If sSplit(i) <> edgeID Then sEIDsNew = sEIDsNew + "," + sSplit(i)

        Next i
        If sEIDsNew <> "" Then sEIDsNew = Mid(sEIDsNew, 2, Len(sEIDsNew))

        dct.Item(CStr(JunctionID)) = sEIDsNew

    End Sub


    Private Sub replaceFieldValue(ByVal pFCls As IFeatureClass, ByVal FieldName As String, ByVal oldV As Long, ByVal newV As Long)
        Dim pFilt As IQueryFilter
        pFilt = New QueryFilter
        pFilt.WhereClause = FieldName & "=" & oldV

        Dim pCs As ICursor
        Dim pR As IRow
        Dim fld As Long
        fld = pFCls.FindField(FieldName)
        pCs = pFCls.Search(pFilt, False)
        pR = pCs.NextRow
        Do Until pR Is Nothing
            '        If pR.value(3) = 131057 Then MsgBox "replace value from " & CStr(oldV) & " to " & CStr(newV)
            pR.value(fld) = newV
            pR.Store()
            pR = pCs.NextRow
        Loop

        pCs = Nothing
        pFilt = Nothing
    End Sub


    Public Function CreateScenarioEdge(ByVal pInputFCls As IFeatureClass, ByVal pFWS As IFeatureWorkspace, ByVal OutputName As String) As IFeatureClass
        Dim pFlds As IFields
        Dim pClone As IClone
        pClone = pInputFCls.Fields
        pFlds = pClone.Clone

        Dim pOutFCls As IFeatureClass
        pOutFCls = pFWS.CreateFeatureClass(OutputName, pFlds, Nothing, Nothing, esriFeatureType.esriFTSimple, pInputFCls.ShapeFieldName, "")
        addField(pOutFCls, "PSRC_E2ID", esriFieldType.esriFieldTypeInteger, , 0)
        addField(pOutFCls, "ScenarioID", esriFieldType.esriFieldTypeInteger)
        addField(pOutFCls, "Scen_Link", esriFieldType.esriFieldTypeInteger)

        addField(pOutFCls, "TR_I", esriFieldType.esriFieldTypeInteger)
        addField(pOutFCls, "TR_J", esriFieldType.esriFieldTypeInteger)
        addField(pOutFCls, "HOV_I", esriFieldType.esriFieldTypeInteger)
        addField(pOutFCls, "HOV_J", esriFieldType.esriFieldTypeInteger)
        addField(pOutFCls, "TK_I", esriFieldType.esriFieldTypeInteger)
        addField(pOutFCls, "TK_J", esriFieldType.esriFieldTypeInteger)
        addField(pOutFCls, "UseEmmeN", esriFieldType.esriFieldTypeInteger)
        addField(pOutFCls, "SplitHOV", esriFieldType.esriFieldTypeString)
        addField(pOutFCls, "SplitTR", esriFieldType.esriFieldTypeString)
        addField(pOutFCls, "SplitTK", esriFieldType.esriFieldTypeString)

        'these fields used to flag that a future project will change the attributes
        'meaning to not get the attributes from modeAttributes table but instead the tblLine/PointProjects table
        addField(pOutFCls, "Updated1", esriFieldType.esriFieldTypeString, 10, "No")
        addField(pOutFCls, "prjRte", esriFieldType.esriFieldTypeString, 20)
        addField(pOutFCls, "shptype", esriFieldType.esriFieldTypeString, 20)
        addField(pOutFCls, "Dissolve", esriFieldType.esriFieldTypeString.esriFieldTypeInteger, , 0)

        'Stefan: add a field to flag if the IJ nodes are reversed in the project route

        addField(pOutFCls, "Direction", esriFieldType.esriFieldTypeString.esriFieldTypeInteger, , 0)
        CreateScenarioEdge = pOutFCls


        pOutFCls = Nothing
        pFlds = Nothing
        pClone = Nothing
    End Function

    Public Function CreateScenarioJunct(ByVal pInputFCls As IFeatureClass, ByVal pFWS As IFeatureWorkspace, ByVal OutputName As String) As IFeatureClass
        Dim pFlds As IFields
        Dim pClone As IClone
        pClone = pInputFCls.Fields
        pFlds = pClone.Clone

        Dim pOutFCls As IFeatureClass

        pOutFCls = pFWS.CreateFeatureClass(OutputName, pFlds, Nothing, Nothing, esriFeatureType.esriFTSimple, pInputFCls.ShapeFieldName, "")

        addField(pOutFCls, "PSRC_E2ID", esriFieldType.esriFieldTypeInteger, , 0)
        addField(pOutFCls, "ScenarioID", esriFieldType.esriFieldTypeInteger)
        addField(pOutFCls, "Scen_Node", esriFieldType.esriFieldTypeInteger)
        CreateScenarioJunct = pOutFCls

        pOutFCls = Nothing
        pFlds = Nothing
        pClone = Nothing
    End Function

    Public Function ExportFeature(ByVal pFromFCS As IFeatureCursor, ByVal pToFCls As IFeatureClass)

        Dim pIns As IFeatureCursor
        Dim pFBuf As IFeatureBuffer

        pIns = pToFCls.Insert(True)

        Dim pFeat As IFeature
        Dim i As Integer, j As Integer
        pFeat = pFromFCS.NextFeature
        Try

        
            Do Until pFeat Is Nothing

                pFBuf = pToFCls.CreateFeatureBuffer
                With pFBuf
                    .Shape = pFeat.ShapeCopy
                    For i = 2 To pFromFCS.Fields.FieldCount - 1
                        '[051407] jaf: cannot assign null values
                        '[051807] hyu: comparing field name to make sure the value is assigned to correct field.
                        If (pFeat.Fields.Field(i).Editable = True) And (pFeat.Fields.Field(i).Type <> esriFieldType.esriFieldTypeGeometry) And Not (IsDBNull(pFeat.Value(i))) Then
                            .Value(i) = pFeat.Value(pFeat.Fields.FindField(.Fields.Field(i).Name))
                            '                    For j = 1 To pFeat.Fields.FieldCount - 1
                            '                        If UCase(.Fields.field(i).name) = UCase(pFeat.Fields.field(j).name) Then
                            '                            .value(i) = pFeat.value(j)
                            '                        End If
                            '                    Next j
                        End If
                    Next i
                End With
                pIns.InsertFeature(pFBuf)

                pFeat = pFromFCS.NextFeature
            Loop
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        pFBuf = Nothing
        pIns = Nothing
        pFromFCS = Nothing
    End Function

    Public Function UpdateFields(ByVal pEdgeTbl As ITable, ByVal pJctTbl As ITable)
        Dim pCs As ICursor
        Dim pRow As IRow
        Dim fldSceID As Long, fldSceLink As Long, fldPSRCEdgeID As Long, fldPSRCE2Id As Long, fldINode As Long, fldJNode As Long, fldUseEmmeN As Long
        Dim fldSceID2 As Long, fldSceNode As Long, fldPSRCJctID As Long, fldEmme2NodeId As Long, fldPSRCE2Id2 As Long
        Dim dctEmme2Node As Dictionary 'key=PSRCJunctID, item=Emme2NodeID; list of the reserved nodes for Emme2 centroids, park and ride
        dctEmme2Node = New Dictionary

        With pEdgeTbl
            fldSceID = .FindField("ScenarioID")
            fldSceLink = .FindField("Scen_Link")
            fldPSRCEdgeID = .FindField(g_PSRCEdgeID)
            fldPSRCE2Id = .FindField("PSRC_E2ID")
            fldINode = .FindField(g_INode)
            fldJNode = .FindField(g_JNode)
            fldUseEmmeN = .FindField("UseEmmeN")

        End With

        With pJctTbl
            fldSceID2 = .FindField("ScenarioID")
            fldSceNode = .FindField("Scen_Node")
            fldPSRCJctID = .FindField(g_PSRCJctID)
            fldEmme2NodeId = .FindField("EMME2NODEID")
            fldPSRCE2Id2 = .FindField("PSRC_E2ID")
        End With

        'Set pCs = pJctTbl.Update(Nothing, False)
        pCs = pJctTbl.Search(Nothing, False)
        pRow = pCs.NextRow
        Dim lEmme2NodeID As Long
        Dim lPSRCJctID As Long

        Do Until pRow Is Nothing
            lEmme2NodeID = IIf(IsDBNull(pRow.Value(fldEmme2NodeId)), 0, pRow.Value(fldEmme2NodeId))
            lPSRCJctID = pRow.value(fldPSRCJctID)

            pRow.value(fldSceID2) = m_ScenarioId
            pRow.value(fldPSRCE2Id2) = lEmme2NodeID

            'update the dictionary of reserved nodes for Emme2
            '[092707] hyu: per Jeff's email on 092607. the Emme2NodeID should be less than the user defined offset
            'If lEmme2NodeID > 0 Then
            If lEmme2NodeID > 0 And lEmme2NodeID <= m_Offset Then
                If Not dctEmme2Node.Exists(CStr(lPSRCJctID)) Then dctEmme2Node.Add(CStr(lPSRCJctID), lEmme2NodeID)
                pRow.value(fldSceNode) = lEmme2NodeID
            Else
                pRow.value(fldSceNode) = pRow.value(fldPSRCJctID) + m_Offset
            End If

            pRow.Store()
            pRow = pCs.NextRow
        Loop

        WriteLogLine("Reserved nodes for Emme/2 Centroids and park and rides: " & dctEmme2Node.count)

        'Set pCs = pEdgeTbl.Update(Nothing, False)
        pCs = pEdgeTbl.Search(Nothing, False)
        pRow = pCs.NextRow
        Do Until pRow Is Nothing
            '        If pRow.value(pRow.Fields.FindField("PSRCEDGEID")) = 52888 Then MsgBox 1
            pRow.value(fldSceLink) = pRow.value(fldPSRCEdgeID) + 1
            pRow.value(fldPSRCE2Id) = 0
            pRow.value(fldSceID) = m_ScenarioId

            'Update "UseEmmeN" to denote whether INode and Jnode are reserved nodes for Emme2 Centroid, park and ride
            'UseEmmeN=0: Neither INode nor JNode is reserved
            'UseEmmeN=1: Only INode is reserved
            'UseEmmeN=2: Only JNode is reserved
            'UseEmmeN=3: Both INode and Jnode are reserved
            If dctEmme2Node.Exists(CStr(pRow.value(fldINode))) Then
                If dctEmme2Node.Exists(CStr(pRow.value(fldJNode))) Then
                    pRow.value(fldUseEmmeN) = 3
                Else
                    pRow.value(fldUseEmmeN) = 1
                End If
            Else
                If dctEmme2Node.Exists(CStr(pRow.value(fldJNode))) Then
                    pRow.value(fldUseEmmeN) = 2
                Else
                    pRow.value(fldUseEmmeN) = 0
                End If
            End If

            pRow.Store()
            If IsDBNull(pRow.Value(fldSceID)) Then MsgBox(1)
            pRow = pCs.NextRow
        Loop

        pCs = Nothing
        dctEmme2Node.RemoveAll()
        dctEmme2Node = Nothing
    End Function

    Private Function createCircle(ByVal pPoint As IPoint) As IPolygon4
        Dim pSegCol As ISegmentCollection
        pSegCol = New Polygon
        pSegCol.SetCircle(pPoint, 0.5)
        createCircle = pSegCol
    End Function

    Public Function MergeEdges(ByVal pNextFeat As IFeature, ByVal pFeat2 As IFeature) As IPolyline4
        'pFeat2's geometry is the base of the merged geometry.  pNextFeat's geometry may be flipped based on how they coincide.
        Dim pSegCol As ISegmentCollection, pSegCol2 As ISegmentCollection
        Dim pPl As IPolyline4, pPl2 As IPolyline4
        Dim pPt2 As IPoint, dAlong As Double, dFrom As Double, bright As Boolean

        '[073007] hyu: new appraoch to merge edges:
        '[073007] hyu: consider the geometry orientation when merging two edges. Replace the old approach of using iTopologicalOperator.ConstructUnion.

        pPl = pNextFeat.ShapeCopy
        pPl2 = pFeat2.ShapeCopy

        pPl2.QueryPointAndDistance(esriSegmentExtension.esriNoExtension, pPl.FromPoint, True, pPt2, dAlong, dFrom, bright)
        pSegCol = pPl
        pSegCol2 = pPl2
        If Math.Round(dFrom, 3) = 0 Then 'pNextFeat's from point is the coincidental point of the pFeat2
            If Math.Round(dAlong, 3) = 0 Then 'pNextFeat's from point coincidents w/ pFeat2's from point
                'need's flip pNextFeat's geometry
                'the final geometry is from pNextFeat to pFeat2
                pPl.ReverseOrientation()
                pSegCol2.InsertSegmentCollection(0, pSegCol)
            Else    'pNextFeat's from point coincidents w/ pFeat2's to point
                'no flip, the final geometry is from pFeat2 to pNextFeat
                pSegCol2.AddSegmentCollection(pSegCol)
            End If
        Else    'pNextFeat's to point is the coincidental point of the pFeat2
            pPl2.QueryPointAndDistance(esriSegmentExtension.esriNoExtension, pPl.ToPoint, True, pPt2, dAlong, dFrom, bright)
            If Math.Round(dAlong, 3) = 0 Then 'pNextFeat's to point coincidents w/ pFeat2's from point
                'no flip, the final geometry is from pNextFeat to pFeat2
                pSegCol2.InsertSegmentCollection(0, pSegCol)
            Else    'pNextFeat's to point coincidents w/ pFeat2's to point
                'need's flip pNextFeat's geometry, the final geometry is from pFeat2 to pNextFeat
                pPl.ReverseOrientation()
                pSegCol2.AddSegmentCollection(pSegCol)
            End If
        End If

        MergeEdges = New Polyline
        MergeEdges = pSegCol2

        pPl = Nothing
        pSegCol = Nothing
    End Function
    Public Sub dissolveTransEdgesNoX()
        '[022706] hyu: revised function to replace [dissolveTransEdgesNoX_0]
        'performs "pseudonode" thinning of the intermediate edge shp file
        'this version does NOT cross check with Tolls, Transit, and Turns
        '[061505] jaf removed transit, toll, and turn cross-check

        'on error GoTo ErrChk

        WriteLogLine("")
        WriteLogLine("======================================")
        WriteLogLine("dissolveTransEdgesNoX started " & Now())
        WriteLogLine("======================================")
        WriteLogLine("")

        If (strLayerPrefix = "SDE") Then
            Dim pwf As IWorkspaceFactory
            pwf = New SdeWorkspaceFactory
        End If

        Dim pFLayerE As IFeatureLayer, pFLayerJ As IFeatureLayer
        pFLayerE = New FeatureLayer
        pFLayerE.FeatureClass = m_edgeShp 'Intermediate Edge
        pFLayerJ = New FeatureLayer
        pFLayerJ.FeatureClass = m_junctShp 'Intermediate Junction
        Dim pWorkspace As IWorkspace
        pWorkspace = get_Workspace()
        Dim pTblMode As ITable
        pTblMode = get_TableClass(m_layers(2)) 'modeAttributes
        '[061505] turn, toll, and transit object vars removed here

        Dim pFeatCursor As IFeatureCursor
        Dim pFC As IFeatureCursor, pFC2 As IFeatureCursor
        Dim pFCj As IFeatureCursor
        Dim pQF As IQueryFilter, pQFt As IQueryFilter, pQF2 As IQueryFilter
        Dim pQFj As IQueryFilter

        Dim pGeometry As IGeometry, pNewGeometry As IGeometry
        Dim pGeometryBag As IGeometryCollection
        pGeometryBag = New GeometryBag
        Dim pTopo As ITopologicalOperator
        Dim pPolyline As IPolyline

        Dim pFeat As IFeature, pFeat2 As IFeature
        Dim pNextFeat As IFeature, pNewFeat As IFeature
        Dim pJunctFeat As IFeature
        Dim pRow As IRow, pNextRow As IRow
        Dim pTC As ICursor
        Dim pFlds As IFields
        Dim lSFld As Long, dnode As Long, idIndex As Long
        Dim lIndex As Long, upIndex As Long, facIndex As Long
        Dim psrcId As String
        Dim i As Long
        Dim fcountI, fcountIJ, mergeCountIJ, mergeCountJI As Long, noMergeCount As Long, attMismatchCount As Long
        Dim idxJunctID As Integer
        Dim match As Boolean
        match = True
        Dim edgesdelete As Boolean
        edgesdelete = False

        Dim newI As Long, newJ As Long
        Dim pWorkspaceEdit As IWorkspaceEdit
        pWorkspaceEdit = pWorkspace
        Dim t As Long
        Dim l As Long

        '[022706] hyu: insert this block to replace the old use of selection set.
        ' use a dictionary to dynamically maintain edges' oid and status, and an cursor to search edges.
        ' the old method of using 2 selection sets and cursor dosn't work.
        ' The cursor is initially set from one selection set, then its position cannot be updated dynamically when edges are deleted or generated.
        Dim dctEdges As Dictionary '(key=PSRCJunctID, Item=joined edges' PSRCEdgeID)
        Dim dctJcts As Dictionary
        Dim dctEdgeID As Dictionary
        Dim bEOFDict As Boolean
        dctEdges = New Dictionary
        'getOID2Dict m_edgeShp, dctEdges
        Debug.Print("CalcEdges@joint: " & Now())

        'commented out code to make it run
        'calcEdgesAtJoint(m_edgeShp, g_INode, g_JNode, dctEdges, dctJcts, dctEdgeID)
        Debug.Print("Dissolve: " & Now())
        bEOFDict = False
        idIndex = m_edgeShp.FindField(g_PSRCEdgeID)
        upIndex = m_edgeShp.FindField("Updated1")
        facIndex = m_edgeShp.FindField(g_FacType)

        pWorkspaceEdit = g_FWS
        pWorkspaceEdit.StartEditing(False)
        pWorkspaceEdit.StartEditOperation()
        Dim pFSelect As IFeatureSelection
        Dim pSSet As ISelectionSet, pSSet2 As ISelectionSet

        '    Set pFSelect = pFLayerE
        '    'pFSelect.SelectFeatures Nothing, esriSelectionResultNew, False 'get all intermediate edges
        '
        '    Set pSSet = m_edgeShp.Select(Nothing, esriSelectionTypeIDSet, esriSelectionOptionNormal, pWorkSpace)  'pFSelect.SelectionSet
        '    Set pSSet2 = m_edgeShp.Select(Nothing, esriSelectionTypeIDSet, esriSelectionOptionNormal, pWorkSpace) 'pFSelect.SelectionSet
        '
        '    pSSet.Search Nothing, False, pFeatCursor
        '    Set pFeat = pFeatCursor.NextFeature
        '    '[011806] pan currently no PSRC_E2ID in edges or modeAttributes.
        '    'idIndex = pFeat.Fields.FindField("PSRC_E2ID")
        '    idIndex = pFeat.Fields.FindField("PSRCEDGEID")
        '    upIndex = pFeat.Fields.FindField("Updated1")

        mergeCountIJ = 0
        mergeCountJI = 0
        noMergeCount = 0
        attMismatchCount = 0

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

        Dim lEdgeCt As Long, sEdgeIDs(), sJctIDs()
        Dim pFeat3 As IFeature
        lEdgeCt = 0
        'sEdgeIDs = dctEdges.Keys
        sJctIDs = dctEdges.Keys

        '[030906] hyu: change the way to loop through ScenarioEdges by looping through a dynamic edge dictionary
        Dim sEdges() As String
        Dim dum1 As Long, dum2 As Long, dum As Long
        Dim lJctCt As Long, sNode As String
        pQF = New QueryFilter
        pQF2 = New QueryFilter
        pQFj = New QueryFilter
        For lJctCt = 0 To dctEdges.count - 1

            sNode = CStr(sJctIDs(lJctCt))
            dnode = Val(sNode)

            sEdges = Split(dctEdges.Item(sNode), ",")
            If UBound(sEdges) - LBound(sEdges) + 1 <> 2 Then
                noMergeCount = noMergeCount + 1
            Else
                pQF.WhereClause = g_PSRCEdgeID & "=" & sEdges(0) & " OR " & g_PSRCEdgeID & "=" & sEdges(1)
                pFC = m_edgeShp.Search(pQF, False)
                pFeat2 = pFC.NextFeature
                pNextFeat = pFC.NextFeature
                pFeat = pFeat2

                If Not (pFeat2 Is Nothing Or pNextFeat Is Nothing) Then
                    '[042706] hyu: exclude centroid connector from dissolving
                    If pFeat2.value(facIndex) <> 16 And pNextFeat.value(facIndex) <> 16 Then
                        If pFeat.value(m_edgeShp.FindField(g_INode)) = dnode Then
                            m = 0
                        Else
                            m = 1
                        End If
                        '    Do Until bEOFDict
                        '        For m = 0 To 1 'first look at this edges INode and then its JNode to see if should dissolve
                        '            'if node dissolved should exit for loop
                        match = True
                        edgesdelete = False
                        If m = 0 Then
                            same = "INode"
                            opposite = "JNode"
                        Else
                            same = "JNode"
                            opposite = "INode"
                        End If

                        '            'get inode
                        '            lSFld = pFeat.Fields.FindField(same)
                        '            dnode = pFeat.value(lSFld)
                        ''            If pFeat.value(m_edgeShp.FindField(g_PSRCEdgeID)) = 201372 _
                        'SEC-commented out to make code run
                        'Or pFeat.value(m_edgeShp.FindField(g_PSRCEdgeID)) = 131057 Then MsgBox "HERE" '201372
                        If (dnode > 0) Then
                            'get jnode
                            WriteLogLine("potential dnode " & CStr(dnode) & "  " & Now)
                            lSFld = pFeat.Fields.FindField(opposite)
                            If m = 0 Then
                                newJ = pFeat.Value(lSFld)
                            Else
                                newI = pFeat.Value(lSFld)
                            End If

                            If pNextFeat.Value(m_edgeShp.FindField(same)) = dnode Then

                                lSFld = pNextFeat.Fields.FindField(opposite)
                                dir = False
                                If m = 0 Then
                                    newI = pNextFeat.Value(lSFld)
                                Else
                                    newJ = pNextFeat.Value(lSFld)
                                End If
                            Else
                                lSFld = pNextFeat.Fields.FindField(same)
                                dir = True
                                If m = 0 Then
                                    newI = pNextFeat.Value(lSFld)
                                Else
                                    newJ = pNextFeat.Value(lSFld)
                                End If
                            End If

                            pTC = pTblMode.Search(pQFt, False)
                            '[031606] hyu: remove the double query, add a judgement on the pNextRow
                            'count the presence of records using this edge in modeAttributes, modeTolls, and tblTransitSegments
                            '                        If (pTblMode.rowcount(pQFt) > 0) Then 'found first edge in modeAttributes {[061505] toll and transit check removed here}
                            '[060705] pre-Transit Xchk code: If (pTblMode.rowcount(pQFt) > 0 And pTolls.rowcount(pQFt) = 0) Then 'found first edge in modeAttributes and NOT in modeTolls

                            pRow = pTC.NextRow
                            pNextRow = pTC.NextRow
                            If Not (pRow Is Nothing Or pNextRow Is Nothing) Then

                                'compare attributes to detect if we have "real" pseudonode
                                '[031406] hyu: "FFS" attributes are excluded.
                                If dir = True Then
                                    For i = 1 To pRow.Fields.FieldCount - 1
                                        If (UCase(pRow.Fields.Field(i).Name) <> UCase("PSRCEdgeID")) And _
                                            (UCase(pRow.Fields.Field(i).Name) <> "IJFFS") And _
                                            (UCase(pRow.Fields.Field(i).Name) <> "JIFFS") Then

                                            If (pRow.Value(i) <> pNextRow.Value(i)) Then
                                                match = False
                                                WriteLogLine(pRow.Fields.Field(i).Name + " diff")
                                            End If

                                        End If
                                    Next i
                                Else 'need to compare IJ to JI to match direction
                                    Dim j As Long 'l As Long
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


                                        '                                        '[011806] pan FFS floating point comparison removed
                                        '                                        index2 = pRow.Fields.FindField(dir2 + "FFS")
                                        '                                        indexNext = pNextRow.Fields.FindField(dirNext + "FFS")
                                        '                                        If (pRow.value(index2) <> pNextRow.value(indexNext)) Then
                                        '                                            match = False
                                        ''                                              MsgBox "ffs diff"
                                        '                                        End If
                                        '

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
                                                    WriteLogLine("dir lanes " & lType(l) & " diff")
                                                End If
                                            Else
                                                index2 = pRow.Fields.FindField(dir2 + "LaneCap" + lType(l))
                                                indexNext = pNextRow.Fields.FindField(dirNext + "LaneCap" + lType(l))
                                                If (pRow.Value(index2) <> pNextRow.Value(indexNext)) Then
                                                    match = False
                                                    WriteLogLine("dir lanecap " & lType(l) & " diff")
                                                End If
                                                For t = 0 To 4
                                                    index2 = pRow.Fields.FindField(dir2 + "Lanes" + lType(l) + timePd(t))
                                                    indexNext = pNextRow.Fields.FindField(dirNext + "Lanes" + lType(l) + timePd(t))

                                                    If (pRow.Value(index2) <> pNextRow.Value(indexNext)) Then
                                                        match = False
                                                        WriteLogLine("dir lanes " & lType(l) & timePd(t) & " diff")
                                                    End If
                                                Next t
                                            End If
                                        Next l
                                    Next j
                                End If

                                If pFeat.Value(upIndex) <> pNextFeat.Value(upIndex) Then
                                    match = False
                                    WriteLogLine("update diff")
                                End If
                                i = pFeat.Fields.FindField("Modes")
                                If pFeat.Value(i) <> pNextFeat.Value(i) Then
                                    match = False
                                    WriteLogLine("modes diff")
                                End If
                                i = pFeat.Fields.FindField("LinkType")
                                If pFeat.Value(i) <> pNextFeat.Value(i) Then
                                    match = False
                                    WriteLogLine("link diff")
                                End If

                                i = pFeat.Fields.FindField("FacilityType")
                                If pFeat.Value(i) <> pNextFeat.Value(i) Then
                                    match = False
                                    WriteLogLine("fac diff")
                                End If

                                If (match = False) Then
                                    attMismatchCount = attMismatchCount + 1
                                Else
                                    WriteLogLine("found match")
                                    '[073007] hyu: new appraoch to merge edges:
                                    '[073007] hyu: consider the geometry orientation when merging two edges. Replace the old approach of using iTopologicalOperator.ConstructUnion.
                                    pPolyline = MergeEdges(pNextFeat, pFeat2)

                                    '                                    Set pGeometryBag = New GeometryBag
                                    '                                    Set pGeometry = pNextFeat.ShapeCopy
                                    '                                    pGeometryBag.AddGeometry pGeometry
                                    '                                    Set pGeometry = pFeat2.ShapeCopy
                                    '                                    pGeometryBag.AddGeometry pGeometry

                                    'update the merge count for the log
                                    If dir Then
                                        mergeCountIJ = mergeCountIJ + 1
                                    Else
                                        mergeCountJI = mergeCountJI + 1
                                    End If

                                    'merge the edge features
                                    '                                    Set pPolyline = New Polyline
                                    '                                    Set pTopo = pPolyline
                                    '                                    pTopo.ConstructUnion pGeometryBag
                                    pNewFeat = pFLayerE.FeatureClass.CreateFeature
                                    pNewFeat.Shape = pPolyline
                                    pNewFeat.Store()

                                    pFlds = pNewFeat.Fields
                                    For i = 1 To pFlds.FieldCount - 1
                                        lIndex = pFeat2.Fields.FindField(pFlds.Field(i).Name)
                                        If (lIndex <> -1) Then
                                            If (pFlds.Field(i).Type <> esriFieldType.esriFieldTypeOID And _
                                            pFlds.Field(i).Type <> esriFieldType.esriFieldTypeGeometry And _
                                            UCase(pFlds.Field(i).Name) <> UCase("Shape")) Then
                                                'inode and jnode need to be set differently
                                                '                                                If (pFlds.field(i).name = same) Then
                                                If UCase(pFlds.Field(i).Name) = UCase(g_INode) Then
                                                    pNewFeat.Value(i) = newI
                                                    '                                                ElseIf (pFlds.field(i).name = opposite) Then
                                                ElseIf UCase(pFlds.Field(i).Name) = UCase(g_JNode) Then
                                                    pNewFeat.Value(i) = newJ
                                                ElseIf UCase(pFlds.Field(i).Name) = UCase("length") Then
                                                    pNewFeat.Value(i) = pPolyline.Length
                                                Else
                                                    '[022606] hyu: add a judgement of whether field is editable
                                                    If pNewFeat.Fields.Field(i).Editable Then pNewFeat.Value(i) = pFeat2.Value(lIndex)
                                                End If
                                            End If
                                        End If
                                    Next i
                                    'pan edit session incompatible with Store
                                    '[121605]pan re-inserted store and edge delete
                                    '022706] hyu: comment out the deletion of the two features
                                    pNewFeat.Value(pNewFeat.Fields.FindField("Dissolve")) = 1
                                    pNewFeat.Store()
                                    edgesdelete = True
                                    '                                    pFeat2.Delete
                                    '                                    pNextFeat.Delete
                                    'find junction node and delete it

                                    pQFj = New QueryFilter
                                    pQFj.WhereClause = "PSRCJunctID = " + CStr(dnode)
                                    pFCj = pFLayerJ.FeatureClass.Search(pQFj, False)
                                    pJunctFeat = pFCj.NextFeature
                                    '                                    If (pFLayerJ.FeatureClass.featurecount(pQFj) > 0) Then
                                    'MsgBox "found junction to delete"
                                    If Not pJunctFeat Is Nothing Then

                                        If fVerboseLog Then
                                            idxJunctID = pJunctFeat.Fields.FindField("PSRCJunctID")
                                            If idxJunctID < 0 Then
                                                WriteLogLine("Deleting intrmed. shp junct OID=" & pJunctFeat.OID)
                                            Else
                                                WriteLogLine("Deleting intrmed. shp junctID=" & pJunctFeat.Value(idxJunctID))
                                            End If
                                        End If
                                        pJunctFeat.Delete()
                                        '                                        pJunctFeat.value(pJunctFeat.Fields.FindField(g_PSRCJctID)) = 1
                                        '                                        pJunctFeat.Store
                                    End If
                                    '                                    pSSet.Add pnewfeat.oid
                                    '                                    'MsgBox pSSet.count


                                    '                                    dum1 = -1
                                    '                                    dum2 = -1
                                    '                                    dum = -1
                                    '                                    If Not IsDBNull(pFeat2.value(pFeat2.Fields.FindField("Dissolve"))) Then dum1 = pFeat2.value(pFeat2.Fields.FindField("Dissolve"))
                                    '                                    If Not IsDBNull(pNextFeat.value(pFeat2.Fields.FindField("Dissolve"))) Then dum2 = pNextFeat.value(pFeat2.Fields.FindField("Dissolve"))
                                    '
                                    '
                                    '                                    If dum1 = -1 And dum2 = -1 Then
                                    '                                        dum = lEdgeCt
                                    '                                    ElseIf dum1 = -1 And dum2 > -1 Then
                                    '                                        dum = dum2
                                    '                                    ElseIf dum1 > -1 And dum2 = -1 Then
                                    '                                        dum = dum1
                                    '                                    ElseIf dum1 > -1 And dum1 = dum2 Then
                                    '                                        dum = dum1
                                    '                                    ElseIf dum1 > -1 And dum1 <> dum2 Then
                                    '                                        dum = dum1
                                    '                                        'replace dum2 w/ dum1
                                    '                                        replaceFieldValue m_edgeShp, "Dissolve", dum2, dum
                                    '                                    End If
                                    '
                                    '                                    pFeat2.value(pFeat2.Fields.FindField("Dissolve")) = dum
                                    '                                    pFeat2.Store
                                    '                                    pNextFeat.value(pFeat2.Fields.FindField("Dissolve")) = dum
                                    '                                    pNextFeat.Store
                                    '
                                    '                                    If (pFLayerJ.FeatureClass.featurecount(pQFj) > 0) Then
                                    '                                        Set pFCj = pFLayerJ.FeatureClass.Search(pQFj, False)
                                    '                                        Set pJunctFeat = pFCj.NextFeature
                                    '                                        pJunctFeat.value(pJunctFeat.Fields.FindField("Dissolve")) = 1
                                    '                                        pJunctFeat.Store
                                    '                                    End If
                                    '                                    edgesdelete = True
                                End If
                            End If  'match = true
                        End If
                    End If
                End If
            End If


            If (edgesdelete = True) Then

                '[022706] hyu: keep dictionary updated about the edge status. 1 means it's been dissolved.
                updateEdgesAtJoint(dctEdges, pFeat2, pNextFeat, pNewFeat)
                dctJcts.Remove(sNode)
                dctEdgeID.Item(sEdges(0)) = 1
                dctEdgeID.Item(sEdges(1)) = 1
                '                dctEdges.Add CStr(pnewfeat.oid), 0
                '
                '                sEdgeIDs = dctEdges.Keys

                pFeat2.Delete()
                pNextFeat.Delete()

                edgesdelete = False
                'MsgBox "edges deleted"
                ''                Exit For
                '                '[031306] Hyu: instead of exiting for, continue with the opposite node
                '                Set pFeat = pnewfeat
            End If
            '        Next m
            '        lEdgeCt = lEdgeCt + 1
            '        If lEdgeCt >= dctEdges.count Then bEOFDict = True
            ''        Set pFeat = pFeatCursor.NextFeature
            '    Loop
        Next lJctCt

        pWorkspaceEdit.StopEditOperation()
        pWorkspaceEdit.StopEditing(True)

        m_Map.ClearSelection()
        Dim PMx As IMxDocument
        PMx = m_App.Document
        PMx.ActiveView.Refresh()

        pFLayerE = Nothing
        pFeat = Nothing
        pFeat2 = Nothing
        pNextFeat = Nothing
        pNewFeat = Nothing


        WriteLogLine("FINISHED dissolveTransEdgesNoX at " & Now())
        WriteLogLine("Junctions selected= " & dctEdges.count)
        WriteLogLine("Same dir thinned= " & mergeCountIJ & "; Opp dir thinned= " & mergeCountJI)
        WriteLogLine("Junctions thinned= " & mergeCountIJ + mergeCountJI)
        WriteLogLine("Junctions geometrically eligible to thin but not thinned due to attribute mismatch= " & attMismatchCount)
        WriteLogLine("Junctions geometrically inelegible to thin= " & noMergeCount)
        WriteLogLine("Junctions final= " & m_junctShp.FeatureCount(Nothing))

        t = 0
        For l = 0 To dctEdgeID.count - 1
            If dctEdgeID.Items(l) = 1 Then t = t + 1
        Next l

        WriteLogLine("Edges selected= " & l)
        WriteLogLine("Edges merged= " & t)

        pQF = New QueryFilter
        pQF.WhereClause = "Dissolve=1"
        WriteLogLine("Edges created= " & m_edgeShp.FeatureCount(pQF))
        WriteLogLine("Edges final= " & m_edgeShp.FeatureCount(Nothing))
        pQF = Nothing
        dctEdges = Nothing
        dctEdgeID = Nothing

        'Note: needs to be a round two version of this based on atrributes that don't match but are not significant
        'need to get these attributes!
        Exit Sub

ErrChk:
        CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.dissolveTransEdgesNoX")
        MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.dissolveTransEdgesNoX")
        'Print #1, Err.Description, vbInformation, "dissolveTransEdgesNoX"
        pFLayerE = Nothing
        pFeat = Nothing
        pFeat2 = Nothing
        pNextFeat = Nothing
        pNewFeat = Nothing
    End Sub
End Module
