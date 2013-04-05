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
Imports ESRI.ArcGIS.System
Imports Microsoft.SqlServer


Imports System.Math

Imports ESRI.ArcGIS.GeoDatabaseUI
Imports Scripting

Imports ESRI.ArcGIS.DataSourcesFile
Module hyuNetFile
    Public Sub SplitFeatureByM(ByVal pFeature As IFeature, ByVal dSplitM As Double, ByVal onode As Long, ByVal NewPoint As IPoint, ByVal lType As String, _
        ByVal dctEdges As Dictionary, ByVal dctJct As Dictionary, ByVal dctWeaveNodes As Dictionary, ByVal dctWeaveNodes2 As Dictionary, ByRef nodeString As String)
        'dSplitM is the weave length from FromPoint
        'onode is the base node, should coincide w/ one of the ends of pfeature.
        '[080807] hyu:
        'on error GoTo ErrChk
        If fVerboseLog Then WriteLogLine("called SplitFeatureByM")
        Dim tempSplitM As Double
        tempSplitM = dSplitM

        Dim pEnumVertex As IEnumVertex
        Dim pGeoColl As IGeometryCollection
        Dim pPolyCurve As IPolycurve2
        'Dim pEnumSplitPoint As IEnumSplitPoint
        Dim pNewFeature As IFeature
        Dim PartCount As Integer
        Dim index As Long, indexJ As Long
        Dim pFCls As IFeatureClass
        pFCls = m_junctShp

        Dim pFilter As ISpatialFilter
        Dim pFC As IFeatureCursor
        Dim pJFeat As IFeature

        Dim found As Boolean
        Dim SplitID As Long
        LargestJunct = LargestJunct + 1
        '[051407] hyu
        'SplitID = largestjunct
        SplitID = LargestJunct + m_Offset
        '    If SplitID = 2241 Then MsgBox 1

        Dim FromPt As IPoint, ToPt As IPoint
        Dim pPline As IPolyline
        Dim bSplit As Boolean, lPart As Long, lSeg As Long, bpart As Boolean, bratio As Boolean
        Dim pSplit1 As IPolyline, pSplit2 As IPolyline  'split1 is the long portion and Split2 is the short portion
        Dim pSegCol As ISegmentCollection
        Dim bUpdateJ As Boolean 'indicate whether the J node of the weave link part should get the new junction id.

        pSegCol = New Polygon
        pPline = pFeature.ShapeCopy
        pPolyCurve = pFeature.Shape
        FromPt = pPline.FromPoint
        ToPt = pPline.ToPoint
        pFilter = New SpatialFilter
        With pFilter
            'Set .Geometry = ToPt
            'SEC- commented code to make it run
            '.Geometry = pSegCol.SetCircle(FromPt, 0.5)
            .GeometryField = m_junctShp.ShapeFieldName
            .SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects
            .WhereClause = "Scen_Node = " & onode
        End With

        'If pFCls.featurecount(pFilter) > 0 Then
        pFC = pFCls.Search(pFilter, False)
        pJFeat = pFC.NextFeature
        If Not pJFeat Is Nothing Then
            '        index = pJFeat.Fields.FindField("Scen_Node")
            '        If pJFeat.value(index) = onode Then

            found = True
            If fVerboseLog Then WriteLogLine("its the from node " & pPline.Length & " splitm " & tempSplitM)

            pPolyCurve.GetSubcurve(0, tempSplitM, bratio, pSplit2)
            '            Set pGeoColl = New GeometryBag
            '            pGeoColl.AddGeometry pSplit2    'pSplit2 is the short part
            pPolyCurve.GetSubcurve(tempSplitM, pPolyCurve.Length, False, pSplit1)
            '            pGeoColl.AddGeometry pSplit1    'pSplit1 is the long part
            NewPoint = pSplit2.ToPoint

            Select Case pFeature.Value(pFeature.Fields.FindField("UseEmmeN"))
                Case Nothing, 2, 0
                    If pFeature.Value(pFeature.Fields.FindField(g_INode)) = onode - m_Offset Then
                        bUpdateJ = True
                    Else
                        bUpdateJ = False
                    End If
                Case 1, 3   'reserved emme2 node
                    If pFeature.Value(pFeature.Fields.FindField(g_INode)) = onode Then
                        bUpdateJ = True
                    Else
                        bUpdateJ = False
                    End If
            End Select
            '        End If
        End If

        If found = False Then 'check from node
            '        Set pFilter = New SpatialFilter
            pSegCol = New Polygon
            With pFilter
                'commented code to make it run
                '.Geometry = pSegCol.SetCircle(ToPt, 0.5)
                '            .GeometryField = m_junctShp.ShapeFieldName
                '            .SpatialRel = esriSpatialRelIntersects
            End With
            '        If pFCls.featurecount(pFilter) > 0 Then
            pFC = pFCls.Search(pFilter, False)
            pJFeat = pFC.NextFeature
            If Not pJFeat Is Nothing Then
                '            index = pJFeat.Fields.FindField("Scen_Node")
                '            If pJFeat.value(index) = onode Then
                '                tempSplitM = tempSplitM
                tempSplitM = pPline.Length - tempSplitM
                found = True
                If fVerboseLog Then WriteLogLine("its the to node" & pPline.Length & "splitm" & tempSplitM)

                pPolyCurve.GetSubcurve(0, tempSplitM, bratio, pSplit1)
                '                Set pGeoColl = New GeometryBag
                '                pGeoColl.AddGeometry pSplit2    'pSplit1 is the short part
                pPolyCurve.GetSubcurve(tempSplitM, pPolyCurve.Length, False, pSplit2)
                '                pGeoColl.AddGeometry pSplit1    'pSplit2 is the long part

                NewPoint = pSplit1.ToPoint

                Select Case pFeature.Value(pFeature.Fields.FindField("UseEmmeN"))
                    Case Nothing, 1, 0
                        If pFeature.Value(pFeature.Fields.FindField(g_JNode)) = onode - m_Offset Then
                            bUpdateJ = False
                        Else
                            bUpdateJ = True
                        End If
                    Case 2, 3   'reserved emme2 node
                        If pFeature.Value(pFeature.Fields.FindField(g_JNode)) = onode Then
                            bUpdateJ = False
                        Else
                            bUpdateJ = True
                        End If
                End Select
                '            End If
            Else
                WriteLogLine("Split feature oid " & pFeature.OID & "FAILED. It doesn't coincide with Scen_Node " & onode)
                SplitM = False
                Exit Sub
            End If
        End If

        SplitM = True
        bpart = True 'so create new part

        '    Dim originalI As Long
        '    Dim originalJ As Long
        '    index = pfeature.Fields.FindField("INode")
        '    originalI = pfeature.value(index)
        '    index = pfeature.Fields.FindField("JNode")
        '    originalJ = pfeature.value(index)
        '    index = pfeature.Fields.FindField("UseEmmeN")
        '    If IsNull(pfeature.value(index)) Or pfeature.value(index) = 0 Then
        '        originalI = originalI + m_Offset
        '        originalJ = originalJ + m_Offset
        '
        '    ElseIf pfeature.value(index) = 1 Then
        '        originalI = dctEmme2Node.Item(CStr(originalI))
        '        originalJ = originalJ + m_Offset
        '    ElseIf pfeature.value(index) = 2 Then
        '        originalJ = dctEmme2Node.Item(CStr(originalJ))
        '        originalI = originalI + m_Offset
        '    Else
        '        originalI = dctEmme2Node.Item(CStr(originalI))
        '        originalJ = dctEmme2Node.Item(CStr(originalJ))
        '    End If
        '
        '    orgstr = CStr(originalI) + " " + CStr(originalJ)
        '    orgstr2 = CStr(originalJ) + " " + CStr(originalI)

        Dim i As Integer
        For i = 0 To 1
            If i = 0 Then
                'has to create a new feature first in order to copy the original IJ nodes.
                pNewFeature = m_edgeShp.CreateFeature
                pNewFeature.Shape = pSplit2 'Remember pSplit2 is short portion, and bUpdateJ indicates the weave portion (short portion)
                CopyAttributes(pFeature, pNewFeature, "")
                If bUpdateJ Then
                    pFeature.Value(pFeature.Fields.FindField(g_JNode)) = SplitID - m_Offset
                Else
                    pFeature.Value(pFeature.Fields.FindField(g_INode)) = SplitID - m_Offset
                End If
                index = pNewFeature.Fields.FindField("Split" + lType)
                pNewFeature.Value(index) = lType
                pNewFeature.Store()
            Else
                pFeature.Shape = pSplit1    'Remember pSplit1 is long portion, and bUpdateJ indicates the weave portion (short portion)
                If bUpdateJ Then
                    pFeature.Value(pFeature.Fields.FindField(g_INode)) = SplitID - m_Offset
                Else
                    pFeature.Value(pFeature.Fields.FindField(g_JNode)) = SplitID - m_Offset
                End If
                pFeature.Store()
                pNewFeature = pFeature
            End If

            '        Set pFilter = New SpatialFilter
            '        With pFilter
            '            Set .Geometry = pNewFeature.Shape
            '            .GeometryField = m_junctShp.ShapeFieldName
            '            .SpatialRel = esriSpatialRelIntersects
            '        End With
            '
            '        If m_junctShp.featurecount(pFilter) > 0 Then 'should get one node- the new one is not in the intermediate junct shp
            '
            '        Set pFC = m_junctShp.Search(pFilter, False)
            '        Set pJFeat = pFC.NextFeature
            '        Do Until pJFeat Is Nothing
            '        Dim nodecnt As Long
            '        nodecnt = 0
            '        Dim temp As Long
            '        'keep original direction
            '        indexJ = pJFeat.Fields.FindField("Scen_Node")
            '        If (IsNull(pJFeat.value(indexJ)) = False) Then
            '
            '            If ((originalI + m_Offset) = pJFeat.value(indexJ)) Then
            '                index = pNewFeature.Fields.FindField("INode")
            '                temp = pJFeat.value(indexJ) - m_Offset
            '                pNewFeature.value(index) = temp
            '                index = pNewFeature.Fields.FindField("JNode")
            '
            '                pNewFeature.value(index) = SplitID - m_Offset
            '
            '                temp = pJFeat.value(indexJ)
            '                If i = 0 Then
            '                   replstr = CStr(temp) + " " + CStr(SplitID)
            '                   replstr2 = CStr(SplitID) + " " + CStr(temp)
            '                Else
            '
            '                  index = pNewFeature.Fields.FindField("Split" + lType)
            '                  pNewFeature.value(index) = lType
            '                End If
            '                pNewFeature.Store
            '
            '            ElseIf ((originalJ + m_Offset) = pJFeat.value(indexJ)) Then
            '                index = pNewFeature.Fields.FindField("JNode")
            '                temp = pJFeat.value(indexJ) - m_Offset
            '                pNewFeature.value(index) = temp
            '                index = pNewFeature.Fields.FindField("INode")
            '                pNewFeature.value(index) = SplitID - m_Offset
            '
            '                temp = pJFeat.value(indexJ)
            '                If i = 0 Then
            '                    replstr = CStr(SplitID) + " " + CStr(temp)
            '                    replstr2 = CStr(temp) + " " + CStr(SplitID)
            '                Else
            '                    index = pNewFeature.Fields.FindField("Split" + lType)
            '                    pNewFeature.value(index) = lType
            '                End If
            '                pNewFeature.Store
            '            End If
            '        Else 'its Null
            '            WriteLogLine "DATA ERROR: PSRCJunctionID is null in SplitFeatureByM"
            '        End If
            '        Set pPline = pNewFeature.ShapeCopy
            '
            '        If pPline.FromPoint.x = FromPt.x And pPline.FromPoint.Y = FromPt.Y Then 'need to get point so can get geomtry
            '            Set NewPoint = pPline.ToPoint
            '        Else
            '            Set NewPoint = pPline.FromPoint
            '        End If
            '        Set pJFeat = pFC.NextFeature
            '    Loop
            '    pNewFeature.Store

            index = pNewFeature.Fields.FindField(pNewFeature.Class.OIDFieldName)
            spID = pNewFeature.Value(index)

            If dctEdges.Exists(CStr(spID)) Then
                dctEdges.Item(CStr(spID)) = pNewFeature
            Else
                dctEdges.Add(CStr(spID), pNewFeature)
            End If

            'SEC-commented out below to make code work
            'End If
            pJFeat = Nothing
        Next i

        Dim pNewJct As IFeature
        pNewJct = m_junctShp.CreateFeature
        pNewJct.Shape = NewPoint
        pNewJct.Value(pNewJct.Fields.FindField(g_PSRCJctID)) = LargestJunct
        pNewJct.Value(pNewJct.Fields.FindField("Scen_Node")) = LargestJunct + m_Offset
        pNewJct.Store()
        dctJct.Add(LargestJunct + m_Offset, pNewJct)
        dctWeaveNodes.Add(CStr(LargestJunct + m_Offset), LargestJunct)
        dctWeaveNodes2.Add(onode, LargestJunct + m_Offset)

        If nodeString = "" Then
            nodeString = "a " + CStr(LargestJunct + m_Offset) + " " + CStr(NewPoint.X) + " " + CStr(NewPoint.Y) + " 0 0 0 0"
        Else
            nodeString = nodeString + vbCrLf + "a " + CStr(LargestJunct + m_Offset) + " " + CStr(NewPoint.X) + " " + CStr(NewPoint.Y) + " 0 0 0 0"
        End If

        Exit Sub

ErrChk:
        CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.SplitFeatureByM")
        MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.SplitFeatureByM")
        'Print #9, Err.Description, vbInformation, "SplitFeatures"

    End Sub



    Private Function CanSplit(ByVal pFeat As IFeature, ByVal pFeatInt As IFeature, ByVal pAttRow As IRow, ByVal lType As String) As Boolean
        Dim fldFacType As Long, fldModes As Long
        Dim index As Long
        Dim pQFt As IQueryFilter
        Dim pTable As ITable, laneFld As String
        Dim pTC As ICursor
        Dim pRow As IRow, iLanes As Double
        Dim timePd(4) As String
        Dim i As Integer
        timePd(0) = "AM"
        timePd(1) = "MD"
        timePd(2) = "PM"
        timePd(3) = "EV"
        timePd(4) = "NI"

        CanSplit = False
        If Not (IsdbNull(pFeatInt.value(pFeatInt.Fields.FindField("SPLIT" & lType)))) Then Exit Function

        fldFacType = pFeat.Fields.FindField("FacilityType")
        fldModes = pFeat.Fields.FindField("Modes")

        If pFeat.value(fldFacType) >= 6 And pFeatInt.value(fldFacType) <> 1 And pFeatInt.value(fldFacType) <> 5 Then
            Select Case pFeatInt.value(fldModes)
                Case "wk", "r", "f"
                    CanSplit = False
                    Exit Function
                Case Else
            End Select

            index = pFeat.Fields.FindField(g_PSRCEdgeID)
            '        Set pTable = get_TableClass(m_layers(2)) 'need to put in check to see if need to pull out of tblLine or evt tables
            '        Set pQFt = New QueryFilter
            '        pQFt.WhereClause = g_PSRCEdgeID & " = " + CStr(pFeatInt.value(index))
            '        Set pTC = pTable.Search(pQFt, False)
            '        Set pRow = pTC.NextRow
            pRow = pAttRow
            If lType = "HOV" Or lType = "GP" Then
                For i = 0 To 4
                    laneFld = CStr(lType) + CStr(timePd(i))
                    index = pRow.Fields.FindField("IJLanes" + laneFld)
                    iLanes = iLanes + pRow.value(index)
                    index = pRow.Fields.FindField("JILanes" + laneFld)
                    iLanes = iLanes + pRow.value(index)
                Next i
            Else
                laneFld = CStr(lType)
                index = pRow.Fields.FindField("IJLanes" + laneFld)
                iLanes = pRow.value(index)
                index = pRow.Fields.FindField("JILanes" + laneFld)
                iLanes = iLanes + pRow.value(index)
            End If

            If iLanes = 0 Then CanSplit = True
        End If
    End Function
    Public Sub create_NetFile2(ByVal pathnameN As String, ByVal filenameN As String, ByVal pWS As IWorkspace)
        'on error GoTo eh
        'called by GlobalMod.create_ScenarioShapefiles
        'pathnameN and filenameN input by user in frmNetLayer
        'm_edgeShp and m_junctShp are the intermediate shapefiles used to create the Netfile
        '    Dim strFFT As String    'holds formatted FFT number
        '    strFFT = ""

        '    Dim pathname As String
        '    pathname = "c:\createNet_ERROR.txt"
        '    Open pathname For Output As #9

        WriteLogLine("")
        WriteLogLine("========================================")
        WriteLogLine("create_NetFile2 started " & Now())
        WriteLogLine("Model year " & GlobalMod.inserviceyear)
        WriteLogLine("========================================")
        WriteLogLine("")

        'SEC 031711, commenting out status bar for now.....
        'Dim pStatusBar As IStatusBar
        'Dim pProgbar As IStepProgressor
        'pStatusBar = m_App.StatusBar
        'pProgbar = pStatusBar.ProgressBar

        'Dim pathnameAm As String, pathnameM As String, pathnamePm As String, pathnameE As String, pathnameNi As String
        Dim pathTOD(0 To 4) As String
        Dim attribname As String
        Dim lType, timePd As ArrayList
        lType = New ArrayList
        timePd = New ArrayList
        Dim mydate As Date
        Dim t As Integer, l As Integer, i As Integer

        lType.Add("GP")
        lType.Add("TR")
        lType.Add("HOV")
        lType.Add("TK")
        timePd.Add("AM")
        timePd.Add("MD")
        timePd.Add("PM")
        timePd.Add("EV")
        timePd.Add("NI")
        mydate = Date.Now


        For t = 0 To 4
            pathTOD(t) = pathnameN + "\" + filenameN + CStr(timePd(t)) + ".txt"
            FileClose(t + 1)
            FileOpen(t + 1, pathTOD(t), OpenMode.Output, OpenAccess.Default, OpenShare.Shared)
            PrintLine((t + 1), "c Exported AM Network File from ArcMap /Emme/2 Interface: " + CStr(mydate))
            PrintLine((t + 1), "t nodes init")
        Next t
        attribname = pathnameN + "\" + filenameN + "Attr.txt"

        FileOpen(6, attribname, OpenMode.Output)


        'open temp files
        PrintLine(6, "c Exported Link Attribute File from ArcMap /Emme/2 Interface: " + CStr(mydate))

        Dim pWorkspace As IWorkspace
        pWorkspace = g_FWS
        Dim pTblPrjEdge As ITable
        Dim pTblMode As ITable
        Dim pEvtLine As ITable
        Dim pevtPoint As ITable

        Dim wString As String, pstring As String, nodes As String, astring As String
        Dim tempString As String

        wString = ""
        nodes = ""

        'get tables needed for the attribute data need to send to emme/2
        '[073007] hyu: pTblLine is substituted w/ pTblPrjEdge, and set to tblScenarioProject
        '[062807] sec: pTblLine is being set to tblLineProjects. Edges are given Attributes from this if they are
        'from a project. The IJ node information is contained in the ScenarioEdge Shapefile, which takes the
        'IJ nodes from the underlying network TransRefEdges, not ProjectRoutes (this is correct). Since pTbLine does not
        'have any IJ node information, I suspect that there is no check to make sure that the IJ node are the same in the
        'ScenarioEdge and the ProjectRoutes

        pTblPrjEdge = get_TableClass(m_layers(24)) 'use if future project
        pTblMode = get_TableClass(m_layers(2))
        pEvtLine = get_TableClass(m_layers(4)) 'use if future event project
        pevtPoint = get_TableClass(m_layers(6)) 'use if future event project

        Dim pTC As ICursor, pTCt As ICursor
        Dim pRow As IRow
        Dim tempAttrib As String
        Dim pFeatCursor As IFeatureCursor
        Dim pFeat As IFeature
        Dim pFlds As IFields
        Dim lSFld As Long, upIndex As Long, tempIndex As Long
        Dim Tindex As Long
        Dim pWSedit As IWorkspaceEdit
        pWSedit = pWS

        Dim count As Long
        Dim frmNetLayer As New frmNetLayer
        count = m_EdgeSSet.Count '+ m_JunctSSet.count
        'pProgbar.position = 0
        'pStatusBar.ShowProgressBar("Creating NetFile...", 0, count, 1, True)
        'jaf--let's keep the user updated
        frmNetLayer.lblStatus.Text = "GlobalMod.create_Netfile: Starting Netfile Build..."
        'pan frmNetLayer.Refresh

        GetLargestID(m_junctShp, "PSRCJunctID", LargestJunct)

        pFeatCursor = m_junctShp.Search(Nothing, False)

        Dim test As Boolean
        Dim pPoint As IPoint
        Dim pMRow As IRow

        test = False
        Dim j As Long, shpindex As Long
        Dim tempprj As ClassPrjSelect
        Dim HaveSplits As Boolean
        Dim Subtype As Long
        Dim Emme2NodeID As Long

        HaveSplits = False
        pFeat = pFeatCursor.NextFeature

        lSFld = pFeat.Fields.FindField("Scen_Node") 'this id used by emme2
        upIndex = pFeat.Fields.FindField("Updated1")
        shpindex = pFeat.Fields.FindField("shptype")
        Subtype = pFeat.Fields.FindField("JunctionType")
        'sec082108 added this variable to make sure park and rides have an Emme2 Node ID greater than 0
        Emme2NodeID = pFeat.Fields.FindField("EMME2nodeID")

        'jaf--let's keep the user updated
        frmNetLayer.lblStatus.Text = "GlobalMod.create_Netfile: Starting Junction portion of NetFile..."
        'pan frmNetLayer.Refresh

        Dim pcount As Long
        Dim dctReservedNodes As Dictionary  '[042506] hyu: dictionary for reserved nodes: key=PSRCJunctId, item=Emme2Node
        Dim dctWeaveNodes(3) As clsCollection       '[051106] hyu: dictionary for weave nodes not physically in the ScenarioJunct layer
        Dim dctWeaveNodes2(3) As clsCollection    '[052007] hyu: dictionary for junctions and corresponding weave nodes

        Dim lPSRCJctID As Long
        dctReservedNodes = New Dictionary



        For l = 0 To 3

            dctWeaveNodes(l) = New clsCollection
            dctWeaveNodes2(l) = New clsCollection
        Next l

        Dim dctJcts As Dictionary, dctJoints As Dictionary, dctEdgeID As Dictionary, lMaxEdgeOID As Long
        Dim pNewFeat As IFeature
        dctJcts = New Dictionary

        WriteLogLine("start getting all nodes " & Now())
        pcount = 1
        '[062007] jaf: populate the junctions dictionary
        Try


            Do Until pFeat Is Nothing 'loop through features in intermediate junction shapefile
                pPoint = pFeat.Shape
                '[042706] pan:  added parkandride node subtype 7 to test-they need a * also
                '[090408] pan:  there are junctiontype parkandride subtype 7 that are not currently modeled and
                '               would bump us over 1200 limit if used, so check EMME2nodeid > 0
                If (pFeat.Value(Subtype) = 6) Or ((pFeat.Value(Subtype) = 7 And pFeat.Value(Emme2NodeID) > 0)) Then
                    'jaf--centroid junctions (subtype 6) are special cases
                    tempString = "a* " + CStr(pFeat.Value(lSFld)) + " " + CStr(pPoint.X) + " " + CStr(pPoint.Y) + " 0 0 0 0" '+ CStr(pFeat.value(lSFld))
                Else
                    tempString = "a " + CStr(pFeat.Value(lSFld)) + " " + CStr(pPoint.X) + " " + CStr(pPoint.Y) + " 0 0 0 0" '+ CStr(pFeat.value(lSFld))
                End If
                pcount = pcount + 1

                '[042506] hyu: populate lookup dictionary of reserved nodes for Emme2
                lPSRCJctID = pFeat.Value(pFeat.Fields.FindField(g_PSRCJctID))

                '[050407] hyu: recording the PSRCJunctID and Scen_Node correspondance
                dctJcts.Add(pFeat.Value(lSFld), pFeat)

                If lPSRCJctID + m_Offset <> pFeat.Value(lSFld) Then
                    If Not dctReservedNodes.Exists(CStr(lPSRCJctID)) Then dctReservedNodes.Add(CStr(lPSRCJctID), CStr(pFeat.Value(lSFld)))
                End If

                '************************************
                'jaf--removed block 1 for readability.  see removed_block_1.txt
                '************************************
                For t = 1 To 5
                    PrintLine(t, tempString)
                Next t

                'pStatusBar.StepProgressBar
                pFeat = pFeatCursor.NextFeature
            Loop 'end of loop through every feature in intermediate junction shapefile
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        'jaf--let's keep the user updated
        frmNetLayer.lblStatus.Text = "GlobalMod.create_Netfile: Finished creating Junction portion of Netfile"
        'pan frmNetLayer.Refresh

        Dim NextLine As String
        Dim lanecap As Double, length As Double
        Dim oneIndex As Long, cnt As Long
        oneIndex = m_edgeShp.FindField("Oneway")

        'jaf--this appears to commence the links portion of the buildfiles
        Dim tempstringAm As String, tempstringM As String, tempstringPm As String, tempstringE As String, tempstringN As String
        Dim strTOD(0 To 4) As String

        '       pathname = "c:\MODULEcreateweave.txt"
        '       Open pathname For Append As #2
        WriteLogLine("Model Input create weave links ")

        'jaf--let's keep the user updated
        frmNetLayer.lblStatus.Text = "GlobalMod.create_Netfile: Starting Edge portion of NetFile..."
        'pan frmNetLayer.Refresh

        '[051507]hyu: no need of cursor
        upIndex = m_edgeShp.FindField("Updated1")
        lSFld = m_edgeShp.FindField("PSRCEdgeID") 'this used to get out of tblmode
        shpindex = m_edgeShp.FindField("shptype")

        Dim Wprint As Boolean
        Dim intWays As Integer
        Dim uindex As Integer
        Dim lanes As Double
        Dim direction As String
        Dim lflag As String
        Dim mode As String, linkType As String
        Dim laneFld As String
        Dim pos As Long
        Dim Leng
        Dim newstr As String
        Dim rString As String
        Dim rLocation As Long
        Dim sVDF As String, sCap As String, sFFS As String, sFacType As String, sNewFacilityType As String, sOldFacilityType As String
        Dim lancapfld As String
        Dim intFacType As Integer
        Dim dctWeaveLinks As Dictionary
        Dim sDir() As String
        Dim sLeng As String

        dctWeaveLinks = New Dictionary

        '[033106] hyu: declare a lookup dictionary for edges in the ModeAttributes table.
        '              This dictionary replaces the query for row count in the ModeAttributes table.
        Dim dctMAtt As Dictionary, dctEdges As Dictionary
        Dim bEOF As Boolean, lEdgeCt As Long, sEdgeOID As String, sEdgeID As String
        Dim dctSplit As Dictionary
        dctSplit = New Dictionary

        WriteLogLine("start getting all edges and mode attributes " & Now())
        '[062007] jaf: appears to populate dctEdges with intermed. edge features...
        '              ...and dctMAtt with related modeAttributes records
        '              pTblMode is handle to ModeAttributes
        '              m_edgeShp is handle to ScenarioEdges
        dctEdges = New Dictionary
        dctMAtt = New Dictionary
        getMAtts(m_edgeShp, pTblMode, g_PSRCEdgeID, dctEdges, dctMAtt)

        WriteLogLine("start writing links " & Now())
        pWSedit.StartEditing(False)
        pWSedit.StartEditOperation()

        bEOF = False
        lEdgeCt = 0

        If dctEdges.Count = 0 Then Exit Sub

        '************************************************
        '[062007] jaf: I THINK this only adds Weave links (and related splits) where warranted...
        '              ...by looping through all edges in dctEdges.
        Try


            Do Until bEOF
                pFeat = dctEdges.Items(lEdgeCt)
                If dctMAtt.Exists(CStr(pFeat.Value(pFeat.Fields.FindField(g_PSRCEdgeID)))) Then
                    pMRow = dctMAtt.Item(CStr(pFeat.Value(pFeat.Fields.FindField(g_PSRCEdgeID))))
                    'getAttributes row returns proper attributes for the edge feature based on Updated1 field
                    '...from related ModeAttributes, tblLineProjects, or evtLineProjects records (as row obj)
                    pRow = getAttributesRow(pFeat, pMRow, pTblPrjEdge, pEvtLine)

                    '************************************************
                    'PROBLEM:  should check TK also!!!
                    '************************************************

                    '[062007] jaf: checks lane types 1, 2 (TR, HOV) to see if need weave; SHOULD check 3 (TK) ALSO!!
                    '[092710] SEC: All combinations of HOV/TR are handled in the HOV fields (coded values). No need to go check TR fields any longer
                    l = 2
                    If getAllLanes(pRow, pMRow, CStr(lType(l))) > 0 Then
                        If pRow.Value(pRow.Fields.FindField("IJLanesGPAM")) > 0 Or pRow.Value(pRow.Fields.FindField("JILanesGPAM")) > 0 Or pRow.Value(pMRow.Fields.FindField("IJLanesGPAM")) > 0 Or pMRow.Value(pRow.Fields.FindField("JILanesGPAM")) > 0 Then

                            createWeaveLink(CStr(lType(l)), pFeat, pRow, wString, nodes, astring, dctReservedNodes, dctWeaveNodes(l).myDictionary, dctWeaveNodes2(l).myDictionary, dctJcts, dctEdges, pstring)
                        End If
                    End If

                End If
                lEdgeCt = lEdgeCt + 1
                If lEdgeCt = dctEdges.Count Then bEOF = True
                WriteLogLine("edge count=" & lEdgeCt & " " & Now())
                '        If pFeat.OID = 1508230 Then MsgBox 1
                'If lEdgeCt = 15000 Then MsgBox lEdgeCt
                getSplitLink(pFeat, dctSplit, dctReservedNodes)
            Loop

            For t = 0 To 4
                PrintLine(t + 1, pstring)

                PrintLine(t + 1)
                PrintLine(t + 1, "t links init")
            Next t

            pWSedit.StopEditOperation()
            pWSedit.StopEditing(True)
            '    pWSEdit.StartEditing False
            '    pWSEdit.StartEditOperation


            Dim wNodes As String
            Dim wasHOV As Boolean

            '************************************************
            ' [062007] jaf: dictionaries have been prepped; write out actual buildfile

            lEdgeCt = 0
            bEOF = False
            Do Until bEOF

                pFeat = dctEdges.Items(lEdgeCt)

                lEdgeCt = lEdgeCt + 1
                If lEdgeCt = dctEdges.Count Then bEOF = True
                WriteLogLine("edge count=" & lEdgeCt & " " & Now())
                SplitM = False
                'MsgBox "pfeat oid " + CStr(pFeat.OID)
                'jaf--create the basic links for one-way or two-way facilities based on field Oneway:
                '       0 = one way from I
                '       1 = one way from J
                '       2 = two way

                intWays = 2
                If IsDBNull(pFeat.Value(oneIndex)) Then intWays = 2 Else intWays = pFeat.Value(oneIndex)
                '[033006] hyu: write wrong value for oneway
                'If fVerboseLog Then WriteLogLine "net oneway" + CStr(oneIndex)
                If fVerboseLog Then WriteLogLine("net oneway" + CStr(intWays))
                Select Case intWays
                    Case 0
                        'one way from I
                        ReDim sDir(0)
                        sDir(0) = "IJ"
                    Case 1
                        'one way from J
                        ReDim sDir(0)
                        sDir(0) = "JI"
                    Case 2, 3
                        'two way
                        ReDim sDir(1)
                        sDir(0) = "IJ"
                        sDir(1) = "JI"
                End Select

                uindex = m_edgeShp.FindField("UseEmmeN")
                If fVerboseLog Then WriteLogLine("useemmen is index " & CStr(uindex))
                For i = LBound(sDir) To UBound(sDir)
                    tempString = ""
                    nodes = ""
                    'jaf--this block writes the inode and jnode values in appropriate direction
                    'jaf--(columns inode and jnode)


                    nodes = getIJNodes(pFeat, sDir(i), m_Offset, dctReservedNodes)

                    'jaf--this block writes the length in miles--6 ASCII CHARS MAX LEN!
                    'jaf--(column len(miles))

                    If fVerboseLog Then WriteLogLine("nodes string=" & nodes)

                    sLeng = getLength(strLayerPrefix, pFeat)

                    'jaf--lane type is derived from projects, if any
                    'jaf--assume this all works for now
                    '[033006] hyu: get correct link attributes as following
                    'a real link from TransRefEdges: modeAttributes (pTC)
                    'a project link: pTblPrjEdge
                    'a event link: pEvtLine
                    Tindex = pFeat.Fields.FindField("Scen_Link") 'this links the attribute file (which gets included in the punch) back tot he intermediate shp
                    tempAttrib = nodes + " " + CStr(pFeat.Value(Tindex))

                    '[033106] hyu: change to lookup a dictionary before doing the query
                    If Not dctMAtt.Exists(CStr(pFeat.Value(lSFld))) Then
                        'jaf--modified to give more info
                        WriteLogLine("DATA ERROR: Edge " & CStr(pFeat.Value(lSFld)) & " is missing in ModeAttributes--GlobalMod.create_NetFile")
                        'move to next one
                    Else
                        '[040406] hyu: change to lookup a dictionary
                        pMRow = dctMAtt.Item(CStr(pFeat.Value(lSFld)))
                        pRow = getAttributesRow(pFeat, pMRow, pTblPrjEdge, pEvtLine)
                        If fVerboseLog Then WriteLogLine("setting direction")

                        For l = 0 To 2 'for now just do gp and hov and tr
                            'wasHOV = False
                            Wprint = False

                            lflag = getFlag(CStr(lType(l)), sDir(i))
                            mode = getModes(pFeat, CStr(lType(l)))
                            linkType = getLinkType(pFeat)

                            'jaf--all times of day buildfiles get same info up to this point
                            'assign so far a iNode, jNode, length, modes, linkType if not hov
                            If l = 0 Then
                                tempString = nodes + tempString + " " + sLeng
                                nodes = ""
                            End If

                            wString = ""
                            astring = ""

                            For t = 0 To 4
                                'NOW assign: a iNode, jNode, length, modes, linkType, #lane

                                strTOD(t) = tempString
                                lanes = getLanes(pRow, pMRow, CStr(lType(l)), sDir(i), CStr(timePd(t)))
                                If fVerboseLog Then WriteLogLine("checking lanes for " & sDir(i) & "=" & lanes)
                                '***********this still needs to be FIXED! Not a big deal right now, but will be if we have HOV and BAT lanes
                                'that are variable in time and space. E.G- present in AM, not in MD, present again in PM
                                '*************RemoveWeaveLinks is my first attempt at a sub to deal with this
                                'SEC:112410 We need to check to see if the nodes have been added to
                                'the weave dictionary from other time periods when there were HOV/Managed Lanes but not during the current time period
                                '. If so, we need to remove them so that if there are HOV/Managed lanes
                                'subsequent time periods, all the necessary weave links are built!!!
                                'If t > 0 And t < 3 And lanes = 0 And l = 2 And wasHOV = True Then
                                'might have been HOV in an earlier time period. Need to remove weavelinks from dictionary in case it is HOV again in subsequent time periods.
                                'RemoveWeaveLinks pFeat, pRow, pMRow, dctWeaveLinks, dctSplit, dctReservedNodes, CStr(lType(l)), sDir(i), CStr(timePd(t)), lanes, nodes, astring, wString

                                'End If


                                If lanes > 0 Then


                                    If l = 0 Then 'GP
                                        strTOD(t) = "a " + strTOD(t) + " " + mode + " " + linkType + " " + CStr(lanes)
                                    Else 'l>0 non-GP
                                        Wprint = True
                                        If fVerboseLog Then WriteLogLine("go to weave")

                                        nodes = ""
                                        'sec added 4/21/09
                                        If pRow.Value(pRow.Fields.FindField("IJLanesGPAM")) > 0 Or pRow.Value(pRow.Fields.FindField("JILanesGPAM")) > 0 Or pRow.Value(pMRow.Fields.FindField("IJLanesGPAM")) > 0 Or pMRow.Value(pRow.Fields.FindField("JILanesGPAM")) > 0 Then
                                            GetWeaveLinks(pFeat, pRow, pMRow, dctWeaveLinks, dctSplit, dctReservedNodes, CStr(lType(l)), sDir(i), CStr(timePd(t)), lanes, nodes, astring, wString)
                                            '                                length = getWeaveLen(CStr(lType(l))) / 5280 'convert feet to mile
                                        End If

                                        If l <> 2 Then 'not HOV
                                            lanes = 1
                                            strTOD(t) = "a " + nodes + " " + sLeng + " " + mode + " " + linkType + " " + CStr(lanes)

                                            '*************TO DO*********This is where we deal with HOV/TR/BAT/HOT/TR lanes- Needs to be coded to handle all possibilities
                                        Else    'hov modes depends on the lanes
                                            'wasHOV = True
                                            Select Case lanes
                                                Case 1
                                                    mode = "ahijb"
                                                Case 2
                                                    mode = "aijb"
                                                    'TR or BAT
                                                Case 3
                                                    mode = "b"
                                                Case 4
                                                    mode = "b"
                                            End Select
                                            lanes = 1
                                            strTOD(t) = "a " + nodes + " " + sLeng + " " + mode + " " + linkType + " " + CStr(lanes)
                                        End If

                                        If t = 0 Then
                                            If astring = "" Then
                                                astring = nodes + " " + CStr(pFeat.Value(Tindex)) + " " + lflag
                                            Else
                                                astring = astring + vbNewLine + nodes + " " + CStr(pFeat.Value(Tindex)) + " " + lflag
                                            End If
                                            PrintLine(6, astring)

                                            If fVerboseLog Then WriteLogLine("ASTRING* " + astring)
                                        End If
                                    End If

                                    If fVerboseLog Then WriteLogLine("checking VDFunc")

                                    'NOW assign: a iNode, jNode, length, modes, linkType, #lane, functional class
                                    sNewFacilityType = getNewFacilityType(pFeat)
                                    sVDF = getVDF(pRow, pMRow, t, sDir(i), mode, linkType, sNewFacilityType)
                                    sCap = getLaneCap(pRow, pMRow, sDir(i), CStr(lType(l)))

                                    sOldFacilityType = getOldFacilityType(pFeat)
                                    sFFS = getSpeedLimit(pRow, pMRow, CStr(lType(l)), sDir(i), sOldFacilityType, pFeat)
                                    sFacType = getNewFacilityType(pFeat)
                                    strTOD(t) = strTOD(t) + " " + sVDF + " " + sCap + " " + sFFS + " " + sFacType
                                    'SEC--took out SFFS string (this attribute no longer exists in ModeAttributes
                                    'strTOD(t) = strTOD(t) + " " + sVDF + " " + sCap + " " + sFacType
                                    If Not wString = "" Then
                                        If fVerboseLog Then WriteLogLine("wstring=" & wString)
                                        PrintLine(t + 1, wString)
                                    End If

                                    'SEC 010809: changed the following code so that reversibles run in both directions in the mid-day
                                    If l = 0 Then
                                        If intWays = 3 And t <> 1 Then 'reversible
                                            If i = 0 Then   'IJ
                                                If t = 0 Then PrintLine(t + 1, strTOD(t)) 'AM, MD
                                            Else
                                                If t > 1 Then PrintLine(t + 1, strTOD(t)) 'PM, EV, NI
                                            End If
                                        Else
                                            PrintLine(t + 1, strTOD(t))
                                        End If
                                    ElseIf strTOD(t) <> "" Then
                                        PrintLine(t + 1, strTOD(t))

                                    End If
                                End If

                                '                        tempString = ""
                            Next t
                            tempString = ""
                            'If l = 0 Or Wprint = True Then
                            If l = 0 Then
                                PrintLine(6, tempAttrib + " " + lflag)
                            End If
                            '                    tempString = ""
                        Next l
                    End If
                Next i

                'pStatusBar.StepProgressBar()
            Loop 'end of loop through every feature in intermediate edge shapefile
            'now add an additional links from splitting due to weave link case 2
        Catch ex As Exception
            MessageBox.Show(ex.ToString)

        End Try
        For t = 0 To 5
            FileClose(t + 1)
        Next t

        '    pWSEdit.StopEditOperation
        '    pWSEdit.StopEditing True
        '
        'now copy temp ouput to built file
        WriteLogLine("copying temp output files to final build files")
        '    CreateFinalNetfiles pathnameN, filenameN

        'jaf--let's keep the user updated
        frmNetLayer.lblStatus.Text = "GlobalMod.create_Netfile: Finished creating Edge portion of Netfile"
        'pan frmNetLayer.Refresh
        'MsgBox "Finished creating Edge portion of Netfile", , "GlobalMod.create_Netfile"
        'pStatusBar.HideProgressBar()

        pFeat = Nothing
        pWorkspace = Nothing
        pTblPrjEdge = Nothing
        'Set ptblPoint = Nothing
        pTblMode = Nothing
        pEvtLine = Nothing
        pevtPoint = Nothing
        pRow = Nothing

        dctEdges.RemoveAll()
        dctMAtt.RemoveAll()
        dctReservedNodes.RemoveAll()
        dctSplit.RemoveAll()
        For l = 1 To 3
            dctWeaveNodes(l).myDictionary.RemoveAll()
            dctWeaveNodes2(l).myDictionary.RemoveAll()
            dctWeaveNodes(l) = Nothing
            dctWeaveNodes2(l) = Nothing
        Next l
        dctSplit = Nothing
        dctEdges = Nothing
        dctMAtt = Nothing
        dctReservedNodes = Nothing

        WriteLogLine("FINISHED create_NetFile2 at " & Now())
        Exit Sub
eh:
        CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.create_NetFile2")
        MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.create_NetFile2")

        'Print #9, Err.Description, vbExclamation, "GlobalMod.create_NetFile"
        If pWSedit.IsBeingEdited Then
            pWSedit.AbortEditOperation()
            pWSedit.StopEditing(False)
        End If
        For t = 0 To 4
            FileClose(t + 1)
        Next t
        FileClose(6)

        pFeat = Nothing
        pWorkspace = Nothing
        pTblPrjEdge = Nothing
        'Set ptblPoint = Nothing
        pTblMode = Nothing
        pEvtLine = Nothing
        pevtPoint = Nothing
        pRow = Nothing

        If Not dctEdges Is Nothing Then dctEdges.RemoveAll()
        If Not dctMAtt Is Nothing Then dctMAtt.RemoveAll()
        dctReservedNodes.RemoveAll()
        For l = 1 To 3
            dctWeaveNodes(l).myDictionary.RemoveAll()
            dctWeaveNodes2(l).myDictionary.RemoveAll()
            dctReservedNodes(l) = Nothing
            dctWeaveNodes(l) = Nothing
        Next l
        dctEdges = Nothing
        dctMAtt = Nothing


    End Sub
    Private Function getAllLanes(ByVal pRow As IRow, ByVal pMRow As IRow, ByVal sType As String) As Double
        '[062007] jaf:  sums ALL lane field values across both directions and all time periods...
        '               ...for lane type 'sType'

        Dim dLanes As Double
        Dim laneFld As Long

        dLanes = 0
        Dim sDir(0 To 1) As String
        sDir(0) = "IJ"
        sDir(1) = "JI"


        Dim sTpd(0 To 4) As String
        sTpd(0) = "AM"
        sTpd(1) = "MD"
        sTpd(2) = "PM"
        sTpd(3) = "EV"
        sTpd(4) = "NI"

        Dim i As Integer, j As Integer
        If sType = "GP" Or sType = "HOV" Then 'If l < 2 Then
            For i = LBound(sDir) To UBound(sDir)
                For j = LBound(sTpd) To UBound(sTpd)
                    dLanes = dLanes + getLanes(pRow, pMRow, sType, sDir(i), sTpd(j))
                Next j
            Next i
        Else
            'check for TR!!!!!
            For i = 0 To 1
                dLanes = dLanes + getLanes(pRow, pMRow, sType, sDir(i))
            Next i
        End If
        getAllLanes = dLanes
    End Function
    Private Function getLanes(ByVal pRow As IRow, ByVal pMRow As IRow, ByVal sType As String, ByVal direction As String, Optional ByVal tpd As String = "") As Double
        '[062007] jaf:  returns (as double) value of lane field from record pRow (could be from tblLineProjects, etc.) or pMRow (modeAttributes)
        '               for lane type 'sType' in IJ or JI direction 'Direction' for (optional) time period 'tpd'
        '               NOTES:
        '                  a) if both project-driven and baseyear attribs have no lanes, defaults to 1 GP
        '                  b) returns raw field value NOT number of actual lanes!!!!!

        Dim dLanes As Double
        Dim laneFld As String
        Dim Tindex As Long

        dLanes = 0

        If sType = "GP" Or sType = "HOV" Then 'If l < 2 Then
            laneFld = direction + "Lanes" + sType + tpd
        Else
            laneFld = direction + "Lanes" + sType
        End If

        Tindex = pRow.Fields.FindField(laneFld)
        If (IsDBNull(pRow.Value(Tindex)) = False And pRow.Value(Tindex) > 0) Then
            dLanes = pRow.Value(Tindex)
        Else
            Tindex = pMRow.Fields.FindField(laneFld)
            If (IsDBNull(pMRow.Value(Tindex)) = False And pMRow.Value(Tindex) > 0) Then
                dLanes = pMRow.Value(Tindex)
            Else
                If sType = "GP" Then
                    dLanes = 1
                    'If t = 0 Then MsgBox "no lane value " + tempstringAm
                End If
            End If
        End If

        getLanes = dLanes
    End Function
    Private Function getNewFacilityType(ByVal pFeat As IFeature) As String
        Dim Tindex As Long
        Dim intFacType As Integer
        'jaf--write out the Functional Class as model Facility Type
        Dim frmNetLayer As New frmNetLayer
        'SEC 070909- Changed FacilityType to NewFacilityType to reflect Emme2 FacilityType, NewFacilityType is the field
        'temporarily holding Emme1 FT.
        Tindex = pFeat.Fields.FindField("NewFacilityType") 'this is changing to E2FacType
        If Tindex < 0 Then
            'ERROR--No FacilityType field!!!
            intFacType = 0

            'write out error message
            WriteLogLine("DATA ERROR: GlobalMod.create_NetFile: edge pFeat.OID=" & CStr(pFeat.OID) & " has no Functional field, using 0")
            frmNetLayer.lblStatus.text = "GlobalMod.create_NetFile: edge pFeat.OID=" & CStr(pFeat.OID) & " has no Functional field, using 0"
            'pan frmNetLayer.Refresh
        Else
            'we have a FacilityType field in pFeat
            If IsDBNull(pFeat.Value(Tindex)) Then
                'jaf--facility type is null, which is bad, but we'll fix later
                intFacType = 0
            Else
                'jaf--got good facility type
                intFacType = pFeat.Value(Tindex)
            End If
        End If

        getNewFacilityType = CStr(intFacType)
    End Function
    Private Function getIJNodes(ByVal pFeat As IFeature, ByVal IJorJI As String, ByVal lOffset As Long, ByVal dctRNode As Dictionary) As String
        'dctRNode is the dictionary of reserved nodes
        Dim fldI As String
        Dim fldJ As String
        Dim str As String
        Dim nodeI As Long, nodeJ As Long
        fldI = pFeat.Fields.FindField(g_INode)
        fldJ = pFeat.Fields.FindField(g_JNode)
        nodeI = pFeat.value(fldI)
        nodeJ = pFeat.value(fldJ)
        If dctRNode.Exists(CStr(nodeI)) Then
            nodeI = dctRNode.Item(CStr(nodeI))
        Else
            nodeI = nodeI + lOffset
        End If

        If dctRNode.Exists(CStr(nodeJ)) Then
            nodeJ = dctRNode.Item(CStr(nodeJ))
        Else
            nodeJ = nodeJ + lOffset
        End If

        If IJorJI = "IJ" Then
            str = CStr(nodeI) & " " & CStr(nodeJ)
        Else
            str = CStr(nodeJ) & " " & CStr(nodeI)
        End If
        getIJNodes = str
    End Function

    Private Function getModes(ByVal pFeat As IFeature, ByVal sType As String) As String
        'SEC 101610: For some reason, this is coded to take the default modestring for TR weave links. I changed it to assign TR the
        'mode string "b".


        Dim Tindex As Long
        Dim bDefaultModes As Boolean

        Tindex = pFeat.Fields.FindField("Modes")
        bDefaultModes = False

        If IsDBNull(pFeat.Value(Tindex)) Then
            bDefaultModes = True
        ElseIf Trim(pFeat.Value(Tindex)) = "" Then
            bDefaultModes = True
        End If


        If bDefaultModes Then
            'jaf--default value of modes string
            '[051507] hyu default modes
            Select Case sType
                Case "GP"
                    getModes = "ashi"
                Case "TR"  'tr
                    getModes = "ab"
                Case "HOV"  'hov

                Case "TK"  'tk
                    getModes = "auv"
            End Select
        ElseIf sType = "TR" Then
            getModes = "b"

            'supplied value of modes string
        Else
            getModes = CStr(pFeat.value(Tindex))
        End If

    End Function

    Private Function getLength(ByVal strLayerPrefix As String, ByVal pFeat As IFeature) As String
        'jaf--this block writes the modes string, which has defaults if not supplied
        'jaf--(column modes)
        Dim length As Double



        Dim Tindex As Long
        If strLayerPrefix = "SDE" Then
            Tindex = pFeat.Fields.FindField("Shape.len") 'SDE
        Else
            Tindex = pFeat.Fields.FindField("Shape_Length") 'PGDB
        End If

        '[040406] hyu: Length info can be null
        If IsDBNull(pFeat.Value(Tindex)) Then
            length = 0.015


        Else
            length = pFeat.Value(Tindex)


            length = length * 0.00018939393939
            length = FormatNumber(length, 8)
            If length < 0.01 Then
                length = 0.01
            End If


            ' obj = CDec(length)
        End If

        getLength = Left(CStr(length), 10)

    End Function

    Public Function getAttributesRow(ByVal pFeat As IFeature, ByVal pMRow As IRow, ByVal pTblLine As ITable, ByVal pEvtLine As ITable) As IRow
        '[062007] jaf:  determines which attributes row the ScenarioEdge gets in prep for buildfile creation
        '               choices being: ModeAttributes, tblLineProjects, or evtLineProjectsOutcomes

        '**************************************************
        '[062007] jaf
        'PROBLEM: this algorithm assumes ProjectRoute IJ direction matches RefEdge IJ direction!!!
        '**************************************************

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
        If Not IsDBNull(pFeat.Value(upIndex)) Then
            If pFeat.Value(upIndex) = "Yes" Then

                If Not IsDBNull(pFeat.Value(shpindex)) Then
                    If pFeat.Value(shpindex) <> "Event" Then
                        'get future attributes from tblScenarioProjects
                        '        tempIndex = pFeat.Fields.FindField("prjRte")
                        '        tstring = pFeat.value(tempIndex)
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
                    End If
                End If
            Else
                'get future attributes from evtLineProjects
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
        End If 'now set which table to pull the attributes from


        getAttributesRow = pRow
    End Function

    Private Function getFlag(ByVal sType As String, ByVal direction As String) As String
        '    If (i = 0 And fromI = True) Then
        '        direction = "IJ"
        '        If l = 0 Then lflag = " 1 "      ' flag for IJ
        '        If l = 1 Then lflag = " 3 "   ' flag for IJHOV
        '        If l = 2 Then lflag = " 5 "   ' flag for IJTR
        '        'If l = 3 Then lflag =  " 7 "  ' flag for IJHOV
        '    ElseIf (i = 1 Or fromI = False) Then
        '        direction = "JI"
        '        If l = 0 Then lflag = " 2 "      ' flag for JI
        '        If l = 1 Then lflag = " 4 "      'flag for JIHOV
        '        If l = 2 Then lflag = " 6 "   ' flag for IJTR
        '        'If l = 3 Then lflag = " 8 "  ' flag for IJHOV
        '    End If

        Dim s As String
        Select Case sType
            Case "GP"
                s = IIf(direction = "IJ", "1", "2")
            Case "HOV"
                s = IIf(direction = "IJ", "3", "4")
            Case "TR"
                s = IIf(direction = "IJ", "5", "6")
        End Select
        getFlag = s
    End Function

    Private Function getLinkType(ByVal pFeat As IFeature) As String
        'jaf--this block writes the linktype number (between 1-99, used to flag screen lines)
        'jaf--(column linktype)
        'jaf--default to 90 for now, we'll have to figure out how to capture this in GeoDB

        Dim Tindex As Long
        Tindex = pFeat.Fields.FindField("LinkType")
        If (IsDBNull(pFeat.Value(Tindex)) = False And pFeat.Value(Tindex) > 0) Then
            getLinkType = pFeat.Value(Tindex)
        Else
            getLinkType = "90"
        End If
    End Function
    Private Function CreateWeaveNode(ByVal pFeat As IFeature, ByVal nodeID As Long, ByVal IorJ As Integer, ByVal WeaveType As String, ByVal lWeaveNode As Long, _
    ByVal dctJcts As Dictionary, ByVal dctWNodes As Dictionary, ByVal dctWNodes2 As Dictionary, ByRef nodeString As String)

        Dim pPoint As IPoint
        Dim pJFeat As IFeature
        Dim fld As Long

        If IorJ = 0 Then
            fld = pFeat.Fields.FindField(WeaveType & "_I")
        Else
            fld = pFeat.Fields.FindField(WeaveType & "_J")
        End If

        pFeat.value(fld) = lWeaveNode - m_Offset
        pFeat.Store()

        pJFeat = dctJcts.Item(CType(nodeID, Integer))
        pPoint = getWeavenode(pJFeat, WeaveType)
        If nodeString = "" Then
            nodeString = "a " + CStr(lWeaveNode) + " " + CStr(pPoint.x) + " " + CStr(pPoint.y) + " 0 0 0 0"
        Else
            nodeString = nodeString + vbCrLf + "a " + CStr(lWeaveNode) + " " + CStr(pPoint.x) + " " + CStr(pPoint.y) + " 0 0 0 0"
        End If

        dctWNodes.Add(Str(lWeaveNode), lWeaveNode - m_Offset)
        Try


            dctWNodes2.Add(CType(nodeID, Integer), CType(lWeaveNode, Integer))
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        pPoint = Nothing
        '    Set pJFeat = Nothing
    End Function


    Private Function GetWeaveLinks(ByVal pFeat As IFeature, ByVal pRow As IRow, ByVal pMRow As IRow, ByVal dctWLinks As Dictionary, ByVal dctSplitLinks As Dictionary, _
        ByVal dctEmme2Nodes As Dictionary, ByVal lType As String, _
        ByVal direction As String, ByVal tpd As String, ByVal lanes As Double, ByRef wNodes As String, ByRef astring As String, ByRef wString As String)
        'pRow is attribute row in modeAttributes or tblLine or eventLine tables
        'pMRow is exclusively the row in modeAttributes

        Dim modeWeave As String
        Dim Leng As Double
        Dim index As Long, iIndex As Long, jIndex As Long, iwIndex As Long, jwIndex As Long
        Dim iwNode As Long, jwNode As Long, INode As Long, JNode As Long
        Dim wID_Type As String, wDefault As String
        Dim ijtemp As String, jitemp As String
        '    Dim lanes As Double

        wID_Type = " 0 0"

        modeWeave = "ahijstuvbw"

        iwIndex = pFeat.Fields.FindField(lType & "_I")
        jwIndex = pFeat.Fields.FindField(lType & "_J")

        If IsDBNull(pFeat.Value(iwIndex)) Or IsDBNull(pFeat.Value(jwIndex)) Then
            Exit Function
        Else
            iwNode = pFeat.Value(iwIndex) + m_Offset
            jwNode = pFeat.Value(jwIndex) + m_Offset
        End If

        If direction = "IJ" Then
            wNodes = CStr(iwNode) + " " + CStr(jwNode)
        Else
            wNodes = CStr(jwNode) + " " + CStr(iwNode)
        End If

        index = pFeat.Fields.FindField("LinkType")
        'now assign defaults to weave: length, modes, link type, lanes
        Leng = getWeaveLen(lType) / 5280
        '    lanes = getLanes(pRow, pMRow, lType, Direction, tpd)

        If lanes = 0 Then Exit Function

        If fVerboseLog Then WriteLogLine("leng of weave" + CStr(Leng))
        If Not IsDBNull(pFeat.Value(index)) And pFeat.Value(index) > 0 Then
            wDefault = wDefault + " " + CStr(Leng) + " " + modeWeave + " " + CStr(pFeat.Value(index)) + " " + CStr(lanes)
        Else
            wDefault = wDefault + " " + CStr(Leng) + " " + modeWeave + " 90 " + CStr(lanes)
        End If

        iIndex = pFeat.Fields.FindField("INode")
        jIndex = pFeat.Fields.FindField("JNode")
        INode = pFeat.value(iIndex)
        JNode = pFeat.value(jIndex)
        index = pFeat.Fields.FindField("UseEmmeN")

        If pFeat.Value(index) = 0 Or IsDBNull(pFeat.Value(index)) Then
            INode = INode + m_Offset
            JNode = JNode + m_Offset
        ElseIf pFeat.Value(index) = 1 Then 'inode is emme2id
            JNode = JNode + m_Offset
            INode = dctEmme2Nodes.Item(CStr(INode))
        ElseIf pFeat.Value(index) = 2 Then 'jnode is emme2id
            INode = INode + m_Offset
            JNode = dctEmme2Nodes.Item(CStr(JNode))
        Else
            INode = dctEmme2Nodes.Item(CStr(INode))
            JNode = dctEmme2Nodes.Item(CStr(JNode))
        End If

        'vdf, capacity, FFT, facility type

        Dim sFT As String
        sFT = pFeat.value(pFeat.Fields.FindField("NewFacilityType"))
        ijtemp = "1 1800 60 " & sFT
        jitemp = "1 1800 60 " & sFT


        'Only gives WeaveLinks if it has an HOV attribute in the IJ direction for that time of day!!!!!!Need to check to see if
        Dim wLink As String
        wLink = CStr(iwNode) + " " + CStr(INode)
        If Not dctWLinks.Exists(wLink) And Not dctSplitLinks.Exists(wLink) Then
            If wString = "" Then
                wString = "a " + CStr(iwNode) + " " + CStr(INode) + " " + wDefault + " " + ijtemp
            Else
                wString = wString + vbNewLine + "a " + CStr(iwNode) + " " + CStr(INode) + " " + wDefault + " " + ijtemp
            End If

            If astring = "" Then
                astring = CStr(iwNode) + " " + CStr(INode) + wID_Type
            Else
                astring = astring + vbNewLine + CStr(iwNode) + " " + CStr(INode) + wID_Type
            End If

            'now write opposite
            wString = wString + vbNewLine + "a " + CStr(INode) + " " + CStr(iwNode) + " " + wDefault + " " + jitemp
            astring = astring + vbNewLine + CStr(INode) + " " + CStr(iwNode) + wID_Type
            dctWLinks.Add(wLink, wLink)
        End If

        wLink = CStr(jwNode) + " " + CStr(JNode)
        If Not dctWLinks.Exists(wLink) And Not dctSplitLinks.Exists(wLink) Then
            If wString = "" Then
                wString = "a " + wLink + " " + wDefault + " " + ijtemp
            Else
                wString = wString + vbNewLine + "a " + wLink + " " + wDefault + " " + ijtemp
            End If

            If astring = "" Then
                astring = CStr(jwNode) + " " + CStr(JNode) + wID_Type
            Else
                astring = astring + vbNewLine + CStr(jwNode) + " " + CStr(JNode) + wID_Type
            End If

            'now write opposite
            wString = wString + vbNewLine + "a " + CStr(JNode) + " " + CStr(jwNode) + " " + wDefault + " " + jitemp
            astring = astring + vbNewLine + CStr(JNode) + " " + CStr(jwNode) + wID_Type
            dctWLinks.Add(wLink, wLink)
        End If

        If fVerboseLog Then WriteLogLine(wNodes)
        If fVerboseLog Then WriteLogLine(wString)


    End Function
    Public Sub createWeaveLink(ByVal lType As String, ByVal pFeat As IFeature, ByVal pAttRow As IRow, _
    ByVal wString As String, ByVal wNodes As String, ByVal astring As String, _
    ByVal dctEmme2Nodes As Dictionary, ByVal dctWNodes As Dictionary, ByVal dctWNodes2 As Dictionary, _
    ByVal dctJcts As Dictionary, ByVal dctEdges As Dictionary, ByRef nodeString As String)
        'dctWNodes is a dictionary of weave node IDs. key=cstr(weavenode id), item=weavenode id
        'dctWNodes2 is a dictionary of Scen_Node and the corresponding weave node. key=Scen_Node, item=weave node

        'ptstring: string for weave nodes
        'wstring: string for weave links
        'astring: string for attributes

        'on error GoTo eh
        Const wID_Type = " 0 0"
        '    If fVerboseLog Then WriteLogLine "createWeaveLink called with pstring=" & pstring
        If fVerboseLog Then WriteLogLine("weave link at feat id " & pFeat.OID)
        If fVerboseLog Then WriteLogLine("-------------------------------------------------")

        Dim pFC As IFeatureCursor
        Dim pNewFeat As IFeature
        Dim pefeat As IFeature, pJFeat As IFeature
        Dim pGeom As IGeometry
        Dim pPt As IPoint, pOpt As IPoint
        Dim pFilter As ISpatialFilter
        Dim pPoint As IPoint

        Dim INode As Long, iwNode As Long
        Dim JNode As Long, jwNode As Long
        Dim iIndex As Integer, iwIndex As Integer
        Dim jIndex As Integer, jwIndex As Integer
        Dim index As Integer
        Dim wDefault As String
        Dim strFFT As String    'holds formatted FFT number
        Dim modeWeave As String

        modeWeave = "ahijstuvbw"
        Try


            Dim laneFld As String
            '[071206] - pan shadow link node ids set to unique values with offset
            'of largestwjunct. Public largestjunct has been incremented in createShapefiles and
            'during splits, but GetLargestID used anyway to find top of stack.

            'GetLargestID m_junctShp, "PSRCJunctID", largestjunct
            'largestwjunct = largestjunct + 1

            '[050407] hyu
            Dim nodeID As Long

            iIndex = pFeat.Fields.FindField("INode")
            jIndex = pFeat.Fields.FindField("JNode")
            INode = pFeat.Value(iIndex)
            JNode = pFeat.Value(jIndex)
            index = pFeat.Fields.FindField("UseEmmeN")

            '[021706] hyu: shouldn't UseemmeN=3 means both ij nodes are emme2id?
            '    If pFeat.value(index) = 3 Then
            If pFeat.Value(index) = 0 Or IsDBNull(pFeat.Value(index)) Then
                INode = INode + m_Offset
                JNode = JNode + m_Offset
            ElseIf pFeat.Value(index) = 1 Then 'inode is emme2id
                JNode = JNode + m_Offset
                INode = dctEmme2Nodes.Item(CStr(INode))
            ElseIf pFeat.Value(index) = 2 Then 'jnode is emme2id
                INode = INode + m_Offset
                JNode = dctEmme2Nodes.Item(CStr(JNode))
            Else
                INode = dctEmme2Nodes.Item(CStr(INode))
                JNode = dctEmme2Nodes.Item(CStr(JNode))
            End If
            If fVerboseLog Then WriteLogLine("linkij " + CStr(INode) + " " + CStr(JNode))
            'if equal 3 then both ia nd j are emmeid's and are ok for scen_node values
            Dim tempnode As Long
            Dim i As Integer, j As Integer
            Dim Leng As Double
            Dim facType As Integer
            Dim NewPoint As IPoint

            index = pFeat.Fields.FindField("FacilityType")
            facType = pFeat.Value(index)
            If fVerboseLog Then WriteLogLine("factype " + CStr(facType))


            'vdf, capacity, FFT, facility type
            Dim ijtemp As String, jitemp As String
            Dim direction As String
            'SEC 092710- Added New Facility Type field to weaves. Should loi
            Dim sFT As String
            Dim sSpeed As String
            Dim sCap As String
            sFT = pFeat.Value(pFeat.Fields.FindField("NewFacilityType"))
            ijtemp = "1 1800 60 " & sFT
            jitemp = "1 1800 60 " & sFT

            iwIndex = pFeat.Fields.FindField(lType + "_I")
            jwIndex = pFeat.Fields.FindField(lType + "_J")
            'If pFeat.OID = 691 Then MsgBox 1
            '    Debug.Print pFeat.OID
            If Not IsDBNull(pFeat.Value(iwIndex)) Then
                If pFeat.Value(iwIndex) > 0 Then 'jnode should also have a value
                    iwNode = pFeat.Value(iwIndex) + m_Offset
                    jwNode = pFeat.Value(jwIndex) + m_Offset

                    If fVerboseLog Then WriteLogLine("have weave' " + wNodes)

                    If Not dctWNodes.Exists(CType(iwNode, Integer)) Then
                        '[051507]hyu: if the weavenode is a junction, it's the split point for case 2 during weave link creation
                        If dctJcts.Exists(CType(iwNode, Integer)) Then
                            pJFeat = dctJcts.Item(CType(iwNode, Integer))
                            pPoint = pJFeat.ShapeCopy
                            updateTransitLines(pPoint)
                        Else
                            pJFeat = dctJcts.Item(CType(INode, Integer))
                            pPoint = getWeavenode(pJFeat, lType)
                            If nodeString = "" Then
                                nodeString = "a " + CStr(iwNode) + " " + CStr(pPoint.X) + " " + CStr(pPoint.Y) + " 0 0 0 0"
                            Else
                                nodeString = nodeString + vbCrLf + "a " + CStr(iwNode) + " " + CStr(pPoint.X) + " " + CStr(pPoint.Y) + " 0 0 0 0"
                            End If
                        End If

                        dctWNodes.Add(CType(iwNode, Integer), iwNode - m_Offset)
                        If Not dctWNodes2.Exists(CType(INode, Integer)) Then dctWNodes2.Add(CType(INode, Integer), CType(iwNode, Integer))
                    End If
                End If

                If Not dctWNodes.Exists(CType(jwNode, Integer)) Then
                    '[051507]hyu: if the weavenode is a junction, it's the split point for case 2 during weave link creation
                    If dctJcts.Exists(CType(jwNode, Integer)) Then
                        pJFeat = dctJcts.Item(CType(jwNode, Integer))
                        pPoint = pJFeat.ShapeCopy
                        updateTransitLines(pPoint)
                    Else
                        pJFeat = dctJcts.Item(CType(JNode, Integer))
                        pPoint = getWeavenode(pJFeat, lType)
                        If nodeString = "" Then
                            nodeString = "a " + CStr(jwNode) + " " + CStr(pPoint.X) + " " + CStr(pPoint.Y) + " 0 0 0 0"
                        Else
                            nodeString = nodeString + vbCrLf + "a " + CStr(jwNode) + " " + CStr(pPoint.X) + " " + CStr(pPoint.Y) + " 0 0 0 0"
                        End If

                    End If

                    dctWNodes.Add(CType(jwNode, Integer), CType(jwNode, Integer) - m_Offset)
                    If Not dctWNodes2.Exists(CType(JNode, Integer)) Then dctWNodes2.Add(CType(JNode, Integer), CType(jwNode, Integer))
                End If
                'have merge nodes and therefore create weave link
                'so can return now with proper inode and jnode for wltype(i)
                Exit Sub

            Else    'no existing weave node recorded in the attributes
                Dim bWeaveNodeExists As Boolean
                Dim bCanSplit As Boolean



                For j = 0 To 1
                    bWeaveNodeExists = False
                    If j = 0 Then
                        If dctWNodes2.Exists(CType(INode, Integer)) = True Then
                            pFeat.Value(pFeat.Fields.FindField(lType & "_I")) = dctWNodes2.Item(CType(INode, Integer)) - m_Offset
                            bWeaveNodeExists = True
                            iwNode = dctWNodes2.Item(CType(INode, Integer))
                        End If
                    Else
                        If dctWNodes2.Exists(CType(JNode, Integer)) Then
                            pFeat.Value(pFeat.Fields.FindField(lType & "_J")) = dctWNodes2.Item(CType(JNode, Integer)) - m_Offset
                            bWeaveNodeExists = True
                            jwNode = dctWNodes.Item(CType(JNode, Integer))
                        End If
                    End If
                    Dim test As Integer
                    pFeat.Store()
                    '[052007] hyu: per discuss w/ Jeff & Andy, temperarily we assume one base node can only correspond to one weave node.
                    'so here we will use the existing weave node.
                    If bWeaveNodeExists = False Then

                        If fVerboseLog Then WriteLogLine("for looping " + CStr(j))
                        'first need to check if connected edges have weave i,j's
                        If j = 0 Then
                            tempnode = INode


                            pJFeat = dctJcts.Item(CType(INode, Integer))
                        ElseIf j = 1 Then
                            tempnode = JNode
                            pJFeat = dctJcts.Item(CType(JNode, Integer))
                        End If
                        If pJFeat Is Nothing Then
                            MessageBox.Show("nothing")
                        End If

                        'get the center point
                        pGeom = pJFeat.ShapeCopy
                        pOpt = pGeom
                        If fVerboseLog Then WriteLogLine("createWeaveLink new point")
                        pFilter = New SpatialFilter
                        With pFilter
                            .Geometry = pOpt 'all the edges in intermediate shp
                            .GeometryField = m_edgeShp.ShapeFieldName
                            .SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects
                            .WhereClause = m_edgeShp.OIDFieldName & "<>" & pFeat.OID
                        End With

                        'now get the opposite point- used in calculating angles for intersections
                        If j = 0 Then
                            pJFeat = dctJcts.Item(CType(JNode, Integer))
                        ElseIf j = 1 Then
                            pJFeat = dctJcts.Item(CType(INode, Integer))
                        End If

                        pGeom = pJFeat.ShapeCopy
                        pPt = pGeom
                        '                Set pJFeat = Nothing

                        'search intersecting edges
                        If m_edgeShp.FeatureCount(pFilter) < 1 Then
                            'if only one edge intersecting, just take the weave node from the intersecting edge.
                            'LargestWJunct = LargestWJunct + 1
                            LargestJunct = LargestJunct + 1
                            'CreateWeaveNode pFeat, tempnode, j, lType, LargestWJunct, dctJcts, dctWNodes, dctWNodes2, nodeString
                            CreateWeaveNode(pFeat, tempnode, j, lType, LargestJunct + m_Offset, dctJcts, dctWNodes, dctWNodes2, nodeString)
                        Else '>1 other edges intersecting
                            pFC = m_edgeShp.Search(pFilter, False)
                            pefeat = pFC.NextFeature
                            '                    If Not UpdateWeaveNode(pFeat, pefeat, j, ltype) Then
                            'the intersecting edge has no weave nodes

                            Do Until pefeat Is Nothing
                                bCanSplit = CanSplit(pFeat, pefeat, pAttRow, lType)
                                If bCanSplit Then
                                    SplitFeatureByM(pefeat, getWeaveLen(lType), tempnode, NewPoint, lType, dctEdges, dctJcts, dctWNodes, dctWNodes2, nodeString)
                                    updateTransitLines(NewPoint)
                                    If j = 0 Then
                                        pFeat.Value(iwIndex) = LargestJunct '+ m_Offset
                                    Else
                                        pFeat.Value(jwIndex) = LargestJunct '+ m_Offset
                                    End If
                                    pFeat.Store()
                                    Exit Do
                                End If
                                pefeat = pFC.NextFeature
                            Loop

                            If Not bCanSplit Then
                                'LargestWJunct = LargestWJunct + 1
                                LargestJunct = LargestJunct + 1
                                'CreateWeaveNode pFeat, tempnode, j, lType, LargestWJunct, dctJcts, dctWNodes, dctWNodes2, nodeString
                                CreateWeaveNode(pFeat, tempnode, j, lType, LargestJunct + m_Offset, dctJcts, dctWNodes, dctWNodes2, nodeString)
                            End If

                        End If
                    End If  'bWeavenodeExists=true
                Next j
            End If  '(Not IsNull(pFeat.value(iwIndex)) And pFeat.value(iwIndex) > 0)
            '    If fVerboseLog Then WriteLogLine "finished createWeaveLink call with ptstring=" & ptstring
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

        Exit Sub

eh:
        CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.createWeaveLink")
        MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.createWeaveLink")
        'Print #9, Err.Description, vbInformation, "createWeaveLink"
        '   If pWSEdit.IsBeingEdited Then
        '      pWSEdit.AbortEditOperation
        '      pWSEdit.StopEditing False
        '   End If

    End Sub
    Private Function getVDF(ByVal pRow As IRow, ByVal pMRow As IRow, ByVal t As Integer, ByVal direction As String, ByVal mode As String, ByVal linkType As String, ByVal FacilityType As String) As String
        't is the time of day array index used in the Create_NetFile2 procedure

        Dim vdf As Integer, sVDF As String
        Dim Tindex As Long
        Dim rString As String, rLocation As Integer

        Tindex = pRow.Fields.FindField(direction + "VDFunc") 'this will some have by timepd
        rString = "r"
        rLocation = InStr(mode, rString)
        vdf = 0

        If (IsDBNull(pRow.Value(Tindex)) = False And pRow.Value(Tindex) > 0) Then
            vdf = pRow.Value(Tindex)
        Else
            Tindex = pMRow.Fields.FindField(direction + "VDFunc") 'this will some have by timepd
            If IsDBNull(pMRow.Value(Tindex)) = False And pMRow.Value(Tindex) > 0 Then
                vdf = pMRow.Value(Tindex)
            Else
                sVDF = "1"
            End If
        End If
        '09/24/2008 SEC Changed the following according to current VDF rules:
        'If vdf > 0 Then
        'If linkType = "32" Then
        'If IsNull(rLocation) Or rLocation = 0 Then
        'If t = 0 Or t = 1 Or t = 3 Then vdf = vdf - 30
        'If t = 2 Then vdf = vdf
        'If t = 4 Then vdf = vdf - 20
        'Else
        'If t = 3 Then vdf = vdf + 10
        'If t = 4 Then vdf = vdf + 20
        'End If
        'Else
        'If t = 3 Then vdf = vdf + 10
        'If t = 4 Then vdf = vdf + 20
        'End If
        'sVDF = CStr(vdf)
        'End If

        'If vdf = 40 Then

        If vdf = 31 Then
            If t = 3 Then
                vdf = 33
            ElseIf t = 4 Then
                vdf = 35
            End If

        ElseIf vdf > 0 And vdf <> 40 And (CInt(FacilityType) < 9 Or CInt(FacilityType) = 16) Then
            If t = 3 Then
                vdf = vdf + 10
            ElseIf t = 4 Then
                vdf = vdf + 20
            End If
        End If



        If vdf < 1 Then
            If t = 3 Then
                vdf = 15
            ElseIf t = 4 Then
                vdf = 25
            Else
                vdf = 5
            End If
        End If

        If vdf > 40 Then
            MsgBox("wrong vdf")
        End If





        sVDF = CStr(vdf)
        getVDF = sVDF
    End Function
    Private Function getSplitLink(ByVal pFeat As IFeature, ByVal dctSplitLinks As Dictionary, ByVal dctEmme2Nodes As Dictionary)
        Dim fldI As Long, fldJ As Long, fldOneWay As Long, fldUseEmmN As Long
        Dim fldHOV As Long, fldTR As Long, fldTK As Long
        Dim INode As Long, JNode As Long, oneWay As Integer
        Dim sKey As String

        fldHOV = pFeat.Fields.FindField("SplitHOV")
        fldTR = pFeat.Fields.FindField("SplitTR")
        fldTK = pFeat.Fields.FindField("SplitTK")

        If (IsDBNull(pFeat.Value(fldHOV)) And IsDBNull(pFeat.Value(fldTR)) And IsDBNull(pFeat.Value(fldTK))) Then
            Exit Function
        End If

        fldI = pFeat.Fields.FindField(g_INode)
        fldJ = pFeat.Fields.FindField(g_JNode)
        fldOneWay = pFeat.Fields.FindField(g_OneWay)
        fldUseEmmN = pFeat.Fields.FindField(g_UseEmmeN)

        INode = pFeat.value(fldI)
        JNode = pFeat.value(fldJ)
        If pFeat.Value(fldUseEmmN) = 0 Or IsDBNull(pFeat.Value(fldUseEmmN)) Then
            INode = INode + m_Offset
            JNode = JNode + m_Offset
        ElseIf pFeat.Value(fldUseEmmN) = 1 Then 'inode is emme2id
            JNode = JNode + m_Offset
            INode = dctEmme2Nodes.Item(CStr(INode))
        ElseIf pFeat.Value(fldUseEmmN) = 2 Then 'jnode is emme2id
            INode = INode + m_Offset
            JNode = dctEmme2Nodes.Item(CStr(JNode))
        Else
            INode = dctEmme2Nodes.Item(CStr(INode))
            JNode = dctEmme2Nodes.Item(CStr(JNode))
        End If

        oneWay = pFeat.value(fldOneWay)

        Select Case oneWay
            Case 0
                sKey = CStr(INode) + " " + CStr(JNode)
                dctSplitLinks.Add(sKey, sKey)
            Case 1
                sKey = CStr(JNode) + " " + CStr(INode)
                dctSplitLinks.Add(sKey, sKey)
            Case 2, 3
                sKey = CStr(INode) + " " + CStr(JNode)
                dctSplitLinks.Add(sKey, sKey)
                sKey = CStr(JNode) + " " + CStr(INode)
                dctSplitLinks.Add(sKey, sKey)
        End Select


    End Function
    Private Function getLaneCap(ByVal pRow As IRow, ByVal pMRow As IRow, ByVal direction As String, ByVal sType As String) As String
        'pRow can be the corresponding attributes row in tblLine, evtLine, or modeAttributes
        'pMRow is the attributes row in modeAttributes.

        Dim Tindex As Long
        Dim sCap As String

        'NOW assign: a iNode, jNode, length, modes, linkType, #lane, functional class, capacity
        If sType = "GP" Or sType = "HOV" Then 'If l < 2 Then
            Tindex = pRow.Fields.FindField(direction + "LaneCap" + sType)
        Else
            Tindex = pRow.Fields.FindField(direction + "LaneCap" + "GP")
        End If

        If Not IsDBNull(pRow.Value(Tindex)) And pRow.Value(Tindex) > 0 Then
            sCap = CStr(pRow.Value(Tindex))
        Else
            If sType = "GP" Or sType = "HOV" Then 'If l < 2 Then
                Tindex = pMRow.Fields.FindField(direction + "LaneCap" + sType)
            Else
                Tindex = pMRow.Fields.FindField(direction + "LaneCap" + "GP")
            End If
            If Not IsDBNull(pMRow.Value(Tindex)) And pMRow.Value(Tindex) > 0 Then
                sCap = CStr(pMRow.Value(Tindex))
            Else
                sCap = "1800"
            End If
        End If
        getLaneCap = sCap

    End Function
    Private Function getSpeedLimit(ByVal pRow As IRow, ByVal pMRow As IRow, ByVal sType As String, ByVal direction As String, ByVal FacilityType As String, ByVal pFeature As IFeature) As String
        Dim Tindex As Long
        Dim strSpeedLimit As String
        Dim strName As String
        Dim strSounderName As String
        Dim dblLength As Double
        Dim dblSpeed As Double
        Dim dblSpeed2 As Double
        Dim lngSpeed As Long
        Tindex = pRow.Fields.FindField(direction + "SpeedLimit")


        dblLength = pFeature.value(pFeature.Fields.FindField("Shape.len"))


        If FacilityType = "14" Or FacilityType = "15" Then
            strName = pFeature.value(pFeature.Fields.FindField("Fullname"))
            Select Case strName
                Case "Clinton"
                    dblSpeed = 0.001369241 * dblLength
                    dblSpeed = Round(dblSpeed, 2)
                    If dblSpeed = 0 Then
                        strSpeedLimit = "0.01"
                    Else
                        strSpeedLimit = CStr(dblSpeed)
                    End If
                Case "Edmonds"
                    dblSpeed = 0.001026187 * dblLength
                    dblSpeed = Round(dblSpeed, 2)
                    If dblSpeed = 0 Then
                        strSpeedLimit = "0.01"
                    Else
                        strSpeedLimit = CStr(dblSpeed)
                    End If

                Case "Bainbridge"
                    dblSpeed = 0.000803369 * dblLength
                    dblSpeed = Round(dblSpeed, 2)
                    If dblSpeed = 0 Then
                        strSpeedLimit = "0.01"
                    Else
                        strSpeedLimit = CStr(dblSpeed)
                    End If

                Case "Bremerton"
                    dblSpeed = 0.000677703 * dblLength
                    dblSpeed = Round(dblSpeed, 2)
                    If dblSpeed = 0 Then
                        strSpeedLimit = "0.01"
                    Else
                        strSpeedLimit = CStr(dblSpeed)
                    End If

                Case "SouthworthWS"
                    dblSpeed = 0.001548318 * dblLength
                    dblSpeed = Round(dblSpeed, 2)
                    If dblSpeed = 0 Then
                        strSpeedLimit = "0.01"
                    Else
                        strSpeedLimit = CStr(dblSpeed)
                    End If

                Case "SouthworthVashon"
                    dblSpeed = 0.001074801 * dblLength
                    dblSpeed = Round(dblSpeed, 2)
                    If dblSpeed = 0 Then
                        strSpeedLimit = "0.01"
                    Else
                        strSpeedLimit = CStr(dblSpeed)
                    End If

                Case "VashonWS"
                    dblSpeed = 0.001488144 * dblLength
                    dblSpeed = Round(dblSpeed, 2)
                    If dblSpeed = 0 Then
                        strSpeedLimit = "0.01"
                    Else
                        strSpeedLimit = CStr(dblSpeed)
                    End If

                Case "PtDef"
                    dblSpeed = 0.001488144 * dblLength
                    dblSpeed = Round(dblSpeed, 2)
                    If dblSpeed = 0 Then
                        strSpeedLimit = "0.01"
                    Else
                        strSpeedLimit = CStr(dblSpeed)
                    End If

                Case "VPOF"
                    dblSpeed = 0.00048542 * dblLength
                    dblSpeed = Round(dblSpeed, 2)
                    If dblSpeed = 0 Then
                        strSpeedLimit = "0.01"
                    Else
                        strSpeedLimit = CStr(dblSpeed)
                    End If
                Case "ferry"
                    dblSpeed = 0.00181197 * dblLength
                    dblSpeed = Round(dblSpeed, 2)
                    If dblSpeed = 0 Then
                        strSpeedLimit = "0.01"
                    Else
                        strSpeedLimit = CStr(dblSpeed)
                    End If
                Case "Elliot Bay Water Taxi"
                    dblSpeed = 0.001082051 * dblLength
                    dblSpeed = Round(dblSpeed, 2)
                    If dblSpeed = 0 Then
                        strSpeedLimit = "0.01"
                    Else
                        strSpeedLimit = CStr(dblSpeed)
                    End If
                Case "WFSKN"
                    dblSpeed = 0.000375410231 * dblLength
                    dblSpeed = Round(dblSpeed, 2)
                    If dblSpeed = 0 Then
                        strSpeedLimit = "0.01"
                    Else
                        strSpeedLimit = CStr(dblSpeed)
                    End If

                Case "WFBDM"
                    dblSpeed = 0.00047829521 * dblLength
                    dblSpeed = Round(dblSpeed, 2)
                    If dblSpeed = 0 Then
                        strSpeedLimit = "0.01"
                    Else
                        strSpeedLimit = CStr(dblSpeed)
                    End If
                Case "WFKUW"
                    dblSpeed = 0.00058933092 * dblLength
                    dblSpeed = Round(dblSpeed, 2)
                    If dblSpeed = 0 Then
                        strSpeedLimit = "0.01"
                    Else
                        strSpeedLimit = CStr(dblSpeed)
                    End If
                Case "WFKUW2"
                    dblSpeed = 0.000454225951 * dblLength
                    dblSpeed = Round(dblSpeed, 2)
                    If dblSpeed = 0 Then
                        strSpeedLimit = "0.01"
                    Else
                        strSpeedLimit = CStr(dblSpeed)
                    End If
                Case "Renton_Leschi-PO"
                    dblSpeed = 0.0005000164 * dblLength
                    dblSpeed = Round(dblSpeed, 2)
                    If dblSpeed = 0 Then
                        strSpeedLimit = "0.01"
                    Else
                        strSpeedLimit = CStr(dblSpeed)
                    End If

                Case "DesMoines-DTS-PO"
                    dblSpeed = 0.000364854684 * dblLength
                    dblSpeed = Round(dblSpeed, 2)
                    If dblSpeed = 0 Then
                        strSpeedLimit = "0.01"
                    Else
                        strSpeedLimit = CStr(dblSpeed)
                    End If

                Case "Shilshole-Seattle-PO"
                    dblSpeed = 0.00055975697289 * dblLength
                    dblSpeed = Round(dblSpeed, 2)
                    If dblSpeed = 0 Then
                        strSpeedLimit = "0.01"
                    Else
                        strSpeedLimit = CStr(dblSpeed)
                    End If
                Case "WFSSW"
                    dblSpeed = 0.00029331341327 * dblLength
                    dblSpeed = Round(dblSpeed, 2)
                    If dblSpeed = 0 Then
                        strSpeedLimit = "0.01"
                    Else
                        strSpeedLimit = CStr(dblSpeed)
                    End If
                Case Else
                    dblSpeed = 0.000735354 * dblLength
                    dblSpeed = Round(dblSpeed, 2)
                    If dblSpeed = 0 Then
                        strSpeedLimit = "0.01"
                    Else
                        strSpeedLimit = CStr(dblSpeed)
                    End If

            End Select
        ElseIf FacilityType = "10" Then
            dblSpeed = 0.00378787878787879 * dblLength
            dblSpeed = Round(dblSpeed, 2)
            If dblSpeed = 0 Then
                strSpeedLimit = "0.01"
            Else
                strSpeedLimit = CStr(dblSpeed)
            End If



        ElseIf FacilityType = "12" Or FacilityType = "11" Or FacilityType = "13" Then
            If IsDBNull(pFeature.Value(pFeature.Fields.FindField("Processing"))) = False Then
                dblSpeed = pFeature.Value(pFeature.Fields.FindField("Processing")) / 1000
                dblSpeed = Round(dblSpeed, 2)
                If dblSpeed = 0 Then
                    strSpeedLimit = "0.01"
                Else
                    strSpeedLimit = CStr(dblSpeed)
                End If
            Else

                strName = pFeature.Value(pFeature.Fields.FindField("Fullname"))
                'Dim dblRailLength As Double
                'dblRailLength = pFeature.value(pFeature.Fields.FindField("Shape.len"))
                'strSounderName = pFeature.value(pFeature.Fields.FindField("Fullname"))
                'Dim dblRailSpeed As Double
                Select Case strName
                    Case "Sounder-TS"
                        dblSpeed = 0.000283311487 * dblLength
                        dblSpeed = Round(dblSpeed, 2)
                        If dblSpeed = 0 Then
                            strSpeedLimit = "0.01"
                        Else
                            strSpeedLimit = CStr(dblSpeed)
                        End If
                    Case "Sounder-ES"
                        dblSpeed = 0.00032637 * dblLength
                        dblSpeed = Round(dblSpeed, 2)
                        If dblSpeed = 0 Then
                            strSpeedLimit = "0.01"
                        Else
                            strSpeedLimit = CStr(dblSpeed)
                        End If

                    Case "TacomaLR"
                        dblSpeed = 0.0007695367 * dblLength
                        dblSpeed = Round(dblSpeed, 2)
                        If dblSpeed = 0 Then
                            strSpeedLimit = "0.01"
                        Else
                            strSpeedLimit = CStr(dblSpeed)
                        End If
                    Case "Monorail"
                        dblSpeed = 0.00044056162 * dblLength
                        dblSpeed = Round(dblSpeed, 2)
                        If dblSpeed = 0 Then
                            strSpeedLimit = "0.01"
                        Else
                            strSpeedLimit = CStr(dblSpeed)
                        End If
                    Case "BNSF_RR"
                        dblSpeed = 0.00047414 * dblLength
                        dblSpeed = Round(dblSpeed, 2)
                        If dblSpeed = 0 Then
                            strSpeedLimit = "0.01"
                        Else
                            strSpeedLimit = CStr(dblSpeed)
                        End If
                    Case "HCT-1"
                        dblSpeed = 0.00023238 * dblLength
                        dblSpeed = Round(dblSpeed, 2)
                        If dblSpeed = 0 Then
                            strSpeedLimit = "0.01"
                        Else
                            strSpeedLimit = CStr(dblSpeed)
                        End If
                    Case "HCT-2"
                        dblSpeed = 0.0003108826 * dblLength
                        dblSpeed = Round(dblSpeed, 2)
                        If dblSpeed = 0 Then
                            strSpeedLimit = "0.01"
                        Else
                            strSpeedLimit = CStr(dblSpeed)
                        End If
                    Case "LRT-B-S-WS"
                        dblSpeed = 0.0003108826 * dblLength
                        dblSpeed = Round(dblSpeed, 2)
                        If dblSpeed = 0 Then
                            strSpeedLimit = "0.01"
                        Else
                            strSpeedLimit = CStr(dblSpeed)
                        End If
                    Case "Link Light Rail"
                        dblSpeed = 0.00036631172 * dblLength
                        dblSpeed = Round(dblSpeed, 2)
                        If dblSpeed = 0 Then
                            strSpeedLimit = "0.01"
                        Else
                            strSpeedLimit = CStr(dblSpeed)
                        End If
                    Case "LRT-B-UW-Red"
                        dblSpeed = 0.00041426 * dblLength
                        dblSpeed = Round(dblSpeed, 2)
                        If dblSpeed = 0 Then
                            strSpeedLimit = "0.01"
                        Else
                            strSpeedLimit = CStr(dblSpeed)
                        End If

                    Case Else
                        dblSpeed = 0.000558045 * dblLength
                        dblSpeed = Round(dblSpeed, 2)
                        If dblSpeed = 0 Then
                            strSpeedLimit = "0.01"
                        Else
                            strSpeedLimit = CStr(dblSpeed)
                        End If
                End Select
            End If

        ElseIf FacilityType = "9" Then
            If IsDBNull(pFeature.Value(pFeature.Fields.FindField("Processing"))) = False Then
                dblSpeed = pFeature.Value(pFeature.Fields.FindField("Processing")) / 1000
                dblSpeed = Round(dblSpeed, 2)
                If dblSpeed = 0 Then
                    strSpeedLimit = "0.01"
                Else
                    strSpeedLimit = CStr(dblSpeed)
                End If
            Else
                strName = pFeature.Value(pFeature.Fields.FindField("Fullname"))
                Select Case strName
                    Case "E-3 Busway"
                        dblSpeed = 0.000568181 * dblLength
                        dblSpeed = Round(dblSpeed, 2)
                        If dblSpeed = 0 Then
                            strSpeedLimit = "0.01"
                        Else
                            strSpeedLimit = CStr(dblSpeed)
                        End If
                    Case "Bus Tunnel"
                        dblSpeed = 0.000568181 * dblLength
                        dblSpeed = Round(dblSpeed, 2)
                        If dblSpeed = 0 Then
                            strSpeedLimit = "0.01"
                        Else
                            strSpeedLimit = CStr(dblSpeed)
                        End If
                    Case Else
                        dblSpeed = 0.000568181 * dblLength
                        dblSpeed = Round(dblSpeed, 2)
                        If dblSpeed = 0 Then
                            strSpeedLimit = "0.01"
                        Else
                            strSpeedLimit = CStr(dblSpeed)
                        End If
                End Select

            End If

        Else

            If (IsDBNull(pRow.Value(Tindex)) = False And pRow.Value(Tindex) > 0) Then
                'tempstring = tempstring + " " + CStr(pRow.value(Tindex))
                'jaf--converted to FormatNumber 12.02.03
                'tempstring = tempstring + " " + FormatNumber(pRow.value(Tindex), 2, vbTrue)
                strSpeedLimit = FormatNumber(pRow.Value(Tindex), 2, vbTrue)
                'If strFFT = "0.00" Then strFFT = "0.01"
            Else
                Tindex = pMRow.Fields.FindField(direction + "SpeedLimit")
                If (IsDBNull(pMRow.Value(Tindex)) = False And pMRow.Value(Tindex) > 0) Then
                    strSpeedLimit = FormatNumber(pMRow.Value(Tindex), 2, vbTrue)
                    'If strFF = "0.00" Then strFFT = "0.01"
                Else
                    If sType = "GP" Then
                        strSpeedLimit = "30"
                    Else
                        strSpeedLimit = "60"
                    End If
                End If
            End If
        End If
        getSpeedLimit = strSpeedLimit
    End Function
    Private Function getOldFacilityType(ByVal pFeat As IFeature) As String

        Dim frmNetLayer As frmNetLayer
        Dim Tindex As Long
        Dim intFacType As Integer
        'jaf--write out the Functional Class as model Facility Type

        'SEC 070909- Changed FacilityType to NewFacilityType to reflect Emme2 FacilityType, NewFacilityType is the field
        'temporarily holding Emme1 FT.
        Tindex = pFeat.Fields.FindField("FacilityType") 'this is changing to E2FacType
        If Tindex < 0 Then
            'ERROR--No FacilityType field!!!
            intFacType = 0

            'write out error message
            WriteLogLine("DATA ERROR: GlobalMod.create_NetFile: edge pFeat.OID=" & CStr(pFeat.OID) & " has no Functional field, using 0")
            frmNetLayer.lblStatus.Text = "GlobalMod.create_NetFile: edge pFeat.OID=" & CStr(pFeat.OID) & " has no Functional field, using 0"
            'pan frmNetLayer.Refresh
        Else
            'we have a FacilityType field in pFeat
            If IsDBNull(pFeat.Value(Tindex)) Then
                'jaf--facility type is null, which is bad, but we'll fix later
                intFacType = 0
            Else
                'jaf--got good facility type
                intFacType = pFeat.Value(Tindex)
            End If
        End If

        getOldFacilityType = CStr(intFacType)
    End Function

    
End Module

