Imports System
Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.ADF



Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.SystemUI
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Carto
Imports System.Windows.Forms
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.DataSourcesGDB

Imports Microsoft.SqlServer


Imports System.Math

Imports ESRI.ArcGIS.GeoDatabaseUI
Imports ESRI.ArcGIS.DataSourcesFile


Module test

    Public Sub create_NetFile3(ByVal pathnameN As String, ByVal filenameN As String, ByVal pWS As IWorkspace, ByVal App As IApplication)
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
        timePd.Add("am")
        timePd.Add("md")
        timePd.Add("pm")
        timePd.Add("ev")
        timePd.Add("ni")
        mydate = Date.Now


        For t = 0 To 4
            'pathTOD(t) = pathnameN + "\" + filenameN + CStr(timePd(t)) + ".txt"
            pathTOD(t) = pathnameN + "\" + CStr(timePd(t)) + "_" + filenameN
            FileClose(t + 1)
            FileOpen(t + 1, pathTOD(t), OpenMode.Output, OpenAccess.Default, OpenShare.Shared)
            PrintLine((t + 1), "c Exported " + (timePd(t)) + " Network File from ArcMap /Emme/2 Interface: " + CStr(mydate))
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
        Dim pFeatlayerIE As IFeatureLayer
        Dim pFeatlayerJunctions As IFeatureLayer
        pFeatlayerJunctions = New FeatureLayer

        pFeatlayerIE = New FeatureLayer
        pFeatlayerIE.FeatureClass = m_edgeShp
        pFeatlayerJunctions.FeatureClass = m_junctShp


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
        Dim dctReservedNodes As New Dictionary(Of Long, Long)  '[042506] hyu: dictionary for reserved nodes: key=PSRCJunctId, item=Emme2Node
        'Dim dctWeaveNodes(3) As clsCollection       '[051106] hyu: dictionary for weave nodes not physically in the ScenarioJunct layer
        'Dim dctWeaveNodes2(3) As clsCollection    '[052007] hyu: dictionary for junctions and corresponding weave nodes
        Dim dctWeaveNodes(3) As Dictionary(Of String, Long)
        Dim dctWeaveNodes2(3) As Dictionary(Of Long, Long)

        'Dictionary to hold New Weave junctions to write to shapefiles:
        Dim dctWeaveJunctions As Dictionary(Of Long, ESRI.ArcGIS.Geometry.Point)
        dctWeaveJunctions = New Dictionary(Of Long, ESRI.ArcGIS.Geometry.Point)
        Dim dctExportWeaveLinks As Dictionary(Of Long, Long)
        dctExportWeaveLinks = New Dictionary(Of Long, Long)
        Dim dctTODWeaveLinks(4) As Dictionary(Of Long, Long)
        Dim dctHOVModeAtts As New Dictionary(Of Long, Row)


        Dim testt(3) As Dictionary(Of Long, Long)
        Dim lPSRCJctID As Long
        ' dctReservedNodes = New Dictionary



        For l = 0 To 3

            dctWeaveNodes(l) = New Dictionary(Of String, Long)
            dctWeaveNodes2(l) = New Dictionary(Of Long, Long)
        Next l





        Dim dctJcts As New Dictionary(Of Long, Feature), dctJoints As New Dictionary(Of Long, Long), dctEdgeID As New Dictionary(Of Long, Long), lMaxEdgeOID As Long
        Dim pNewFeat As IFeature
        ' dctJcts = New Dictionary(Of Feature, Feature)

        WriteLogLine("start getting all nodes " & Now())
        pcount = 1
        '[062007] jaf: populate the junctions dictionary
        Try


            Do Until pFeat Is Nothing 'loop through features in intermediate junction shapefile
                pPoint = pFeat.Shape
                '[042706] pan:  added parkandride node subtype 7 to test-they need a * also
                '[090408] pan:  there are junctiontype parkandride subtype 7 that are not currently modeled and
                '               would bump us over 1200 limit if used, so check EMME2nodeid > 0
                'check to see if scene_node is a zone
                If pFeat.Value(lSFld) <= m_Offset Then
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
                    If Not dctReservedNodes.ContainsKey(CStr(lPSRCJctID)) Then dctReservedNodes.Add(CStr(lPSRCJctID), CStr(pFeat.Value(lSFld)))
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
        Dim dctWeaveLinks(5) As Dictionary(Of String, String)
        Dim sDir() As String
        Dim switchNodes As String()
        Dim onewayWalkLink As String
        Dim buildOneWayReverseWalkLink As Boolean = False
        Dim sLeng As String
        Dim m_App As IApplication
        m_App = App

        'dctWeaveLinks = New Dictionary

        '[033106] hyu: declare a lookup dictionary for edges in the ModeAttributes table.
        '              This dictionary replaces the query for row count in the ModeAttributes table.
        Dim dctMAtt As New Dictionary(Of Long, Row), dctEdges As New Dictionary(Of Long, Feature)
        Dim bEOF As Boolean, lEdgeCt As Long, sEdgeOID As String, sEdgeID As String
        Dim dctSplit As New Dictionary(Of String, String)
        'dctSplit = New Dictionary(Of Long, Long)

        WriteLogLine("start getting all edges and mode attributes " & Now())
        '[062007] jaf: appears to populate dctEdges with intermed. edge features...
        '              ...and dctMAtt with related modeAttributes records
        '              pTblMode is handle to ModeAttributes
        '              m_edgeShp is handle to ScenarioEdges
        'dctEdges = New Dictionary
        'dctMAtt = New Dictionary
        getMAtts2(m_edgeShp, pTblMode, g_PSRCEdgeID, dctEdges, dctMAtt)

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

            Dim pair As KeyValuePair(Of Long, Feature)




            For Each pair In dctEdges
                'key = dctEdges.GetEnumerator.Current.
                pFeat = pair.Value
                If dctMAtt.ContainsKey(pFeat.Value(pFeat.Fields.FindField(g_PSRCEdgeID))) Then
                    pMRow = dctMAtt.Item(pFeat.Value(pFeat.Fields.FindField(g_PSRCEdgeID)))
                    'getAttributes row returns proper attributes for the edge feature based on Updated1 field
                    '...from related ModeAttributes, tblLineProjects, or evtLineProjects records (as row obj)
                    pRow = getAttributesRow(pFeat, pMRow, pTblPrjEdge, pEvtLine)

                    '************************************************
                    'PROBLEM:  should check TK also!!!
                    '************************************************

                    '[062007] jaf: checks lane types 1, 2 (TR, HOV) to see if need weave; SHOULD check 3 (TK) ALSO!!
                    '[092710] SEC: All combinations of HOV/TR are handled in the HOV fields (coded values). No need to go check TR fields any longer
                    l = 2
                    If getAllLanes2(pRow, pMRow, CStr(lType(l))) > 0 Then
                        'If pRow.Value(pRow.Fields.FindField("IJLanesGPAM")) > 0 Or pRow.Value(pRow.Fields.FindField("JILanesGPAM")) > 0 Or pRow.Value(pMRow.Fields.FindField("IJLanesGPAM")) > 0 Or pMRow.Value(pRow.Fields.FindField("JILanesGPAM")) > 0 Then

                        'No Time of day at this point
                        createWeaveLink2(CStr(lType(l)), pFeat, pRow, wString, nodes, astring, dctReservedNodes, dctWeaveNodes(CType(l, String)), dctWeaveNodes2(l), dctJcts, dctEdges, pstring, dctWeaveJunctions)
                        'End If
                    End If

                End If
                lEdgeCt = lEdgeCt + 1
                If lEdgeCt = dctEdges.Count Then bEOF = True
                WriteLogLine("edge count=" & lEdgeCt & " " & Now())
                '        If pFeat.OID = 1508230 Then MsgBox 1
                'If lEdgeCt = 15000 Then MsgBox lEdgeCt
                getSplitLink2(pFeat, dctSplit, dctReservedNodes)
            Next


            For t = 0 To 4
                PrintLine(t + 1, pstring)

                PrintLine(t + 1)
                PrintLine(t + 1, "t links init")
                dctWeaveLinks(t) = New Dictionary(Of String, String)
                dctTODWeaveLinks(t) = New Dictionary(Of Long, Long)

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
            For Each pair In dctEdges

                buildOneWayReverseWalkLink = False
                pFeat = pair.Value

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


                    nodes = getIJNodes2(pFeat, sDir(i), m_Offset, dctReservedNodes)


                    'jaf--this block writes the length in miles--6 ASCII CHARS MAX LEN!
                    'jaf--(column len(miles))

                    If fVerboseLog Then WriteLogLine("nodes string=" & nodes)

                    sLeng = getLength2(strLayerPrefix, pFeat)
                    'if one way and mode has w in in it, need to create reverse walk link
                    

                    'jaf--lane type is derived from projects, if any
                    'jaf--assume this all works for now
                    '[033006] hyu: get correct link attributes as following
                    'a real link from TransRefEdges: modeAttributes (pTC)
                    'a project link: pTblPrjEdge
                    'a event link: pEvtLine
                    Tindex = pFeat.Fields.FindField("Scen_Link") 'this links the attribute file (which gets included in the punch) back tot he intermediate shp
                    tempAttrib = nodes + " " + CStr(pFeat.Value(Tindex))

                    '[033106] hyu: change to lookup a dictionary before doing the query
                    If Not dctMAtt.ContainsKey(CStr(pFeat.Value(lSFld))) Then
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

                            lflag = getFlag2(CStr(lType(l)), sDir(i))
                            mode = getModes2(pFeat, CStr(lType(l)))

                            'if one way and mode has w in in it, need to create reverse walk link. Only for One way GP lanes.
                            If (intWays = 0 Or intWays = 1) And mode.Contains("w") And l <> 2 Then
                                switchNodes = nodes.Split(New Char() {" "c})
                                onewayWalkLink = "a " + CType(switchNodes(1), String) + " " + CType(switchNodes(0), String) + " " + sLeng + " wk 90 1 9 0 3 0"
                                buildOneWayReverseWalkLink = True
                                
                            End If

                            linkType = getLinkType2(pFeat)

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
                                lanes = getLanes2(pRow, pMRow, CStr(lType(l)), sDir(i), CStr(timePd(t)))
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
                                        'sec added 4/21/09. 
                                        'If pRow.Value(pRow.Fields.FindField("IJLanesGPAM")) > 0 Or pRow.Value(pRow.Fields.FindField("JILanesGPAM")) > 0 Or pRow.Value(pMRow.Fields.FindField("IJLanesGPAM")) > 0 Or pMRow.Value(pRow.Fields.FindField("JILanesGPAM")) > 0 Then
                                        GetWeaveLinks2(pFeat, pRow, pMRow, dctWeaveLinks(t), dctSplit, dctReservedNodes, CStr(lType(l)), sDir(i), CStr(timePd(t)), lanes, nodes, astring, wString, dctTODWeaveLinks(t))
                                        '                                length = getWeaveLen(CStr(lType(l))) / 5280 'convert feet to mile
                                        'End If

                                        If l <> 2 Then 'not HOV
                                            lanes = 1
                                            strTOD(t) = "a " + nodes + " " + sLeng + " " + mode + " " + linkType + " " + CStr(lanes)

                                            '*************TO DO*********This is where we deal with HOV/TR/BAT/HOT/TR lanes- Needs to be coded to handle all possibilities
                                        Else    'hov modes depends on the lanes
                                            'wasHOV = True
                                            Select Case lanes
                                                Case 1
                                                    'HOV 2+
                                                    mode = "ahijbdmgp"

                                                    lanes = 1
                                                Case 2
                                                    'HOV 3+
                                                    mode = "aijbmgp"
                                                    lanes = 1

                                                Case 3
                                                    '1 lane transit only
                                                    mode = "bp"
                                                    lanes = 1

                                                Case 4
                                                    '1 lane bat
                                                    mode = "bp"
                                                    lanes = 1

                                                Case 5
                                                    'One lane HOT
                                                    mode = "ashijtuvbedmgp"

                                                    lanes = 1


                                                Case 7
                                                    'Two lanes of HOV
                                                    mode = "ahijbdmgp"
                                                    lanes = 2
                                                Case 8
                                                    mode = "aijbmgp"
                                                    lanes = 2

                                                Case 12
                                                    'Two lanes HOT, 2+ free
                                                    mode = "ashijtuvbedmgp"

                                                    lanes = 2
                                                Case 13
                                                    'Two lanes HOT, 3+ free
                                                    mode = "ashijtuvbedmgp"
                                                    lanes = 2
                                                Case 14
                                                    'One lanes HOT, No heavy/med trucks
                                                    mode = "ashijvbedmgp"
                                                    lanes = 1
                                                Case 15
                                                    'Two lanes HOT, No heavy/med trucks
                                                    mode = "ashijvbedmgp"
                                                    lanes = 2

                                            End Select

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
                                    'Bus/Bat lanes New Facility Type & VDF do not Pivot of GP Link
                                    If mode = "bp" And l = 2 Then
                                        sVDF = "9"
                                        sNewFacilityType = "0"
                                        sCap = 0

                                        sOldFacilityType = getOldFacilityType2(pFeat)
                                        sFFS = getSpeedLimit2(pRow, pMRow, CStr(lType(l)), sDir(i), sOldFacilityType, pFeat)
                                        sFacType = 0
                                        strTOD(t) = strTOD(t) + " " + sVDF + " " + sCap + " " + sFFS + " " + sFacType

                                    Else
                                        'Shadow link gets these attributes from GP link
                                        sNewFacilityType = getNewFacilityType2(pFeat)
                                        sVDF = getVDF2(pRow, pMRow, t, sDir(i), mode, linkType, sNewFacilityType)
                                        sCap = getLaneCap2(pRow, pMRow, sDir(i), CStr(lType(l)))

                                        sOldFacilityType = getOldFacilityType2(pFeat)
                                        sFFS = getSpeedLimit2(pRow, pMRow, CStr(lType(l)), sDir(i), sOldFacilityType, pFeat)
                                        sFacType = getNewFacilityType2(pFeat)
                                        strTOD(t) = strTOD(t) + " " + sVDF + " " + sCap + " " + sFFS + " " + sFacType
                                    End If
                                    'SEC--took out SFFS string (this attribute no longer exists in ModeAttributes
                                    'strTOD(t) = strTOD(t) + " " + sVDF + " " + sCap + " " + sFacType
                                    If Not wString = "" Then
                                        If fVerboseLog Then WriteLogLine("wstring=" & wString)
                                        PrintLine(t + 1, wString)
                                    End If

                                    'SEC 010809: changed the following code so that reversibles run in both directions in the mid-day
                                    If l = 0 Then
                                        'for oneway with a walk mode, make a walk link in the opposite direction
                                        If buildOneWayReverseWalkLink Then

                                            PrintLine(t + 1, onewayWalkLink)

                                        End If
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


                                astring = ""
                                wString = ""


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
            Next
            'end of loop through every feature in intermediate edge shapefile
            'now add an additional links from splitting due to weave link case 2
        Catch ex As Exception
            MessageBox.Show(ex.ToString)

        End Try


        If g_ExportShapefiles = True Then



            'only need to export one junction file for all times of day
            Dim outJunctions As IFeatureClass
            'Export ScenarioJunctions
            outJunctions = FeatureClassToShapefile(pathnameN, pFeatlayerJunctions, "junctions")
            addField(outJunctions, "IsZone", esriFieldType.esriFieldTypeInteger)

            'get a dictionary of ModeAttribute Records for HOV only
            '**********Need to make sure its getting atts from right place- projects!!!!!!!!!!!!!
            getEdgeAtts(m_edgeShp, pTblMode, pTblPrjEdge, "PSRCEdgeID", dctHOVModeAtts, "HOV_I > 0 AND HOV_J > 0")



            Dim myWorkspaceFactory As IWorkspaceFactory = New ShapefileWorkspaceFactoryClass()
            Dim workspace As IWorkspace = myWorkspaceFactory.OpenFromFile(pathnameN, 0)
            Dim workspaceEdit As IWorkspaceEdit = CType(workspace, IWorkspaceEdit)
            workspaceEdit.StartEditing(False)
            calcIsZone(outJunctions, m_Offset)
            '8/22/12 Stefan
            'Adds Weave/HOV junctions to exported ScenarioJunctions
            'dctWeaveJunctions and dctTODWeaveLinks are populated during build file creation above. All other logic is new and much of it is 
            'duplicitive. A major overhaul of the code should look into combining/optimizing code/logic. For example, 
            'checking for HOV by TOD is done at least 3 times right now. Once for build, once for Export and once for 
            'transit. 
            AddWeaveJunctions(outJunctions, dctWeaveJunctions)

            workspaceEdit.StopEditOperation()
            workspaceEdit.StopEditing(True)




            Dim outEdges(5) As IFeatureClass

            Dim dctHOVByTOD As New Dictionary(Of Long, edgeHOVAtts)
            'Time of Day HOV Attributes. In the code below, the term weave is used to for the portion of the system that
            'provides network access to the HOV system while HOV is used for the actual HOV edges/links.
            dctHOVByTOD = todHOVatts(m_edgeShp, dctHOVModeAtts)
            'Fill a dictionary with HOV polylines and attributes using the edgeHOVAtts structure. Also get weavelinks
            'associated with HOV system. 
            Dim dctHOVEdges As Dictionary(Of edgeHOVAtts, ESRI.ArcGIS.Geometry.Polyline)
            dctHOVEdges = New Dictionary(Of edgeHOVAtts, ESRI.ArcGIS.Geometry.Polyline)
            'Returns HOV edges adn popultes dctExportWeaveLinks by passing in dctHOVbyTOD
            dctHOVEdges = getHOVEdges2(m_edgeShp, dctExportWeaveLinks, dctHOVByTOD, m_Offset)

            'Populate a dictionary of weave polylines and their I & J node attributes
            Dim dctWeaveEdges As New Dictionary(Of TwoValue, ESRI.ArcGIS.Geometry.Polyline)
            dctWeaveEdges = getWeaveFeatures(dctExportWeaveLinks, outJunctions)
            'workspaceEdit.StartEditing(False)


            'for each time period
            For t = 0 To 4
                'Export ScenarioEdges for all times of day
                outEdges(t) = FeatureClassToShapefile(pathnameN, pFeatlayerIE, "edges_" & t)
                'add some fields
                addField(outEdges(t), "NewINode", esriFieldType.esriFieldTypeInteger, 20)
                addField(outEdges(t), "NewJNode", esriFieldType.esriFieldTypeInteger, 20)
                'Populate new I & J Fi
                CalcGP_IJNOdes(outEdges(t), dctReservedNodes, t, m_Offset)
                workspaceEdit.StartEditing(False)
                'add weave edges to the exported scenario edges
                AddWeaveEdges(outEdges(t), dctWeaveEdges, dctTODWeaveLinks(t), t)
                'add HOV edges tot he exported scearnio edges
                AddHOVEdges(outEdges(t), dctHOVEdges, t, m_Offset)
                workspaceEdit.StopEditOperation()
                workspaceEdit.StopEditing(True)
            Next t

            

        End If
        For t = 0 To 5
            'exporting starts here:
            'export test


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

        dctEdges.Clear()
        dctMAtt.Clear()
        dctReservedNodes.Clear()

        dctSplit.Clear()
        For l = 1 To 3
            dctWeaveNodes(l).Clear()
            dctWeaveNodes2(l).Clear()
            dctWeaveNodes(l) = Nothing
            dctWeaveNodes2(l) = Nothing
        Next l
        dctSplit = Nothing
        dctEdges = Nothing
        dctMAtt = Nothing
        dctReservedNodes = Nothing

        '09282011 added following to create vehicle count attribute file
        Dim strVehicleCounts As String
        strVehicleCounts = pathnameN + "\" + filenameN + "AWDTCounts.txt"

        FileOpen(7, strVehicleCounts, OpenMode.Output)

        Dim pVehicleCounts As IFeatureLayer
        pVehicleCounts = get_FeatureLayer3(m_layers(25), m_App)
        'loop through each row
        Dim p_vcFeature As IFeature
        Dim p_vcFCursor As IFeatureCursor
        Dim tempVCString As String
        p_vcFCursor = pVehicleCounts.FeatureClass.Search(Nothing, False)
        p_vcFeature = p_vcFCursor.NextFeature
        Dim indexInode As Long
        Dim indexJnode As Long
        Dim indexAWDT As Long
        indexInode = p_vcFeature.Fields.FindField("INode")
        indexJnode = p_vcFeature.Fields.FindField("JNode")
        indexAWDT = p_vcFeature.Fields.FindField("trucks")


        Do Until p_vcFeature Is Nothing
            tempString = CType(p_vcFeature.Value(indexInode) + m_Offset, String) & " " & CType(p_vcFeature.Value(indexInode) + m_Offset, String) & " " & CType(p_vcFeature.Value(indexAWDT), String)
            PrintLine(7, tempString)
            p_vcFeature = p_vcFCursor.NextFeature
        Loop

        FileClose(6)








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

        If Not dctEdges Is Nothing Then dctEdges.Clear()
        If Not dctMAtt Is Nothing Then dctMAtt.Clear()
        dctReservedNodes.Clear()
        For l = 1 To 3
            dctWeaveNodes(l).Clear()
            dctWeaveNodes2(l).Clear()
            dctReservedNodes(l) = Nothing
            dctWeaveNodes(l) = Nothing
        Next l
        dctEdges = Nothing
        dctMAtt = Nothing


    End Sub
    Public Sub getMAtts2(ByVal pFClsEdge As IFeatureClass, ByVal pTbl As ITable, ByVal EdgeIDFld As String, ByVal dctEdge As Dictionary(Of Long, Feature), ByVal dctMAtt As Dictionary(Of Long, Row), Optional ByVal whereClause As String = "")

        '[062007] jaf:  appears to populate dctEdge and dctMAtt with ScenarioEdge and related modeAttributes records respectively

        Dim dctEdge2 As New Dictionary(Of Long, Long)
        'dctEdge2 = New Dictionary
        If dctEdge Is Nothing Then dctEdge = New Dictionary(Of Long, Feature)
        If dctMAtt Is Nothing Then dctMAtt = New Dictionary(Of Long, Row)


        Dim pDS As IDataset
        pDS = pTbl
        Debug.Print("getMAtt: " & pDS.Name & " start at " & Now())

        Dim pCs As ICursor, pRow As IRow
        Dim pFCS As IFeatureCursor, pFt As IFeature
        Dim sEdgeID As String, sEdgeOID As String
        Dim fld As Long
        Dim pQueryFilter As IQueryFilter = New QueryFilter
        pQueryFilter.WhereClause = whereClause


        If whereClause = "" Then
            fld = pFClsEdge.FindField(EdgeIDFld)
            pFCS = pFClsEdge.Search(Nothing, False)
        Else
            pFCS = pFClsEdge.Search(pQueryFilter, False)
        End If
        fld = pFClsEdge.FindField(EdgeIDFld)
        'pFCS = pFClsEdge.Search(Nothing, False)
        pFt = pFCS.NextFeature
        Do Until pFt Is Nothing
            sEdgeOID = CStr(pFt.OID)
            sEdgeID = CStr(pFt.Value(fld))

            If Not dctEdge2.ContainsKey(sEdgeID) Then dctEdge2.Add(sEdgeID, sEdgeOID)
            '[050407] hyu
            'If Not dctEdge.Exists(sEdgeOID) Then dctEdge.Add sEdgeOID, sEdgeOID
            If Not dctEdge.ContainsKey(sEdgeOID) Then dctEdge.Add(sEdgeOID, pFt)
            pFt = pFCS.NextFeature
        Loop

        Debug.Print("finish get Edges: " & Now())

        fld = pTbl.FindField(EdgeIDFld)
        pCs = pTbl.Search(Nothing, False)
        pRow = pCs.NextRow
        Do Until pRow Is Nothing

            If Not IsDBNull(pRow.Value(fld)) Then
                sEdgeID = CStr(pRow.Value(fld))
                If dctEdge2.ContainsKey(sEdgeID) Then
                    If Not dctMAtt.ContainsKey(sEdgeID) Then dctMAtt.Add(sEdgeID, pRow)
                End If

            End If
            pRow = pCs.NextRow
        Loop

        Debug.Print("getMAtt: " & pDS.Name & " end at " & Now())

        dctEdge2.Clear()
        dctEdge2 = Nothing
        pCs = Nothing
        '     pEnumV = Nothing
        '     pDataStat = Nothing
        pDS = Nothing

    End Sub
    Private Function getAllLanes2(ByVal pRow As IRow, ByVal pMRow As IRow, ByVal sType As String) As Double
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
                    dLanes = dLanes + getLanes2(pRow, pMRow, sType, sDir(i), sTpd(j))
                Next j
            Next i
        Else
            'check for TR!!!!!
            For i = 0 To 1
                dLanes = dLanes + getLanes2(pRow, pMRow, sType, sDir(i))
            Next i
        End If
        getAllLanes2 = dLanes
    End Function
    Private Function getLanes2(ByVal pRow As IRow, ByVal pMRow As IRow, ByVal sType As String, ByVal direction As String, Optional ByVal tpd As String = "", Optional ByVal dctHOVAtts As Dictionary(Of Long, edgeHOVAtts) = Nothing) As Double
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


        getLanes2 = dLanes





    End Function
    Private Function getNewFacilityType2(ByVal pFeat As IFeature) As String
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

        getNewFacilityType2 = CStr(intFacType)
    End Function
    Private Function getIJNodes2(ByVal pFeat As IFeature, ByVal IJorJI As String, ByVal lOffset As Long, ByVal dctRNode As Dictionary(Of Long, Long)) As String
        'dctRNode is the dictionary of reserved nodes
        Dim fldI As String
        Dim fldJ As String
        Dim str As String
        Dim nodeI As Long, nodeJ As Long
        fldI = pFeat.Fields.FindField(g_INode)
        fldJ = pFeat.Fields.FindField(g_JNode)
        nodeI = pFeat.Value(fldI)
        nodeJ = pFeat.Value(fldJ)
        If dctRNode.ContainsKey(CStr(nodeI)) Then
            nodeI = dctRNode.Item(CStr(nodeI))
        Else
            nodeI = nodeI + lOffset
        End If

        If dctRNode.ContainsKey(CStr(nodeJ)) Then
            nodeJ = dctRNode.Item(CStr(nodeJ))
        Else
            nodeJ = nodeJ + lOffset
        End If

        If IJorJI = "IJ" Then
            str = CStr(nodeI) & " " & CStr(nodeJ)
        Else
            str = CStr(nodeJ) & " " & CStr(nodeI)
        End If
        getIJNodes2 = str
    End Function

    Private Function getModes2(ByVal pFeat As IFeature, ByVal sType As String) As String
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
                    getModes2 = "ashi"
                Case "TR"  'tr
                    getModes2 = "ab"
                Case "HOV"  'hov

                Case "TK"  'tk
                    getModes2 = "auv"
            End Select
        ElseIf sType = "TR" Then
            getModes2 = "b"

            'supplied value of modes string
        Else
            getModes2 = CStr(pFeat.Value(Tindex))
        End If

    End Function

    Private Function getLength2(ByVal strLayerPrefix As String, ByVal pFeat As IFeature) As String
        'jaf--this block writes the modes string, which has defaults if not supplied
        'jaf--(column modes)
        Dim length As Double



        Dim Tindex As Long
        If strLayerPrefix = "SDE" Then
            Tindex = pFeat.Fields.FindField("Shape.len") 'SDE
            If Tindex = -1 Then
                Tindex = pFeat.Fields.FindField("Shape.STlength()")

            End If
        Else
            Tindex = pFeat.Fields.FindField("Shape_Length") 'PGDB
        End If

        '[040406] hyu: Length info can be null
        If IsDBNull(pFeat.Value(Tindex)) Then
            length = 0.015


        Else
            length = pFeat.Value(Tindex)


            length = length * 0.00018939393939
            length = FormatNumber(length, 3)
            If length < 0.01 Then
                length = 0.01
            End If


            ' obj = CDec(length)
        End If

        getLength2 = CStr(length)

    End Function

    Public Function getAttributesRow2(ByVal pFeat As IFeature, ByVal pMRow As IRow, ByVal pTblLine As ITable, ByVal pEvtLine As ITable) As IRow
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


        getAttributesRow2 = pRow
    End Function
    Private Function getSplitLink2(ByVal pFeat As IFeature, ByVal dctSplitLinks As Dictionary(Of String, String), ByVal dctEmme2Nodes As Dictionary(Of Long, Long))
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

        INode = pFeat.Value(fldI)
        JNode = pFeat.Value(fldJ)
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

        oneWay = pFeat.Value(fldOneWay)

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
    Private Function getFlag2(ByVal sType As String, ByVal direction As String) As String
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
        getFlag2 = s
    End Function

    Private Function getLinkType2(ByVal pFeat As IFeature) As String
        'jaf--this block writes the linktype number (between 1-99, used to flag screen lines)
        'jaf--(column linktype)
        'jaf--default to 90 for now, we'll have to figure out how to capture this in GeoDB

        Dim Tindex As Long
        Tindex = pFeat.Fields.FindField("LinkType")
        If (IsDBNull(pFeat.Value(Tindex)) = False And pFeat.Value(Tindex) > 0) Then
            getLinkType2 = pFeat.Value(Tindex)
        Else
            getLinkType2 = "90"
        End If
    End Function
    Private Function GetWeaveLinks2(ByVal pFeat As IFeature, ByVal pRow As IRow, ByVal pMRow As IRow, ByVal dctWLinks As Dictionary(Of String, String), ByVal dctSplitLinks As Dictionary(Of String, String), _
       ByVal dctEmme2Nodes As Dictionary(Of Long, Long), ByVal lType As String, _
       ByVal direction As String, ByVal tpd As String, ByVal lanes As Double, ByRef wNodes As String, ByRef astring As String, ByRef wString As String, ByVal dctTODWeaveLinks As Dictionary(Of Long, Long))
        'This function returns lane attributes for weave links

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

        'modeWeave = "ahijstuvbw"
        modeWeave = "ashijtuvbwledmgp"

        iwIndex = pFeat.Fields.FindField(lType & "_I")
        jwIndex = pFeat.Fields.FindField(lType & "_J")
        Try


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

            'index = pFeat.Fields.FindField("LinkType")
            'now assign defaults to weave: length, modes, link type, lanes
            'Leng = getWeaveLen(lType) / 5280
            '    lanes = getLanes(pRow, pMRow, lType, Direction, tpd)

            If lanes = 0 Then Exit Function

            If fVerboseLog Then WriteLogLine("leng of weave" + CStr(Leng))

            'wDefault holds the weave attributes:
            wDefault = wDefault + " " + "0.01" + " " + modeWeave + " 90 1 10 2000 0 0"





            iIndex = pFeat.Fields.FindField("INode")
            jIndex = pFeat.Fields.FindField("JNode")
            INode = pFeat.Value(iIndex)
            JNode = pFeat.Value(jIndex)
            index = pFeat.Fields.FindField("UseEmmeN")

            If IsDBNull(pFeat.Value(index)) Then
                INode = INode + m_Offset
                JNode = JNode + m_Offset
            ElseIf pFeat.Value(index) = 0 Then
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

            ' Dim sFT As String
            'Dim laneAtts As New clsModeAttributes(pMRow)
            'sFT = pFeat.Value(pFeat.Fields.FindField("NewFacilityType"))
            'ijtemp = "1 1800 60 " & sFT
            'jitemp = "1 1800 60 " & sFT


            'Only gives WeaveLinks if it has an HOV attribute in the IJ direction for that time of day!!!!!!Need to check to see if
            Dim wLink As String
            wLink = CStr(iwNode) + " " + CStr(INode)

            If Not dctTODWeaveLinks.ContainsKey(INode) Then
                dctTODWeaveLinks.Add(INode, iwNode)
            End If


            'this is where to add weave links by TOD

            If Not dctWLinks.ContainsKey(wLink) And Not dctSplitLinks.ContainsKey(wLink) Then
                If wString = "" Then
                    wString = "a " + CStr(iwNode) + " " + CStr(INode) + " " + wDefault
                Else
                    wString = wString + vbNewLine + "a " + CStr(iwNode) + " " + CStr(INode) + " " + wDefault
                End If

                If astring = "" Then
                    astring = CStr(iwNode) + " " + CStr(INode) + wID_Type
                Else
                    astring = astring + vbNewLine + CStr(iwNode) + " " + CStr(INode) + wID_Type
                End If

                'now write opposite
                wString = wString + vbNewLine + "a " + CStr(INode) + " " + CStr(iwNode) + " " + wDefault
                astring = astring + vbNewLine + CStr(INode) + " " + CStr(iwNode) + wID_Type
                dctWLinks.Add(wLink, wLink)
            End If

            wLink = CStr(jwNode) + " " + CStr(JNode)
            If Not dctTODWeaveLinks.ContainsKey(JNode) Then
                dctTODWeaveLinks.Add(JNode, jwNode)
            End If

            If Not dctWLinks.ContainsKey(wLink) And Not dctSplitLinks.ContainsKey(wLink) Then
                If wString = "" Then
                    wString = "a " + wLink + " " + wDefault
                Else
                    wString = wString + vbNewLine + "a " + wLink + " " + wDefault
                End If

                If astring = "" Then
                    astring = CStr(jwNode) + " " + CStr(JNode) + wID_Type
                Else
                    astring = astring + vbNewLine + CStr(jwNode) + " " + CStr(JNode) + wID_Type
                End If

                'now write opposite
                wString = wString + vbNewLine + "a " + CStr(JNode) + " " + CStr(jwNode) + " " + wDefault
                astring = astring + vbNewLine + CStr(JNode) + " " + CStr(jwNode) + wID_Type
                dctWLinks.Add(wLink, wLink)
            End If

            If fVerboseLog Then WriteLogLine(wNodes)
            If fVerboseLog Then WriteLogLine(wString)
        Catch ex As Exception
            MessageBox.Show(ex.ToString)

        End Try

    End Function
    Private Function getVDF2(ByVal pRow As IRow, ByVal pMRow As IRow, ByVal t As Integer, ByVal direction As String, ByVal mode As String, ByVal linkType As String, ByVal FacilityType As String) As String
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
        getVDF2 = sVDF
    End Function
    Private Function getLaneCap2(ByVal pRow As IRow, ByVal pMRow As IRow, ByVal direction As String, ByVal sType As String) As String
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


        If pRow.Value(Tindex) > -1 Then
            sCap = CStr(pRow.Value(Tindex))
        Else
            If sType = "GP" Or sType = "HOV" Then 'If l < 2 Then
                Tindex = pMRow.Fields.FindField(direction + "LaneCap" + sType)
            Else
                Tindex = pMRow.Fields.FindField(direction + "LaneCap" + "GP")
            End If

            If Not IsDBNull(pMRow.Value(Tindex)) Then
                sCap = CStr(pMRow.Value(Tindex))
            Else
                sCap = "1800"
            End If
        End If
        getLaneCap2 = sCap

    End Function
    Private Function getSpeedLimit2(ByVal pRow As IRow, ByVal pMRow As IRow, ByVal sType As String, ByVal direction As String, ByVal FacilityType As String, ByVal pFeature As IFeature) As String
        Dim Tindex As Long
        Dim strSpeedLimit As String
        Dim strName As String
        Dim strSounderName As String
        Dim dblLength As Double
        Dim dblSpeed As Double
        Dim dblSpeed2 As Double
        Dim lngSpeed As Long
        Tindex = pRow.Fields.FindField(direction + "SpeedLimit")

        Dim indexLengthField As Long
        If strLayerPrefix = "SDE" Then
            indexLengthField = pFeature.Fields.FindField("Shape.len") 'SDE
            If indexLengthField = -1 Then
                indexLengthField = pFeature.Fields.FindField("Shape.STlength()")

            End If
        Else
            indexLengthField = pFeature.Fields.FindField("Shape_Length") 'PGDB
        End If
        dblLength = pFeature.Value(indexLengthField)


        If FacilityType = "14" Or FacilityType = "15" Then
            strName = pFeature.Value(pFeature.Fields.FindField("Fullname"))
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
        getSpeedLimit2 = strSpeedLimit
    End Function
    Private Function getOldFacilityType2(ByVal pFeat As IFeature) As String

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

        getOldFacilityType2 = CStr(intFacType)
    End Function

    Public Sub createWeaveLink2(ByVal lType As String, ByVal pFeat As IFeature, ByVal pAttRow As IRow, _
       ByVal wString As String, ByVal wNodes As String, ByVal astring As String, _
       ByVal dctEmme2Nodes As Dictionary(Of Long, Long), ByVal dctWNodes As Dictionary(Of String, Long), ByVal dctWNodes2 As Dictionary(Of Long, Long), _
       ByVal dctJcts As Dictionary(Of Long, Feature), ByVal dctEdges As Dictionary(Of Long, Feature), ByRef nodeString As String, ByVal dctNewJunctions As Dictionary(Of Long, ESRI.ArcGIS.Geometry.Point))
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
        Dim xxxxx As Integer

        'modeWeave = "ashijtuvbw"
        modeWeave = "ashijtuvbwledmgp"

        Try



            If pFeat.Value(2) = 204273 Then
                xxxxx = 0
            End If
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
            If IsDBNull(pFeat.Value(index)) Then
                INode = INode + m_Offset
                JNode = JNode + m_Offset

            ElseIf pFeat.Value(index) = 0 Then
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
            Dim sVDF As Integer
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

                    If Not dctWNodes.ContainsKey(CType(iwNode, String)) Then
                        '[051507]hyu: if the weavenode is a junction, it's the split point for case 2 during weave link creation
                        If dctJcts.ContainsKey(CType(iwNode, Long)) Then
                            pJFeat = dctJcts.Item(CType(iwNode, Long))
                            pPoint = pJFeat.ShapeCopy
                            dctNewJunctions.Add(iwNode, pPoint)
                            updateTransitLines(pPoint)
                        Else
                            pJFeat = dctJcts.Item(CType(INode, Long))
                            pPoint = getWeavenode(pJFeat, lType)
                            dctNewJunctions.Add(iwNode, pPoint)

                            If nodeString = "" Then
                                nodeString = "a " + CStr(iwNode) + " " + CStr(pPoint.X) + " " + CStr(pPoint.Y) + " 0 0 0 0"
                            Else
                                nodeString = nodeString + vbCrLf + "a " + CStr(iwNode) + " " + CStr(pPoint.X) + " " + CStr(pPoint.Y) + " 0 0 0 0"
                            End If
                        End If

                        dctWNodes.Add(CType(iwNode, String), iwNode - m_Offset)
                        If Not dctWNodes2.ContainsKey(CType(INode, Long)) Then dctWNodes2.Add(CType(INode, Long), CType(iwNode, Long))
                    End If
                End If

                If Not dctWNodes.ContainsKey(CType(jwNode, String)) Then
                    '[051507]hyu: if the weavenode is a junction, it's the split point for case 2 during weave link creation
                    If dctJcts.ContainsKey(CType(jwNode, Long)) Then
                        pJFeat = dctJcts.Item(CType(jwNode, Long))
                        pPoint = pJFeat.ShapeCopy
                        dctNewJunctions.Add(jwNode, pPoint)

                        updateTransitLines(pPoint)
                    Else
                        pJFeat = dctJcts.Item(CType(JNode, Long))
                        pPoint = getWeavenode(pJFeat, lType)
                        dctNewJunctions.Add(jwNode, pPoint)
                        If nodeString = "" Then
                            nodeString = "a " + CStr(jwNode) + " " + CStr(pPoint.X) + " " + CStr(pPoint.Y) + " 0 0 0 0"
                        Else
                            nodeString = nodeString + vbCrLf + "a " + CStr(jwNode) + " " + CStr(pPoint.X) + " " + CStr(pPoint.Y) + " 0 0 0 0"
                        End If

                    End If

                    dctWNodes.Add(CType(jwNode, String), CType(jwNode, Long) - m_Offset)
                    If Not dctWNodes2.ContainsKey(CType(JNode, Long)) Then dctWNodes2.Add(CType(JNode, Long), CType(jwNode, Long))
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
                        If dctWNodes2.ContainsKey(CType(INode, Long)) = True Then
                            pFeat.Value(pFeat.Fields.FindField(lType & "_I")) = dctWNodes2.Item(CType(INode, Long)) - m_Offset
                            bWeaveNodeExists = True
                            iwNode = dctWNodes2.Item(CType(INode, Long))
                        End If
                    Else
                        If dctWNodes2.ContainsKey(CType(JNode, Long)) Then
                            pFeat.Value(pFeat.Fields.FindField(lType & "_J")) = dctWNodes2.Item(CType(JNode, Long)) - m_Offset
                            bWeaveNodeExists = True
                            If dctWNodes.ContainsKey(CType(JNode, String)) Then
                                jwNode = dctWNodes.Item(CType(JNode, String))
                            Else
                                jwNode = 0
                            End If

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


                            pJFeat = dctJcts.Item(CType(INode, Long))
                        ElseIf j = 1 Then
                            tempnode = JNode
                            pJFeat = dctJcts.Item(CType(JNode, Long))
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
                            pJFeat = dctJcts.Item(CType(JNode, Long))
                        ElseIf j = 1 Then
                            pJFeat = dctJcts.Item(CType(INode, Long))
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
                            CreateWeaveNode2(pFeat, tempnode, j, lType, LargestJunct + m_Offset, dctJcts, dctWNodes, dctWNodes2, nodeString, dctNewJunctions)
                        Else '>1 other edges intersecting
                            pFC = m_edgeShp.Search(pFilter, False)
                            pefeat = pFC.NextFeature
                            '                    If Not UpdateWeaveNode(pFeat, pefeat, j, ltype) Then
                            'the intersecting edge has no weave nodes

                            Do Until pefeat Is Nothing
                                bCanSplit = CanSplit2(pFeat, pefeat, pAttRow, lType)
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
                                CreateWeaveNode2(pFeat, tempnode, j, lType, LargestJunct + m_Offset, dctJcts, dctWNodes, dctWNodes2, nodeString, dctNewJunctions)
                            End If

                        End If
                    End If  'bWeavenodeExists=true
                Next j
            End If  '(Not IsNull(pFeat.value(iwIndex)) And pFeat.value(iwIndex) > 0)
            '    If fVerboseLog Then WriteLogLine "finished createWeaveLink call with ptstring=" & ptstring
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.createWeaveLink")
            MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.createWeaveLink")
        End Try

        Exit Sub

        'eh:
        'CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.createWeaveLink")
        'MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.createWeaveLink")
        'Print #9, Err.Description, vbInformation, "createWeaveLink"
        '   If pWSEdit.IsBeingEdited Then
        '      pWSEdit.AbortEditOperation
        '      pWSEdit.StopEditing False
        '   End If

    End Sub
    Private Function CanSplit2(ByVal pFeat As IFeature, ByVal pFeatInt As IFeature, ByVal pAttRow As IRow, ByVal lType As String) As Boolean
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

        CanSplit2 = False
        If Not (IsDBNull(pFeatInt.Value(pFeatInt.Fields.FindField("SPLIT" & lType)))) Then Exit Function

        fldFacType = pFeat.Fields.FindField("FacilityType")
        fldModes = pFeat.Fields.FindField("Modes")

        If pFeat.Value(fldFacType) >= 6 And pFeatInt.Value(fldFacType) <> 1 And pFeatInt.Value(fldFacType) <> 5 Then
            Select Case pFeatInt.Value(fldModes)
                Case "wk", "r", "f"
                    CanSplit2 = False
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
                    iLanes = iLanes + pRow.Value(index)
                    index = pRow.Fields.FindField("JILanes" + laneFld)
                    iLanes = iLanes + pRow.Value(index)
                Next i
            Else
                laneFld = CStr(lType)
                index = pRow.Fields.FindField("IJLanes" + laneFld)
                iLanes = pRow.Value(index)
                index = pRow.Fields.FindField("JILanes" + laneFld)
                iLanes = iLanes + pRow.Value(index)
            End If

            If iLanes = 0 Then CanSplit2 = True
        End If
    End Function
    Private Function CreateWeaveNode2(ByVal pFeat As IFeature, ByVal nodeID As Long, ByVal IorJ As Integer, ByVal WeaveType As String, ByVal lWeaveNode As Long, _
    ByVal dctJcts As Dictionary(Of Long, Feature), ByVal dctWNodes As Dictionary(Of String, Long), ByVal dctWNodes2 As Dictionary(Of Long, Long), ByRef nodeString As String, ByVal dctNewJunctions As Dictionary(Of Long, ESRI.ArcGIS.Geometry.Point))

        Dim pPoint As IPoint
        Dim pJFeat As IFeature
        Dim fld As Long

        If IorJ = 0 Then
            fld = pFeat.Fields.FindField(WeaveType & "_I")
        Else
            fld = pFeat.Fields.FindField(WeaveType & "_J")
        End If

        pFeat.Value(fld) = lWeaveNode - m_Offset
        pFeat.Store()


        pJFeat = dctJcts.Item(CType(nodeID, Long))



        pPoint = getWeavenode(pJFeat, WeaveType)
        'add to dict
        If lWeaveNode = 194237 + 4000 Then
            MessageBox.Show("found it")

        End If
        If Not dctNewJunctions.ContainsKey(lWeaveNode) Then
            dctNewJunctions.Add(lWeaveNode, pPoint)
        End If


        If nodeString = "" Then
            nodeString = "a " + CStr(lWeaveNode) + " " + CStr(pPoint.X) + " " + CStr(pPoint.Y) + " 0 0 0 0"
        Else
            nodeString = nodeString + vbCrLf + "a " + CStr(lWeaveNode) + " " + CStr(pPoint.X) + " " + CStr(pPoint.Y) + " 0 0 0 0"
        End If
        If Not dctWNodes.ContainsKey(CType(lWeaveNode, String)) Then
            dctWNodes.Add(CType(lWeaveNode, String), lWeaveNode - m_Offset)
        End If

        Try

            If Not dctWNodes2.ContainsKey(CType(nodeID, Long)) Then
                dctWNodes2.Add(CType(nodeID, Long), CType(lWeaveNode, Long))
            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        pPoint = Nothing
        '    Set pJFeat = Nothing
    End Function



    Public Function get_FeatureLayer3(ByVal featClsname As String, ByVal app As IApplication) As IFeatureLayer
        'returns the feature class stated in featClsname
        'on error GoTo eh

        Dim _doc As IMxDocument
        Dim _activeView As IActiveView
        Dim _map As IMap

        Dim pEnumLy As IEnumLayer


        _doc = app.Document
        _activeView = _doc.ActiveView
        _map = _activeView.FocusMap






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
                        get_FeatureLayer3 = pFLy
                        Exit Do

                    End If
                End If
            End If
            pLy = pEnumLy.Next
        Loop

        If Not get_FeatureLayer3 Is Nothing Then Exit Function

        Dim pFeatLayer As IFeatureLayer
        pFeatLayer = New FeatureLayer
        Dim pfeatcls As IFeatureClass
        pfeatcls = getFeatureClass(featClsname)
        pFeatLayer.FeatureClass = pfeatcls

        get_FeatureLayer3 = pFeatLayer
        If pFeatLayer Is Nothing Then
            MsgBox("Error: did not find the feature class " + featClsname)
        End If

        Exit Function

eh:
        MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.get_FeatureLayer")

    End Function
    Private Function FeatureClassToShapefile(ByVal targetWorkspacePath As String, ByVal sourceFL As IFeatureLayer, ByVal targetFCName As String, Optional ByVal queryString As String = "Nothing") As FeatureClass
        'Exports FeatureClass to Shapefile

        Dim targetWorkspaceFactory As IWorkspaceFactory = New ShapefileWorkspaceFactoryClass()
        Dim pDS As IDataset
        pDS = sourceFL
        Dim sourceWorkspace As IWorkspace = pDS.Workspace

        Dim targetWorkspace As IWorkspace = targetWorkspaceFactory.OpenFromFile(targetWorkspacePath, 0)

        ' Cast the workspaces to the IDataset interface and get name objects.
        Dim sourceWorkspaceDataset As IDataset = CType(sourceWorkspace, IDataset)
        Dim targetWorkspaceDataset As IDataset = CType(targetWorkspace, IDataset)
        Dim sourceWorkspaceDatasetName As IName = sourceWorkspaceDataset.FullName
        Dim targetWorkspaceDatasetName As IName = targetWorkspaceDataset.FullName
        Dim sourceWorkspaceName As IWorkspaceName = CType(sourceWorkspaceDatasetName, IWorkspaceName)
        Dim targetWorkspaceName As IWorkspaceName = CType(targetWorkspaceDatasetName, IWorkspaceName)

        ' Create a name object for the shapefile and cast it to the IDatasetName interface.
        Dim sourceFeatureClassName As IFeatureClassName = New FeatureClassNameClass()
        Dim sourceDatasetName As IDatasetName = CType(sourceFeatureClassName, IDatasetName)
        sourceDatasetName.Name = sourceFL.FeatureClass.AliasName
        sourceDatasetName.WorkspaceName = sourceWorkspaceName

        ' Create a name object for the FGDB feature class and cast it to the IDatasetName interface.
        Dim targetFeatureClassName As IFeatureClassName = New FeatureClassNameClass()
        Dim targetDatasetName As IDatasetName = CType(targetFeatureClassName, IDatasetName)
        targetDatasetName.Name = targetFCName
        targetDatasetName.WorkspaceName = targetWorkspaceName

        ' Open source feature class to get field definitions.
        Dim sourceName As IName = CType(sourceFeatureClassName, IName)
        Dim sourceFeatureClass As IFeatureClass = CType(sourceName.Open(), IFeatureClass)

        ' Create the objects and references necessary for field validation.
        Dim fieldChecker As IFieldChecker = New FieldCheckerClass()
        Dim sourceFields As IFields = sourceFeatureClass.Fields
        Dim targetFields As IFields = Nothing
        Dim enumFieldError As IEnumFieldError = Nothing

        ' Set the required properties for the IFieldChecker interface.
        fieldChecker.InputWorkspace = sourceWorkspace
        fieldChecker.ValidateWorkspace = targetWorkspace

        ' Validate the fields and check for errors.
        fieldChecker.Validate(sourceFields, enumFieldError, targetFields)
        If Not enumFieldError Is Nothing Then
            ' Handle the errors in a way appropriate to your application.
            Console.WriteLine("Errors were encountered during field validation.")
        End If


        ' Find the shape field.
        Dim shapeFieldName As String = sourceFeatureClass.ShapeFieldName
        Dim shapeFieldIndex As Integer = sourceFeatureClass.FindField(shapeFieldName)
        Dim shapeField As IField = sourceFields.Field(shapeFieldIndex)

        ' Get the geometry definition from the shape field and clone it.
        Dim geometryDef As IGeometryDef = shapeField.GeometryDef
        Dim geometryDefClone As IClone = CType(geometryDef, IClone)
        Dim targetGeometryDefClone As IClone = geometryDefClone.Clone()
        Dim targetGeometryDef As IGeometryDef = CType(targetGeometryDefClone, IGeometryDef)

        ' Create a query filter to remove ramps, interstates and highways.
        Dim queryFilter As IQueryFilter = New QueryFilterClass()
        If queryString = "Nothing" Then
            queryFilter = Nothing
        Else
            queryFilter.WhereClause = queryString
        End If


        ' Create the converter and run the conversion.
        Dim featureDataConverter As IFeatureDataConverter = New FeatureDataConverterClass()
        Dim enumInvalidObject As IEnumInvalidObject = featureDataConverter.ConvertFeatureClass(sourceFeatureClassName, queryFilter, Nothing, targetFeatureClassName, targetGeometryDef, targetFields, "", 1000, 0)

        ' Check for errors.
        enumInvalidObject.Reset()
        Dim invalidObjectInfo As IInvalidObjectInfo = enumInvalidObject.Next()
        Do While Not invalidObjectInfo Is Nothing
            ' Handle the errors in a way appropriate to the application.
            Console.WriteLine("Errors occurred for the following feature: {0}", invalidObjectInfo.InvalidObjectID)
        Loop



        Dim shapefileWS As IFeatureWorkspace = targetWorkspace
        Dim myFC As IFeatureClass
        myFC = shapefileWS.OpenFeatureClass(targetFCName)
        Return myFC



    End Function
    Private Sub AddFieldToFeatureClass(ByVal featureClass As IFeatureClass, ByVal field As IField)

        Dim schemaLock As ISchemaLock = CType(featureClass, ISchemaLock)

        Try
            ' A try block is necessary, as an exclusive lock may not be available.
            schemaLock.ChangeSchemaLock(esriSchemaLock.esriExclusiveSchemaLock)

            ' Add the field.
            featureClass.AddField(field)
        Catch exc As Exception
            ' Handle this in a way appropriate to your application.
            Console.WriteLine(exc.Message)
        Finally
            ' Set the lock to shared, whether or not an error occurred.
            schemaLock.ChangeSchemaLock(esriSchemaLock.esriSharedSchemaLock)
        End Try

    End Sub

    Private Function CreateFieldsCollectionForTable() As IFields

        ' Create a new fields collection.
        Dim fields As IFields = New FieldsClass()

        ' Cast to IFieldsEdit to modify the properties of the fields collection.
        Dim fieldsEdit As IFieldsEdit = CType(fields, IFieldsEdit)

        ' Set the number of fields the collection will contain.
        fieldsEdit.FieldCount_2 = 1

        ' Create the ObjectID field.
        'Dim oidField As IField = New FieldClass()

        ' Cast to IFieldEdit to modify the properties of the new field.

        ' Create the text field.
        Dim textField As IField = New FieldClass()
        Dim textFieldEdit As IFieldEdit = CType(textField, IFieldEdit)
        textFieldEdit.Length_2 = 20 ' Only string fields require that you set the length.
        textFieldEdit.Name_2 = "NewINode"
        textFieldEdit.Type_2 = esriFieldType.esriFieldTypeInteger

        ' Add the new field to the fields collection.
        fieldsEdit.Field_2(1) = textField

        Return fields

    End Function
    
    
    Public Function StartEditing(ByVal workspaceToEdit As ESRI.ArcGIS.Geodatabase.IWorkspace) As Boolean
        'Get a reference to the editor.
        Dim uid As UID
        uid = New UIDClass()
        uid.Value = "esriEditor.Editor"
        Dim editor As IEditor
        editor = CType(m_App.FindExtensionByCLSID(uid), IEditor)

        'Check to see if a workspace is already being edited.
        If editor.EditState = esriEditState.esriStateNotEditing Then
            editor.StartEditing(workspaceToEdit)
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub addWeaveJunctions(ByVal featureClass As IFeatureClass, ByVal dctPoints As Dictionary(Of Long, ESRI.ArcGIS.Geometry.Point))
        'Adds New weave Junctions to the exported junction shapefile

        ' Create the feature buffer.
        Dim featureBuffer As IFeatureBuffer = featureClass.CreateFeatureBuffer()

        ' Create insert feature cursor using buffering.
        Dim featureCursor As IFeatureCursor = featureClass.Insert(True)

        ' All of the features to be created were installed by "B Pierce."
        Dim JunctIDFieldIndex As Integer = featureClass.FindField("PSRCjunctI")
        Dim sceneNodeFieldIndex As Integer = featureClass.FindField("Scen_Node")
        featureBuffer.Value(JunctIDFieldIndex) = 99
        Dim pair As KeyValuePair(Of Long, ESRI.ArcGIS.Geometry.Point)
        Dim geometry As IGeometry
        For Each pair In dctPoints
            geometry = pair.Value
            featureBuffer.Shape = geometry
            featureBuffer.Value(sceneNodeFieldIndex) = pair.Key
            featureCursor.InsertFeature(featureBuffer)


        Next

        Try
            featureCursor.Flush()
        Catch comExc As COMException

            ' Attempt to flush the buffer.
            ' Handle the error in a way appropriate to your application.
        Finally
            ' Release the cursor as it's no longer needed.
            Marshal.ReleaseComObject(featureCursor)
        End Try

    End Sub
   
    Public Sub AddHOVEdges(ByVal featureClass As IFeatureClass, ByVal dctHOVEdges As Dictionary(Of edgeHOVAtts, ESRI.ArcGIS.Geometry.Polyline), ByVal timeOfDay As Integer, ByVal nodeOffset As Long)
        'Adds HOV Edges to exported shapefiles

        ' Create the feature buffer.
        Dim featureBuffer As IFeatureBuffer = featureClass.CreateFeatureBuffer()

        ' Create insert feature cursor using buffering.
        Dim featureCursor As IFeatureCursor = featureClass.Insert(True)

        ' All of the features to be created were installed by "B Pierce."
        'Dim JunctIDFieldIndex As Integer = featureClass.FindField("PSRCjunctI")
        Dim edgeIDFieldIndex As Integer = featureClass.FindField("PSRCEdgeID")
        Dim facilityTypeIndex As Integer = featureClass.FindField("FacilityTy")
        Dim newINodeFieldIndex As Integer = featureClass.FindField("NewINode")
        Dim newJNodeFieldIndex As Integer = featureClass.FindField("NewJNode")
        Dim directionFieldIndex As Integer = featureClass.FindField("Direction")
        'featureBuffer.Value(JunctIDFieldIndex) = 99
        Dim pair As KeyValuePair(Of edgeHOVAtts, ESRI.ArcGIS.Geometry.Polyline)
        Dim pGeometry As IGeometry
        Dim myHOVAtts As edgeHOVAtts
        Try


            For Each pair In dctHOVEdges
                pGeometry = pair.Value
                myHOVAtts = pair.Key
                If myHOVAtts.todDirection(timeOfDay) <> -2 Then
                    featureBuffer.Shape = pGeometry
                    featureBuffer.Value(edgeIDFieldIndex) = myHOVAtts.psrcEdgeID
                    'Flag that this is an HOV
                    featureBuffer.Value(facilityTypeIndex) = 99
                    featureBuffer.Value(directionFieldIndex) = myHOVAtts.todDirection(timeOfDay)
                    'change 4000 later to the offset variable
                    featureBuffer.Value(newINodeFieldIndex) = myHOVAtts.newInode + nodeOffset
                    featureBuffer.Value(newJNodeFieldIndex) = myHOVAtts.newJNode + nodeOffset

                    featureCursor.InsertFeature(featureBuffer)
                End If

            Next pair

        Catch ex As Exception
            MessageBox.Show(ex.ToString)

        End Try




        Try
            featureCursor.Flush()
        Catch comExc As COMException

            ' Attempt to flush the buffer.
            ' Handle the error in a way appropriate to your application.
        Finally
            ' Release the cursor as it's no longer needed.
            Marshal.ReleaseComObject(featureCursor)
        End Try

    End Sub
   
    Public Sub getOffSet(ByVal pJFeat As IFeature, ByVal ltype As String, ByVal moveX As Double, ByVal moveY As Double)
        'the amount to +/- to point coordinates is calculated through a2 + b2 = len*
        'on error GoTo eh


        Dim Leng As Double
        Leng = getWeaveLen(ltype)
        moveX = (Math.Sqrt(Leng)) / 2
        moveX = Math.Sqrt(moveX)
        moveY = moveX 'move the same amount as x


        Exit Sub
    End Sub

    Public Function getHOVEdges2(ByVal edges As IFeatureClass, ByVal dctwLinks As Dictionary(Of Long, Long), ByVal dctedgeHOVAtts As Dictionary(Of Long, edgeHOVAtts), ByVal nodeOffset As Long) As Dictionary(Of edgeHOVAtts, ESRI.ArcGIS.Geometry.Polyline)
        'Populates dctwLinks:
        'key= GP Inode + offset, value= HOV JNode

        'And returns dict of EdgeAtts & HOV polylines





        Dim qFilter As IQueryFilter = New QueryFilter
        Dim edgeIDFilter As IQueryFilter = New QueryFilter
        Dim fCursor As IFeatureCursor
        Dim edgeIDCursor As ICursor
        Dim dctHOVEdges As New Dictionary(Of edgeHOVAtts, ESRI.ArcGIS.Geometry.Polyline)
        'get all edges that have HOV at some point during the day
        qFilter.WhereClause = "HOV_I > 0"

        Dim pQueryFilter As IQueryFilter
        Dim edgeIDIndex As Long
        Dim edgeID As Long

        Dim p_polyLine As IPolyline = New Polyline
        edgeIDIndex = edges.FindField("PSRCEdgeID")
        Dim iNodeIndex As Long
        Dim jNodeINdex As Long
        Dim iNode As Long
        Dim jNode As Long '
        Dim hovINodeIndex As Long
        Dim hovJNodeIndex As Long
        Dim hovINode As Long
        Dim hovJNode As Long

        iNodeIndex = edges.FindField("INOde")
        jNodeINdex = edges.FindField("JNode")
        hovINodeIndex = edges.FindField("HOV_I")
        hovJNodeIndex = edges.FindField("HOV_J")




        pQueryFilter = New QueryFilter
        fCursor = edges.Search(qFilter, False)
        Dim pFeature As IFeature
        pFeature = fCursor.NextFeature
        Dim myHOVAtts As edgeHOVAtts

        Dim trans2D As ITransform2D
        Do Until pFeature Is Nothing
            edgeID = pFeature.Value(edgeIDIndex)
            'return a dictionary with weave link pairs
            'edgeID = pFeature.Value(edgeIDIndex)
            ' edgeIDFilter.WhereClause = "PSRCEdgeID = " & edgeID
            'edgeIDCursor = modeAtts.Search(edgeIDFilter, True)

            iNode = pFeature.Value(iNodeIndex)
            jNode = pFeature.Value(jNodeINdex)
            iNode = iNode + nodeOffset
            jNode = jNode + nodeOffset
            hovINode = pFeature.Value(hovINodeIndex)
            hovJNode = pFeature.Value(hovJNodeIndex)
            hovINode = hovINode + nodeOffset
            hovJNode = hovJNode + nodeOffset
            If Not dctwLinks.ContainsKey(iNode) Then
                dctwLinks.Add(iNode, hovINode)
            End If
            If Not dctwLinks.ContainsKey(jNode) Then
                dctwLinks.Add(jNode, hovJNode)
            End If




            'return a dictionary with hov lines


            myHOVAtts = dctedgeHOVAtts.Item(edgeID)
            myHOVAtts.newInode = pFeature.Value(hovINodeIndex)
            myHOVAtts.newJNode = pFeature.Value(hovJNodeIndex)

            p_polyLine = pFeature.Shape
            trans2D = TryCast(p_polyLine, ITransform2D)

            trans2D.Move(g_xMove, g_yMove)
            dctHOVEdges.Add(myHOVAtts, p_polyLine)
            pFeature = fCursor.NextFeature
            trans2D = Nothing
            p_polyLine = Nothing








        Loop
        Return dctHOVEdges


    End Function
    Public Function getWeaveFeatures(ByVal dctWeavePairs As Dictionary(Of Long, Long), ByVal junctFeatures As IFeatureClass) As Dictionary(Of TwoValue, ESRI.ArcGIS.Geometry.Polyline)
        'returns a dictionary with 
        'key= Twovalue(I & J Node values), value = weave polylines
        Dim pair As KeyValuePair(Of Long, Long)
        Dim pGeometry As IGeometry
        Dim qFilter As IQueryFilter = New QueryFilter
        Dim fCursor As IFeatureCursor
        Dim iJunct As IFeature
        Dim jJunct As IFeature
        Dim iNodePoint As IPoint
        Dim jNodePoint As IPoint
        Dim pPolyLine As IPolyline
        Dim pPointCollection As IPointCollection
        'Dim myColl As Collection
        'myColl = New Collection

        Dim myDict As New Dictionary(Of TwoValue, ESRI.ArcGIS.Geometry.Polyline)
        Dim Key As TwoValue



        'need to keep i j values 

        Try

       
            For Each pair In dctWeavePairs
                qFilter = New QueryFilter
                qFilter.WhereClause = "Scen_Node = " & pair.Key
                fCursor = junctFeatures.Search(qFilter, False)
                iJunct = fCursor.NextFeature
                iNodePoint = iJunct.Shape

                qFilter = New QueryFilter
                qFilter.WhereClause = "Scen_Node = " & pair.Value
                fCursor = junctFeatures.Search(qFilter, False)
                jJunct = fCursor.NextFeature
                jNodePoint = jJunct.Shape

                pPointCollection = New Polyline
                pPointCollection.AddPoint(iNodePoint)
                pPointCollection.AddPoint(jNodePoint)
                Key.INode = pair.Key
                Key.JNode = pair.Value
                myDict.Add(Key, pPointCollection)









            Next pair
       
            Return myDict
        Catch ex As Exception
            MessageBox.Show(ex.ToString)

        End Try
    
    End Function
    Public Sub AddWeaveEdges(ByVal featureClass As IFeatureClass, ByVal dctWeavePolyLine As Dictionary(Of TwoValue, ESRI.ArcGIS.Geometry.Polyline), ByVal dctTODWeave As Dictionary(Of Long, Long), ByVal timeOfDay As Integer)
        'Adds the weave edges to the exported shapefiles

        ' Create the feature buffer.
        Dim featureBuffer As IFeatureBuffer = featureClass.CreateFeatureBuffer()

        ' Create insert feature cursor using buffering.
        Dim featureCursor As IFeatureCursor = featureClass.Insert(True)

        ' All of the features to be created were installed by "B Pierce."
        'Dim JunctIDFieldIndex As Integer = featureClass.FindField("PSRCjunctI")
        Dim edgeIDFieldIndex As Integer = featureClass.FindField("PSRCEdgeID")
        Dim facilityTypeIndex As Integer = featureClass.FindField("FacilityType")
        Dim directionIndex As Integer = featureClass.FindField("Direction")

        Dim newINodeFieldIndex As Integer = featureClass.FindField("NewINode")
        Dim newJNodeFieldIndex As Integer = featureClass.FindField("NewJNode")

        'featureBuffer.Value(JunctIDFieldIndex) = 99
        Dim item As IPolyline
        Dim pGeometry As IGeometry
        Dim pair As KeyValuePair(Of TwoValue, ESRI.ArcGIS.Geometry.Polyline)
        Dim Key2 As TwoValue
        For Each pair In dctWeavePolyLine
            Key2 = pair.Key
            If dctTODWeave.ContainsKey(Key2.INode) Then


                pGeometry = pair.Value
                featureBuffer.Shape = pGeometry
                'featureBuffer.Value(edgeIDFieldIndex) = pair.Key
                'Flag that this is an HOV
                featureBuffer.Value(edgeIDFieldIndex) = 98
                featureBuffer.Value(newINodeFieldIndex) = Key2.INode
                featureBuffer.Value(newJNodeFieldIndex) = Key2.JNode
                'direction is always 2 ways for weaves
                featureBuffer.Value(directionIndex) = 0


                featureCursor.InsertFeature(featureBuffer)
            End If

        Next pair


        Try
            featureCursor.Flush()
        Catch comExc As COMException

            ' Attempt to flush the buffer.
            ' Handle the error in a way appropriate to your application.
        Finally
            ' Release the cursor as it's no longer needed.
            Marshal.ReleaseComObject(featureCursor)
        End Try

    End Sub
    Structure TwoValue
        'a struct to hold I & J node attributes for weave links
        Public INode As Long
        Public JNode As Long
    End Structure
    Structure edgeHOVAtts
        'struct to hold some edge attributes. 
        Public psrcEdgeID As Long

        Public newInode As Long
        Public newJNode As Long

        'an array to hold time of day direction (mostly for HOV on the reversibles)
        Public todDirection() As Short
        
    End Structure
   
    
    Public Function todHOVatts(ByVal edgesFC As IFeatureClass, ByVal mAtts As Dictionary(Of Long, Row)) As Dictionary(Of Long, edgeHOVAtts)
        'Returns a dictionary:
        'key= EdgeID, value = edgeHOVatts




        Dim mymodeAttsClass As clsModeAttributes

        Dim boolIJ As Boolean = False
        Dim boolJI As Boolean = False
        Dim dctHOVAtts As New Dictionary(Of Long, edgeHOVAtts)
        Dim pair As KeyValuePair(Of Long, Row)
        Dim x As Integer

        




        For Each pair In mAtts

            Dim myEdgeHOVAtts As edgeHOVAtts
            'need to redim to specify the size of the todDirection Array (which is an attribute of the edgeHOVAtts Struct)
            ReDim myEdgeHOVAtts.todDirection(4)

            mymodeAttsClass = New clsModeAttributes(pair.Value)
            myEdgeHOVAtts.psrcEdgeID = pair.Key

            For x = 0 To 4


                If mymodeAttsClass.IJLANESHOVAM > 0 Then
                    boolIJ = True
                End If
                If mymodeAttsClass.JILANESHOVAM > 0 Then
                    boolJI = True
                End If
                If (boolIJ = False And boolJI = False) Then
                    myEdgeHOVAtts.todDirection(x) = -2
                ElseIf boolIJ = True And boolJI = True Then
                    myEdgeHOVAtts.todDirection(x) = 0
                ElseIf boolIJ = True And boolJI = False Then

                    myEdgeHOVAtts.todDirection(x) = 1

                Else
                    myEdgeHOVAtts.todDirection(x) = -1
                End If
                'myEdgeHOVAtts.psrcEdgeID = pair.Key
                'dctHOVAtts.Add(pair.Key, myEdgeHOVAtts)

                boolIJ = False
                boolJI = False

            Next x

            'add tod attributes for this edge to dict
            myEdgeHOVAtts.psrcEdgeID = pair.Key
            dctHOVAtts.Add(pair.Key, myEdgeHOVAtts)

        Next

            Return dctHOVAtts


    End Function
    Public Sub CalcGP_IJNOdes(ByVal edges As IFeatureClass, ByVal dctReservedNodes As Dictionary(Of Long, Long), ByVal TOD As Integer, ByVal nodeOffset As Long)
        'Calc the New GP I & J Nodes 
        Dim pFeature As IFeature

        Dim pFCursor As IFeatureCursor
        Dim indexEdgeID As Long
        Dim indexINodeField As Long
        Dim indexJNodeField As Long
        Dim indexDirectionField As Long
        Dim indexNewINodeField As Long
        Dim indexNewJNodeField As Long
        Dim indexOnewayField As Long
        Dim indexModeField As Long

        indexEdgeID = edges.FindField("PSRCEdgeId")
        indexINodeField = edges.FindField("INode")
        indexJNodeField = edges.FindField("JNode")
        indexNewINodeField = edges.FindField("NewINode")
        indexNewJNodeField = edges.FindField("NewJNode")
        indexOnewayField = edges.FindField("Oneway")
        indexDirectionField = edges.FindField("Direction")
        indexModeField = edges.FindField("Modes")
        Dim INode As Long
        Dim JNOde As Long
        Dim oneWay As Long
        Dim modes As String


        pFCursor = edges.Update(Nothing, False)
        pFeature = pFCursor.NextFeature()
        Do Until pFeature Is Nothing
            INode = pFeature.Value(indexINodeField)
            JNOde = pFeature.Value(indexJNodeField)
            modes = pFeature.Value(indexModeField)
            If dctReservedNodes.ContainsKey(INode) Then
                pFeature.Value(indexNewINodeField) = dctReservedNodes.Item(INode)
            Else
                pFeature.Value(indexNewINodeField) = INode + nodeOffset
            End If

            If dctReservedNodes.ContainsKey(JNOde) Then
                pFeature.Value(indexNewJNodeField) = dctReservedNodes.Item(JNOde)
            Else
                pFeature.Value(indexNewJNodeField) = JNOde + nodeOffset
            End If

            'now for direction:
            oneWay = pFeature.Value(indexOnewayField)
            Select Case oneWay
                Case 0
                    'If IJ allows walk, need a JI link for walk mode
                    If modes.Contains("w") Then
                        pFeature.Value(indexDirectionField) = 0
                    Else
                        pFeature.Value(indexDirectionField) = 1
                    End If

                Case 1
                    'If JI allows walk, need a IJ link for walk mode
                    If modes.Contains("w") Then
                        pFeature.Value(indexDirectionField) = 0
                    Else
                        pFeature.Value(indexDirectionField) = -1
                    End If

                Case 2
                    pFeature.Value(indexDirectionField) = 0

                Case Else
                    'reversibles

                    If TOD = 0 Then
                        pFeature.Value(indexDirectionField) = 1
                    ElseIf TOD = 1 Then
                        pFeature.Value(indexDirectionField) = 0
                    Else
                        pFeature.Value(indexDirectionField) = -1
                    End If

            End Select



            pFCursor.UpdateFeature(pFeature)
            pFeature = pFCursor.NextFeature()

        Loop
    End Sub
    Public Sub calcIsZone(ByVal junctions As IFeatureClass, ByVal nodeOffSet As Long)
        'adds a 1 to the IsZone field in the Junctions shapefile if it is a Centroid. 
        Dim pFeature As IFeature
        Dim pFCursor As IFeatureCursor
        Dim indexIsZoneField As Long
        Dim indexScenNode As Long
        indexIsZoneField = junctions.FindField("IsZone")
        indexScenNode = junctions.FindField("Scen_Node")






        pFCursor = junctions.Update(Nothing, False)
        pFeature = pFCursor.NextFeature()
        Do Until pFeature Is Nothing
            If pFeature.Value(indexScenNode) <= nodeOffSet Then
                pFeature.Value(indexIsZoneField) = 1
            Else
                pFeature.Value(indexIsZoneField) = 0
            End If


            pFCursor.UpdateFeature(pFeature)
            pFeature = pFCursor.NextFeature()

        Loop



    End Sub

    Public Sub getEdgeAtts(ByVal pFClsEdge As IFeatureClass, ByVal pMaTbl As ITable, ByVal pProjTable As ITable, ByVal EdgeIDFld As String, ByVal dctEdgeAtts As Dictionary(Of Long, Row), Optional ByVal whereClause As String = "")

        'For each edge that has at least one TOD HOV Lane, populates dctHOV_MAtt with rows from modeAttributes or tblScenarioProjects (if the edge is updated by a project)

        Dim dctEdge2 As New Dictionary(Of Long, Boolean)
        'dctEdge2 = New Dictionary

        If dctEdgeAtts Is Nothing Then dctEdgeAtts = New Dictionary(Of Long, Row)

        'struct that tells where the HOV attribute should come from, modeAttributes or tblScenarioProjects (if it's put in by a project)
        'Dim myHOVModeTable As HOVModeTable
        Dim boolUseProjectTable As Boolean
        Dim pDS As IDataset
        pDS = pMaTbl
        Debug.Print("getMAtt: " & pDS.Name & " start at " & Now())

        Dim pCs As ICursor, pRow As IRow
        Dim pFCS As IFeatureCursor, pFt As IFeature
        Dim sEdgeID As String
        Dim fld As Long
        Dim indexUpdated1 As Long
        Dim pQueryFilter As IQueryFilter = New QueryFilter
        pQueryFilter.WhereClause = whereClause


        If whereClause = "" Then
            fld = pFClsEdge.FindField(EdgeIDFld)
            pFCS = pFClsEdge.Search(Nothing, False)
        Else
            pFCS = pFClsEdge.Search(pQueryFilter, False)
        End If
        fld = pFClsEdge.FindField(EdgeIDFld)
        indexUpdated1 = pFClsEdge.FindField("Updated1")
        'pFCS = pFClsEdge.Search(Nothing, False)

        'start looping through edges
        pFt = pFCS.NextFeature
        Do Until pFt Is Nothing
            boolUseProjectTable = False
            sEdgeID = CStr(pFt.Value(fld))
            'This field is null if project is not being updated by a project 
            If Not IsDBNull(pFt.Value(indexUpdated1)) Then
                boolUseProjectTable = True
            End If

            If Not dctEdge2.ContainsKey(sEdgeID) Then dctEdge2.Add(sEdgeID, boolUseProjectTable)
            '[050407] hyu
            'If Not dctEdge.Exists(sEdgeOID) Then dctEdge.Add sEdgeOID, sEdgeOID

            pFt = pFCS.NextFeature
        Loop

        Debug.Print("finish get Edges: " & Now())

        fld = pMaTbl.FindField(EdgeIDFld)
        pCs = pMaTbl.Search(Nothing, False)
        pRow = pCs.NextRow
        'loop through all records in pTbl. If edgeID in dctEdge2, then add row to dctMAtt

        'Mode Attributes
        Do Until pRow Is Nothing

            If Not IsDBNull(pRow.Value(fld)) Then
                sEdgeID = CStr(pRow.Value(fld))
                If dctEdge2.ContainsKey(sEdgeID) Then
                    'only use mode attributes (boolUseProjectTable) is false
                    If dctEdge2.Item(sEdgeID) = False Then
                        If Not dctEdgeAtts.ContainsKey(sEdgeID) Then dctEdgeAtts.Add(sEdgeID, pRow)
                    End If

                End If

            End If
            pRow = pCs.NextRow
        Loop

        'Project Attributes
        pDS = pProjTable
        fld = pProjTable.FindField(EdgeIDFld)
        pCs = pProjTable.Search(Nothing, False)
        pRow = pCs.NextRow
        Do Until pRow Is Nothing

            If Not IsDBNull(pRow.Value(fld)) Then
                sEdgeID = CStr(pRow.Value(fld))
                If dctEdge2.ContainsKey(sEdgeID) Then
                    'only use mode attributes (boolUseProjectTable) is false
                    If dctEdge2.Item(sEdgeID) = True Then
                        If Not dctEdgeAtts.ContainsKey(sEdgeID) Then dctEdgeAtts.Add(sEdgeID, pRow)
                    End If

                End If

            End If
            pRow = pCs.NextRow
        Loop

        Debug.Print("getMAtt: " & pDS.Name & " end at " & Now())

        dctEdge2.Clear()
        dctEdge2 = Nothing
        pCs = Nothing
        '     pEnumV = Nothing
        '     pDataStat = Nothing
        pDS = Nothing

    End Sub

    Private Function getBATLaneSpeed(ByVal pRow As IRow, ByVal pMRow As IRow, ByVal pFeature As IFeature) As String
        Dim Tindex As Long
        Dim strSpeedLimit As String
        Dim strName As String
        Dim strSounderName As String
        Dim dblLength As Double
        Dim dblSpeed As Double
        Dim dblSpeed2 As Double
        Dim lngSpeed As Long
        dblLength = pFeature.Value(pFeature.Fields.FindField("Shape.len"))

        dblSpeed = 0.00378787878787879 * dblLength
        dblSpeed = Round(dblSpeed, 2)
        If dblSpeed = 0 Then
            strSpeedLimit = "0.01"
        Else
            strSpeedLimit = CStr(dblSpeed)
        End If
        getBATLaneSpeed = strSpeedLimit

    End Function



End Module
