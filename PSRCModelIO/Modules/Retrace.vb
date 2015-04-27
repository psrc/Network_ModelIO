Imports Scripting
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry

Module Retrace
    Public Sub RetraceTransitSegments(ByVal app As IApplication, ByVal ModelYear As Long)
        Debug.Print(Now() & "start: ")
        WriteLogLine(Now() & " start retracing transit segments")

        'Dim pStep As IStatusBar
        Dim lTransitFID As Object
        'lTransitFID = Array(857, 1421, 1420, 1870, 1869, 2637, 2638, 2904, 2905, 3195, 3196, 856)

        'pStep = app.StatusBar
        'pStep.ProgressBar.Position = 0
        'pStep.ShowProgressBar "build topology", 0, 1, 1, True
        '    init Application

        Dim pDoc As IMxDocument
        Dim pMap As IMap
        pDoc = app.Document
        pMap = pDoc.FocusMap

        Dim pTransitLinesFL As IFeatureLayer
        'Dim pTransitLinesFC As IFeatureClass

        Dim pFl As IFeatureLayer
        Dim pFCls As IFeatureClass, pTrFCls As IFeatureClass
        Dim pSegTbl As ITable, pSegCs As ICursor
        Dim pFClsE As IFeatureClass, pFClsJ As IFeatureClass, pFClsTP As IFeatureClass
        Dim pGeo As IGeometry, pFt As IFeature, pFt2 As IFeature, pFt3 As IFeature, pFt3_2 As IFeature
        Dim pEnv As IEnvelope2

        Dim lFidN As Long 'actually, it's the psrc junction id
        Dim pPline As IPolyline4, pPline2 As IPolyline4
        Dim l, K, m, n As Long, lSeq As Long
        Dim pPtCol As IPointCollection
        Dim pPt As IPoint, pPt2 As IPoint
        Dim dctSeq As Dictionary(Of Object, Object)
        Dim dctDist As Dictionary(Of Object, Object)
        Dim dDistAlg As Double, dDistFrm As Double, bright As Boolean
        Dim sFidE As String, sFidE2 As String
        Dim sFidsE() As String, sFidsE2() As String
        Dim bFindNode As Boolean, bFindEdge As Boolean
        Dim lNodeID As Long, lPrevNodeID As Long
        Dim pHT As IHitTest
        Dim dDist As Double, lPart As Long, lSeg As Long
        Dim attCur As jctAtt, attNext As jctAtt

        g_FWS = getPGDws(pMap)
        'Set pTransitLinesFL = getfeaturelayer("TransitLines") 'm_layers(13)
        pTransitLinesFL = getfeaturelayer(m_layers(13), app)
        '    Set pTrFCls = pFl.FeatureClass
        'Set pTrFCls = getInterimTransitLines
        pTrFCls = pTransitLinesFL.FeatureClass
        'Set pFL = getfeaturelayer("TransitPoints") 'm_layers(23)
        pFl = getfeaturelayer(m_layers(23), app)
        pFClsTP = pFl.FeatureClass

        '    Set pSegTbl = get_TableClass("tblTransitSegments")
        pSegTbl = get_TableClass(m_layers(14))
        '    Set pSegCs = pSegTbl.Insert(True)

        '    Set pFL = getfeaturelayer("TransRefEdges") 'm_layers(0)
        pFl = getfeaturelayer(m_layers(22), app) 'ScenarioEdge
        pFClsE = pFl.FeatureClass

        '    Set pFL = getfeaturelayer("TransRefJunctions") 'm_layers(1)
        pFl = getfeaturelayer(m_layers(21), app) 'ScenarioJunct
        pFClsJ = pFl.FeatureClass


        Dim sPath As String
        sPath = app.Templates.Item(1)
        sPath = Left(sPath, InStrRev(sPath, "\"))

        'Open "D:\projects\7494 PSRC Model App\phase2_7743\modelApp\t.txt" For Output As 1

        FileOpen(1, sPath & "timeline.txt", OpenMode.Output)
        PrintLine(1, Now() & " start")
        Dim pRelOp As IRelationalOperator
        Dim i As Integer
        Dim pSort As ITableSort, pFCS As IFeatureCursor
        Dim sTrLineNo As String, fldTrLineNo As Long
        Dim dctLine As Dictionary(Of Object, Object)
        Dim sLineID As String
        Dim pWSedit As IWorkspaceEdit
        pWSedit = g_FWS

        dctSeq = New Dictionary(Of Object, Object)

        Dim pFilt As IQueryFilter2
        Dim lModelYear As Long
        lModelYear = ModelYear
        If lModelYear = 0 Then lModelYear = 2000
        pFilt = New QueryFilter
        '090808 sec: pTrFCls is a selection copy of TransitRoutes based on in and out service date. The fields get truncated and
        'have re-written the following query to reflact that.

        pFilt.WhereClause = "INSERVICEDATE=" & CStr(lModelYear) & " and OUTSERVICEDATE>" & CStr(lModelYear)
        'pFilt.WhereClause = "INSERVICED=" & CStr(lModelYear) & " and OUTSERVICE>" & CStr(lModelYear)

        '    pStep.ShowProgressBar "processing...", 0, pTrFCls.featurecount(pFilt), 1, True
        'pStep.ShowProgressBar("processing...", 0, pTrFCls.FeatureCount(Nothing), 1, True)
        'MsgBox pTrFCls.featurecount(pFilt)
        '    Set pFCS = pTrFCls.Search(pFilt, False)
        pFCS = pTrFCls.Search(pFilt, False)
        '    fldTrLineNo = pTrFCls.FindField(g_TrlineID)

        pFt = pFCS.NextFeature
        PrintLine(1, Now() & " delete all rows in the transit segment table.")
        deleteAllRows(pWSedit, pSegTbl)
        PrintLine(1, Now() & " finish deleting.")

        Dim lct As Long
        '    Dim pMapCache As IMapCache

        pWSedit.StartEditing(True)
        pWSedit.StartEditOperation()
        '    Set pMapCache = pMap
        lct = 0
        Dim sSEq As String
        Dim x As Integer
        x = 0
        Do Until pFt Is Nothing
            lct = lct + 1
            sTrLineNo = pFt.Value(fldTrLineNo)
            'pStep.StepProgressBar()
            SortSequence(pFClsE, pFClsJ, pFClsTP, pFt, pSegTbl, dctSeq, lModelYear)
            PrintLine(1, Now() & " finish sorting out nodes sequence " & CStr(lct))

            'Print #1, vbNewLine & "transit Line FID = " & lTransitFID(i)
            PrintLine(1, dctSeq.Count)
            K = 0
            sSEq = ""
            For i = 0 To dctSeq.Count - 1
                K = K + 1
                sSEq = sSEq & "," & dctSeq.Item(i)
                If K = 20 Then
                    PrintLine(1, sSEq)
                    sSEq = ""
                    K = 0
                End If
            Next i
            If sSEq <> "" Then PrintLine(1, sSEq)
            dctSeq.Clear()

            PrintLine(1, vbNewLine)
            pFt = pFCS.NextFeature
            x = x + 1
        Loop
        'MsgBox x

        'pStep.HideProgressBar()

        'pStep = Nothing
        pWSedit.StopEditOperation()
        pWSedit.StopEditing(True)
        WriteLogLine(Now() & " finish retracing transit segments")

        FileClose(1)
        '    MsgBox "OK"
        System.Runtime.InteropServices.Marshal.ReleaseComObject(pFCS)
        pWSedit = Nothing
        dctSeq = Nothing
        dctLine = Nothing
        'pTG = Nothing
        pDoc = Nothing
        'pMapCache = Nothing
        pMap = Nothing
        pFl = Nothing
        pFCls = Nothing
        'pRel = Nothing
        pTrFCls = Nothing
        Debug.Print(Now() & "Stop: ")
    End Sub


    Private Function getTopologyLayer(ByVal app As IApplication) As ITopologyLayer

        Dim pDoc As IMxDocument
        Dim pMap As IMap
        Dim pEnumLy As IEnumLayer
        pDoc = app.Document
        pMap = pDoc.FocusMap
        pEnumLy = pMap.Layers

        Dim pLy As ILayer, pFLy As IFeatureLayer
        pLy = pEnumLy.Next
        Do Until pLy Is Nothing

            If pLy.Valid Then
                If TypeOf pLy Is ITopologyLayer Then
                    getTopologyLayer = pLy
                End If
            End If
            pLy = pEnumLy.Next
        Loop

ReleaseObjs:
        pDoc = Nothing
        pMap = Nothing
        pEnumLy = Nothing
        pLy = Nothing
    End Function


    Private Function getfeaturelayer(ByVal sLayerName As String, ByVal app As IApplication, Optional ByVal bLayerName As Boolean = False) As IFeatureLayer2

        Dim pDoc As IMxDocument
        Dim pMap As IMap
        Dim pEnumLy As IEnumLayer
        pDoc = app.Document
        pMap = pDoc.FocusMap
        pEnumLy = pMap.Layers

        Dim pLy As ILayer, pFLy As IFeatureLayer
        pLy = pEnumLy.Next
        Do Until pLy Is Nothing

            If pLy.Valid Then
                If TypeOf pLy Is IFeatureLayer Then
                    pFLy = pLy
                    If bLayerName Then
                        If UCase(pLy.Name) = UCase(sLayerName) Then
                            getfeaturelayer = pFLy
                            GoTo ReleaseObjs
                        End If
                    Else
                        If Not TypeOf pFLy.FeatureClass Is IRelQueryTable Then
                            If UCase(pFLy.FeatureClass.AliasName) = UCase(sLayerName) Then
                                getfeaturelayer = pFLy
                                GoTo ReleaseObjs
                            End If
                        End If
                    End If
                End If
            End If
            pLy = pEnumLy.Next
        Loop

ReleaseObjs:
        pDoc = Nothing
        pMap = Nothing
        pEnumLy = Nothing
        pLy = Nothing
        'Set pFLy = Nothing
    End Function


    Private Function getPointAtt(ByVal pFt As IFeature) As jctAtt
        Dim att As jctAtt
        att = New jctAtt
        With pFt
            For i = 0 To .Fields.FieldCount - 1

                Select Case UCase(.Fields.Field(i).Name)
                    Case "LINEID"
                        att.LineID = .Value(i)
                    Case "DWTSTOP"
                        att.DwtStop = .Value(i)
                    Case "ISTIMEPOINT"
                        att.isTimePoint = .Value(i)
                    Case "LAYOVER"
                        att.LayOver = .Value(i)
                    Case "PATH"
                        att.Path = .Value(i)
                    Case "POINTORDER"
                        att.PointOrder = .Value(i)
                    Case "TIMEFUNCID"
                        att.TimeFuncID = .Value(i)
                    Case "USEGPONLY"
                        att.UseGPOnly = .Value(i)
                    Case "USER1"
                        att.USER1 = .Value(i)
                    Case "USER2"
                        att.USER2 = .Value(i)
                    Case "USER3"
                        att.USER3 = .Value(i)
                    Case "txtDWT"
                        att.txtDWT = .Value(i)
                End Select
            Next i
        End With

        getPointAtt = att
    End Function


    Private Function writeSegment(ByVal pTbl As ITable, ByVal pCs As ICursor, ByVal fromNode As Long, ByVal toNode As Long, ByVal att As jctAtt, ByVal SegNo As Long, ByVal sLineID As String, ByVal isTransitPoint As Boolean, ByVal SegLength As Long)    'pPointFeature As IFeature, SegNo As Long,  SegLength as long)
        Dim pRBuf As IRowBuffer
        pRBuf = pTbl.CreateRowBuffer

        If att Is Nothing Then att = New jctAtt
        With pRBuf
            For i = 0 To .Fields.FieldCount - 1
                If .Fields.Field(i).Editable Then
                    Select Case UCase(.Fields.Field(i).Name)
                        Case "SEGORDER"
                            .Value(i) = SegNo
                        Case "INODE"
                            .Value(i) = fromNode
                        Case "JNODE"
                            .Value(i) = toNode
                        Case "LINEID"
                            .Value(i) = sLineID ' att.LineID
                        Case "DWTStop"
                            If isTransitPoint = True Or SegNo = 1 Then
                                .Value(i) = att.DwtStop
                            Else
                                .Value(i) = 0
                            End If
                        Case "ISTIMEPOINT"
                            .Value(i) = att.isTimePoint
                        Case "LAYOVER"
                            .Value(i) = att.LayOver
                        Case "PATH"
                            .Value(i) = att.Path
                        Case "TIMEFUNCID"
                            .Value(i) = att.TimeFuncID
                        Case "USEGPONLY"
                            .Value(i) = att.UseGPOnly
                        Case "USER1"
                            .Value(i) = att.USER1
                        Case "USER2"
                            .Value(i) = att.USER2
                        Case "USER3"
                            .Value(i) = att.USER3
                        Case "CHANGE"
                            .Value(i) = SegLength
                    End Select
                    'pRBuf.value (pRBuf.Fields.FindField("cahnge"))

                    '                If UCase(pRBuf.Fields.field(i).name) = "SEGORDER" Then
                    '                    pRBuf.value(i) = SegNo
                    '                ElseIf UCase(pRBuf.Fields.field(i).name) = "INODE" Then
                    '                    pRBuf.value(i) = FromNode
                    '                ElseIf UCase(pRBuf.Fields.field(i).name) = "JNODE" Then
                    '                    pRBuf.value(i) = ToNode
                    '                Else
                    '                    If pPointFeature.Fields.FindField(pRBuf.Fields.field(i).name) > -1 Then
                    '                        pRBuf.value(i) = pPointFeature.value(pPointFeature.Fields.FindField(pRBuf.Fields.field(i).name))
                    '                    End If
                    '                End If
                End If
            Next i
        End With

        pCs.InsertRow(pRBuf)

        pRBuf = Nothing
    End Function


    Private Sub deleteAllRows(ByVal pWSE As IWorkspaceEdit, ByVal pTbl As ITable)
        WriteLogLine("start deleting all rows: " & Now())
        If pTbl.RowCount(Nothing) = 0 Then Exit Sub
        Dim pREdit As IRowEdit
        Dim pRSet As ESRI.ArcGIS.esriSystem.SetClass
        pRSet = New ESRI.ArcGIS.esriSystem.SetClass

        Dim pCs As ICursor, pRow As IRow
        Dim i As Integer
        pCs = pTbl.Search(Nothing, False)

        Dim pDaStats As IDataStatistics
        Dim pStats As IStatisticsResults
        Dim lMax As Long, lMin As Long, l As Long
        Dim pFilt As IQueryFilter2

        pDaStats = New DataStatistics
        pDaStats.Field = pTbl.OIDFieldName
        pDaStats.Cursor = pCs
        pDaStats.SimpleStats = True
        pStats = pDaStats.Statistics
        lMin = pStats.Minimum
        lMax = pStats.Maximum
        l = lMax
        pFilt = New QueryFilter

        'deleting rows in desceding order of OID
        Do Until lMax <= lMin
            If lMax Mod 10000 > 0 Then
                If lMax > 10000 Then
                    'the first round, usually the largest OID is not a multiplication of 10000
                    l = Val(Left(CStr(lMax), Len(CStr(lMax)) - 4) & "0000")
                Else
                    l = 0
                End If
            Else
                l = lMax - 10000
            End If

            pFilt.WhereClause = pTbl.OIDFieldName & ">=" & CStr(l)

            If pWSE.IsBeingEdited = False Then pWSE.StartEditing(True)
            pWSE.StartEditOperation()
            pCs = pTbl.Search(pFilt, False)

            pRow = pCs.NextRow
            Do Until pRow Is Nothing
                pREdit = pRow
                pRSet.Add(pREdit)
                i = i + 1
                If i = 100 Then
                    pREdit.DeleteSet(pRSet)
                    pRSet = New ESRI.ArcGIS.esriSystem.SetClass
                    i = 0
                End If
                pRow = pCs.NextRow
            Loop
            If pRSet.count > 0 Then pREdit.DeleteSet(pRSet)
            lMax = l

            pWSE.StopEditOperation()
            pWSE.StopEditing(True)
        Loop
        pStats = Nothing
        pDaStats = Nothing

        '    Set pRow = pCs.NextRow
        '    If pWSE.IsBeingEdited = False Then pWSE.StartEditing True
        '    pWSE.StartEditOperation
        '    i = 0
        '    If Not pRow Is Nothing Then
        '
        '    Do Until pRow Is Nothing
        '        Set pREdit = pRow
        '        pRset.Add pREdit
        '        i = i + 1
        '        If i = 100 Then
        '            pREdit.DeleteSet pRset
        '            Set pRset = New esriSystem.Set
        '            i = 0
        '        End If
        '        Set pRow = pCs.NextRow
        '    Loop
        '    If pRset.count > 0 Then pREdit.DeleteSet pRset
        '    End If
        '    pWSE.StopEditOperation
        '    pWSE.StopEditing True

        WriteLogLine("finish deleting all rows: " & Now())
        pRSet = Nothing
        pREdit = Nothing
        pCs = Nothing

    End Sub


    Private Function getJunctions(ByVal pFCls As IFeatureClass, ByVal pRt As IPolyline4, ByVal IDFld As String, ByVal lModelYear As Long, ByVal strTransitNumber As String) As IPointCollection4
        Dim pSFilt As ISpatialFilter
        pSFilt = New SpatialFilter
        On Error GoTo eh
        With pSFilt
            .Geometry = getBuffer(pRt)
            '        If .Geometry Is Nothing Then Exit Function
            .GeometryField = pFCls.ShapeFieldName
            .SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects
            .WhereClause = "(INSERVICEDATE<=" & lModelYear & " AND OUTSERVICEDATE>" & lModelYear _
                & ") OR (INSERVICEDATE IS NULL AND OUTSERVICEDATE IS NULL) " _
                & " OR (INSERVICEDATE=0 AND OUTSERVICEDATE=0)"
            'NULL VALUE in the inservicedate and outservicedate indicates it's a new junction created during creating weavelink to split an arterial for serving as weave link

        End With

        Dim pFCS As IFeatureCursor
        Dim pFt As IFeature
        Dim pPtCol As IPointCollection4
        Dim pPt As IPoint
        Dim fldID As Long
        pFCS = pFCls.Search(pSFilt, False)
        pPtCol = New Multipoint
        fldID = pFCls.FindField(IDFld)

        pFt = pFCS.NextFeature
        Do Until pFt Is Nothing
            pPt = pFt.ShapeCopy
            If Not pPt.IsEmpty Then
                pPt.ID = pFt.Value(fldID)
                pPtCol.AddPoint(pPt)
            End If
            pFt = pFCS.NextFeature
        Loop

        getJunctions = pPtCol
        pPt = Nothing
        pFCS = Nothing
        pSFilt = Nothing
        Exit Function


eh:
        WriteLogLine("Error in getJunctions for TransitLine Number " & strTransitNumber)


    End Function

    Private Function getEdges(ByVal pFCls As IFeatureClass, ByVal pRt As IPolyline4, ByVal fromNode As String, ByVal toNode As String, ByVal modelYEAR As Long, Optional ByVal oneWay As String = "") As Dictionary(Of Object, Object)
        Dim pSFilt As ISpatialFilter
        pSFilt = New SpatialFilter
        With pSFilt
            .Geometry = getBuffer(pRt)
            .GeometryField = pFCls.ShapeFieldName
            .SpatialRel = esriSpatialRelEnum.esriSpatialRelContains ' esriSpatialRelintersects
            '        .SubFields = fromNode & "," & toNode
            .WhereClause = "INSERVICEDATE<=" & modelYEAR & " AND OUTSERVICEDATE>" & modelYEAR
            .SearchOrder = esriSearchOrder.esriSearchOrderSpatial
        End With

        Dim pFCS As IFeatureCursor
        Dim pFt As IFeature
        Dim dct As Dictionary(Of Object, Object)
        Dim fldFromNd As Long, fldToNd As Long, fldOneWay As Long
        Dim sKey As String
        Dim pRelOp As IRelationalOperator
        Dim pfldLength As Integer
        Dim pfeatLength As Long

        dct = New Dictionary(Of Object, Object)
        fldFromNd = pFCls.FindField(fromNode)
        fldToNd = pFCls.FindField(toNode)
        pfldLength = pFCls.FindField("Shape.len")
        '    fldOneWay = pFCls.FindField(oneWay)

        pRelOp = pRt
        pFCS = pFCls.Search(pSFilt, True)
        PrintLine(1, Now() & " finish querying edges")
        pFt = pFCS.NextFeature
        Do Until pFt Is Nothing

            '        If pRelOp.Contains(pFt.Shape) Then
            sKey = pFt.Value(fldFromNd) & "-" & pFt.Value(fldToNd)
            pfeatLength = pFt.Value(pfldLength)
            '        If oneWay <> "" Then
            '            dct.Add sKey, pFt.value(fldOneWay)
            '        End If
            'If Not dct.ContainsKey(sKey) Then dct.Add sKey, sKey
            If Not dct.ContainsKey(sKey) Then dct.Add(sKey, pfeatLength)
            '        End If
            pFt = pFCS.NextFeature
        Loop

        getEdges = dct

        pFCS = Nothing
        pSFilt = Nothing
    End Function

    Private Function getAllEdges(ByVal pFCls As IFeatureClass, ByVal fromNode As String, ByVal toNode As String, ByVal modelYEAR As Long) As Dictionary(Of Object, Object)
        Dim pFilt As IQueryFilter2
        Dim pFCS As IFeatureCursor
        Dim pFt As IFeature
        Dim dct As Dictionary(Of Object, Object)
        Dim fldFromNd As Long, fldToNd As Long, fldOneWay As Long
        Dim sKey As String

        pFilt = New QueryFilter
        pFilt.WhereClause = "INSERVICEDATE<=" & modelYEAR & " AND OUTSERVICEDATE>" & modelYEAR
        pFCS = pFCls.Search(pFilt, False)

        dct = New Dictionary(Of Object, Object)
        fldFromNd = pFCls.FindField(fromNode)
        fldToNd = pFCls.FindField(toNode)
        fldOneWay = pFCls.FindField("OneWay")
        PrintLine(1, Now() & " finish querying edges")
        pFt = pFCS.NextFeature
        Do Until pFt Is Nothing

            sKey = pFt.Value(fldFromNd) & "-" & pFt.Value(fldToNd)
            '        If oneWay <> "" Then
            '            If pFt.value(fldOneWay) = 3 Then
            '                dct.Add sKey, 2
            '            Else
            '                dct.Add sKey, 1
            '            End If
            '        End If
            dct.Add(sKey, sKey)

            pFt = pFCS.NextFeature
        Loop

        getAllEdges = dct
        pFCS = Nothing

    End Function

    Private Function GetTransitPoints(ByVal pFCls As IFeatureClass, ByVal pRt As IPolyline4, ByVal lineIDField As String, ByVal lineIDValue As String, ByRef pPtCol As IPointCollection4, ByRef dct As Dictionary(Of Object, Object))
        Dim pSFilt As ISpatialFilter
        pSFilt = New SpatialFilter
        With pSFilt
            .Geometry = getBuffer(pRt)
            .GeometryField = pFCls.ShapeFieldName
            .SpatialRel = esriSpatialRelEnum.esriSpatialRelContains
            .WhereClause = lineIDField & "=" & lineIDValue '& "'"
        End With

        Dim pFCS As IFeatureCursor
        Dim pFt As IFeature
        Dim pPt As IPoint ', pPtCol As IPointCollection4
        'Dim dct As Dictionary
        Dim fldID As Long, fldToNd As Long, fldOneWay As Long
        Dim att As jctAtt
        Dim l As Long
        Dim pSort As ITableSort
        pSort = New TableSort
        With pSort
            .Table = pFCls
            .QueryFilter = pSFilt
            .Fields = "PointOrder"
            .Ascending("PointOrder") = False
            .Sort(Nothing)
        End With
        pFCS = pSort.Rows
        'Set pFCS = pFCls.Search(pSFilt, False)
        pPtCol = New Multipoint
        dct = New Dictionary(Of Object, Object)
        fldID = pFCls.FindField("Line")
        fldToNd = pFCls.FindField("toNode")
        fldOneWay = pFCls.FindField("oneWay")
        l = 0
        pFt = pFCS.NextFeature
        Do Until pFt Is Nothing
            l = l + 1
            pPt = pFt.ShapeCopy
            pPt.ID = l
            att = getPointAtt(pFt)

            pPtCol.AddPoint(pPt)
            dct.Add(CStr(l), att)

            pFt = pFCS.NextFeature
        Loop

        GetTransitPoints = dct

        pFCS = Nothing
        pSFilt = Nothing
    End Function

    Private Function SortSequence(ByVal pFClsE As IFeatureClass, ByVal pFClsJ As IFeatureClass, ByVal pFClsTP As IFeatureClass, ByVal plineFt As IFeature, _
        ByVal pSegTbl As ITable, ByVal dctSeq As Dictionary(Of Object, Object), ByVal modelYEAR As Long)
        Dim pSegCs As ICursor
        Dim pTrLine As IPolyline4
        Dim sLineID As String
        Dim sLineNo As String
        Dim pJctCol As IPointCollection4, pTPCol As IPointCollection4
        Dim dctE As Dictionary(Of Object, Object), dctTP As Dictionary(Of Object, Object), dctN As Dictionary(Of Object, Object)

        pTrLine = plineFt.Shape
        sLineID = plineFt.Value(plineFt.Fields.FindField(g_TrlineID))
        'sLineNo = plineFt.value(plineFt.Fields.FindField("TransLineN"))
        sLineNo = plineFt.Value(plineFt.Fields.FindField("TransLineNo"))
        PrintLine(1, Now() & " start sorting " & sLineID)

        'get the junctions under the line
        pJctCol = getJunctions(pFClsJ, pTrLine, g_PSRCJctID, modelYEAR, sLineNo)
        '081408 sec: i added this to deal with bad transit routes:
        If pJctCol Is Nothing Then Exit Function

        'Set pJctCol = getJunctions(pFClsJ, pTrLine, g_PSRCJctID, modelYEAR)
        PrintLine(1, Now() & " get junctions " & pJctCol.PointCount)

        'get the edges under the line
        dctE = getEdges(pFClsE, pTrLine, g_INode, g_JNode, modelYEAR)
        '    Set dctE = dctEdges
        PrintLine(1, Now() & " get edges " & dctE.Count)

        'get the TransitPoints for the line
        GetTransitPoints(pFClsTP, pTrLine, g_TrlineID, sLineID, pTPCol, dctTP)
        PrintLine(1, Now() & " get transit points " & dctTP.Count)

        Dim pHtJ As IHitTest, pHtTP As IHitTest
        Dim lPart As Long, lSeg As Long, pHtPt As IPoint, bright As Boolean, dDist As Double
        Dim lPart2 As Long, lSeg2 As Long, pHtPt2 As IPoint
        Dim pPtCol As IPointCollection4, pPt As IPoint
        Dim i As Long, lSeq As Long, lct As Long
        Dim PreJct As Long, CurJct As Long
        Dim curAtt As jctAtt, NextAtt As jctAtt
        Dim sKey As String
        Dim boolTransitPoint As Boolean
        Dim junctCount As Long
        Dim edgeCount
        Dim tpCounter As Long
        Dim boolSameTP As Boolean
        boolSameTP = False
        tpCounter = 0
        edgeCount = 0
        junctCount = 0
        Dim edgeLength As Long

        'transit line vertices
        pPtCol = pTrLine
        'hit test
        pHtJ = pJctCol
        pHtTP = pTPCol
        pPt = New Point
        pHtPt = New Point
        lSeq = 0
        lct = 0
        If dctSeq Is Nothing Then dctSeq = New Dictionary(Of Object, Object)
        dctSeq.Clear()

        dctN = New Dictionary(Of Object, Object)
        pSegCs = pSegTbl.Insert(True)
        'loop through all the vertices in the transit line
        Try


            For i = 0 To pPtCol.PointCount - 1
                pPt = pPtCol.Point(i)
                'does it hit a junction
                '1
                If pHtJ.HitTest(pPt, 0.5, esriGeometryHitPartType.esriGeometryPartVertex, pHtPt, dDist, lPart, lSeg, bright) Then
                    junctCount = junctCount + 1

                    'hit a junction

                    pJctCol.QueryPoint(lPart, pHtPt)
                    CurJct = pHtPt.ID

                    'recording the times the nodes has been traversed
                    '2
                    If dctN.ContainsKey(CType(CurJct, Integer)) Then
                        'junctCount = junctCount - 1
                        '***********Need to create a counter that keeps track of transit points bette
                        dctN.Item(CType(CurJct, Integer)) = dctN.Item(CType(CurJct, Integer) + 1)
                    Else
                        dctN.Add(CType(CurJct, Integer), 1)
                        '2
                    End If

                    '3
                    If PreJct > 0 Then
                        '4
                        If PreJct = CurJct Then
                            'hit the same transit point twice in a row so need to skip, but want it's atts for the next link
                            '5
                            'check to see if it's a transit point
                            If pHtTP.HitTest(pPt, 0.5, esriGeometryHitPartType.esriGeometryPartVertex, pHtPt2, dDist, lPart2, lSeg2, bright) Then
                                tpCounter = tpCounter + 1
                                NextAtt = dctTP.Item(CStr(pTPCol.Point(lSeg2).ID))
                                boolSameTP = True

                            End If


                            'hit a transit point
                            'some vertices may colustered within the searching tolerance, so the junction will be hit more than once.
                        Else
                            'is the current junction a transit point
                            'stay with current atts for now
                            NextAtt = curAtt
                            'is the current junction a transit point



                            If pHtTP.HitTest(pPt, 0.5, esriGeometryHitPartType.esriGeometryPartVertex, pHtPt2, dDist, lPart2, lSeg2, bright) Then
                                'hit a transit point
                                boolTransitPoint = True
                                tpCounter = tpCounter + 1
                                'want to use the TP's atts for the I node of the next link, not the J node of the current link!
                                NextAtt = dctTP.Item(CStr(pTPCol.Point(lSeg2).ID))
                                'If NextAtt.PointOrder = dctN.Item(CurJct) Then
                                'pTPCol.RemovePoints lSeg2, 1
                                'Else
                                'Set curAtt = NextAtt

                                '                        If NextAtt.PointOrder = dctN.Item(CurJct) Then
                                '                            Exit Do
                                '                        Else
                                '                            Set NextAtt = curAtt
                                'End If
                                '
                                '                    Loop
                            Else
                                boolTransitPoint = False
                            End If
                            '                        'dctTP.Remove CStr(lSeg2)
                            '                    Else
                            '                        Set NextAtt = curAtt
                            '                    End If

                            'check whether this junction and previous junction forms a valid edge
                            'be sure to check both directions
                            sKey = CStr(PreJct) & "-" & CStr(CurJct)
                            If dctE.ContainsKey(sKey) Then
                                edgeLength = dctE.Item(sKey)
                                edgeCount = edgeCount + 1
                                lSeq = lSeq + 1

                                If lSeq = 1 Then
                                    writeSegment(pSegTbl, pSegCs, PreJct, CurJct, NextAtt, lSeq, sLineID, boolTransitPoint, edgeLength)
                                Else
                                    writeSegment(pSegTbl, pSegCs, PreJct, CurJct, curAtt, lSeq, sLineID, boolTransitPoint, edgeLength)
                                End If

                                curAtt = NextAtt
                                NextAtt = Nothing
                                PreJct = CurJct
                                dctSeq.Add(dctSeq.Count, CType(CurJct, Integer))
                            Else
                                sKey = CStr(CurJct) & "-" & CStr(PreJct)
                                If dctE.ContainsKey(sKey) Then
                                    edgeLength = dctE.Item(sKey)
                                    edgeCount = edgeCount + 1
                                    lSeq = lSeq + 1
                                    If lSeq = 1 Then
                                        writeSegment(pSegTbl, pSegCs, PreJct, CurJct, NextAtt, lSeq, sLineID, boolTransitPoint, edgeLength)
                                    Else
                                        writeSegment(pSegTbl, pSegCs, PreJct, CurJct, curAtt, lSeq, sLineID, boolTransitPoint, edgeLength)
                                    End If

                                    curAtt = NextAtt
                                    NextAtt = Nothing
                                    PreJct = CurJct
                                    dctSeq.Add(dctSeq.Count, CurJct)
                                End If
                            End If
                        End If

                    Else
                        'the beginning of a line
                        PreJct = CurJct
                        'hittest the transit points
                        If pHtTP.HitTest(pPt, 0.01, esriGeometryHitPartType.esriGeometryPartVertex, pHtPt2, dDist, lPart2, lSeg2, bright) Then
                            'hit a transit point
                            boolTransitPoint = True
                            'keep track of which Transit Point we have
                            tpCounter = tpCounter + 1
                            'current Transit Point Attributes
                            curAtt = dctTP.Item(CStr(pTPCol.Point(lSeg2).ID))
                            'If curAtt.PointOrder = dctN.Item(CurJct) Then
                            'remove 1st junction
                            'pTPCol.RemovePoints lSeg2, 1
                            'Else
                            'Set curAtt = Nothing
                            'End If
                            '                    If curAtt.PointOrder = dctN.Item(CurJct) Then
                            '                        Exit Do
                            '                    Else
                            '                        Set curAtt = Nothing
                            '                    End If
                            '                Loop
                        Else
                            boolTransitPoint = False
                        End If



                        dctSeq.Add(dctSeq.Count, PreJct)
                    End If
                End If
            Next i
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        If edgeCount < (dctE.Count - 6) Then
            MsgBox(sLineID)
        End If



        pSegCs.Flush()
        pSegCs = Nothing
        pJctCol = Nothing
        pTPCol = Nothing
        dctE = Nothing
        dctN = Nothing
        dctTP = Nothing
        curAtt = Nothing
        NextAtt = Nothing
        pPtCol = Nothing

    End Function

    Private Function getBuffer(ByVal pPline As IPolyline4) As IPolygon4
        On Error GoTo eh
        Dim pTopOp As ITopologicalOperator2
        Dim pMpt As IMultipoint
        pTopOp = pPline

        'pTopOp.IsKnownSimple = False
        pTopOp.Simplify()
        getBuffer = pTopOp.Buffer(0.5)
        pTopOp = Nothing
        Exit Function
eh:
        'somehow the buffer method fails on some polyline with unknown reason
        Dim pSegCol As ISegmentCollection
        pSegCol = New Polygon
        pSegCol.SetRectangle(pPline.Envelope)
        getBuffer = pSegCol
    End Function
End Module
