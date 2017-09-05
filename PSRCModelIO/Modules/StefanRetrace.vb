

Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry
Module StefanRetrace



    Public Sub RetraceTransitSegments(ByVal app As IApplication, ByVal ModelYear As Long)
        Try

        
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
            Dim dctSeq As New Dictionary(Of Long, Long)
            Dim dctDist As New Dictionary(Of Long, Long)
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
            ' FileClose(1)

            'FileOpen(1, sPath & "timeline.txt", OpenMode.Input, OpenAccess.ReadWrite, OpenShare.Default)
            FileOpen(1, sPath & "timeline.txt", OpenMode.Output, OpenAccess.Write, OpenShare.Default)
            FileOpen(2, sPath & "transiterrors.txt", OpenMode.Output, OpenAccess.Write, OpenShare.Default)
            'FileOpen(17, TpathnameM, OpenMode.Input, OpenAccess.ReadWrite, OpenShare.Default)
            PrintLine(1, Now() & " start")
            Dim pRelOp As IRelationalOperator
            Dim i As Integer
            Dim pSort As ITableSort, pFCS As IFeatureCursor
            Dim sTrLineNo As String, fldTrLineNo As Long
            Dim dctLine As Dictionary(Of Long, Long)
            Dim sLineID As String
            Dim pWSedit As IWorkspaceEdit
            pWSedit = g_FWS

            'dctSeq = New Dictionary

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

                Dim pair As KeyValuePair(Of Long, Long)
                For Each pair In dctSeq



                    K = K + 1
                    sSEq = sSEq & "," & pair.Value

                    If K = 20 Then
                        PrintLine(1, sSEq)
                        sSEq = ""
                        K = 0
                    End If
                Next
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
            FileClose(2)
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
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
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
            If pRSet.Count > 0 Then pREdit.DeleteSet(pRSet)
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
            '.WhereClause = "(INSERVICEDATE<=" & lModelYear & " AND OUTSERVICEDATE>" & lModelYear _
            '& ") OR (INSERVICEDATE IS NULL AND OUTSERVICEDATE IS NULL) " _
            '& " OR (INSERVICEDATE=0 AND OUTSERVICEDATE=0)"
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

    Private Function getEdges(ByVal pFCls As IFeatureClass, ByVal pRt As IPolyline4, ByVal fromNode As String, ByVal toNode As String, ByVal modelYEAR As Long, Optional ByVal oneWay As String = "") As Dictionary(Of String, Long)
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
        Dim dct As New Dictionary(Of String, Long)
        Dim fldFromNd As Long, fldToNd As Long, fldOneWay As Long
        Dim sKey As String
        Dim pRelOp As IRelationalOperator
        Dim pfldLength As Integer
        Dim pfeatLength As Long


        fldFromNd = pFCls.FindField(fromNode)
        fldToNd = pFCls.FindField(toNode)


        If strLayerPrefix = "SDE" Then
            pfldLength = pFCls.Fields.FindField("Shape.len") 'SDE
            If pfldLength = -1 Then
                pfldLength = pFCls.Fields.FindField("Shape.STlength()")

            End If
        Else
            pfldLength = pFCls.Fields.FindField("Shape_Length") 'PGDB
        End If
        'pfldLength = pFCls.FindField("Shape.len")

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
            'If Not dct.Exists(sKey) Then dct.Add sKey, sKey
            If Not dct.ContainsKey(sKey) Then dct.Add(sKey, pfeatLength)
            '        End If
            pFt = pFCS.NextFeature
        Loop
        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(pFCS)

        getEdges = dct

        pFCS = Nothing
        pSFilt = Nothing
    End Function

    Private Function getAllEdges(ByVal pFCls As IFeatureClass, ByVal fromNode As String, ByVal toNode As String, ByVal modelYEAR As Long) As Dictionary(Of String, String)
        Dim pFilt As IQueryFilter2
        Dim pFCS As IFeatureCursor
        Dim pFt As IFeature
        Dim dct As New Dictionary(Of String, String)
        Dim fldFromNd As Long, fldToNd As Long, fldOneWay As Long
        Dim sKey As String

        pFilt = New QueryFilter
        pFilt.WhereClause = "INSERVICEDATE<=" & modelYEAR & " AND OUTSERVICEDATE>" & modelYEAR
        pFCS = pFCls.Search(pFilt, False)


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

    Private Function GetTransitPoints(ByVal pFCls As IFeatureClass, ByVal pRt As IPolyline4, ByVal lineIDField As String, ByVal lineIDValue As String, ByRef pPtCol As IPointCollection4, ByRef dct As Dictionary(Of Long, PSRCModelIO.jctAtt))
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
        'dct = New Dictionary
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
            dct.Add(l, att)

            pFt = pFCS.NextFeature
        Loop

        GetTransitPoints = dct

        pFCS = Nothing
        pSFilt = Nothing
    End Function

    Private Function SortSequence(ByVal pFClsE As IFeatureClass, ByVal pFClsJ As IFeatureClass, ByVal pFClsTP As IFeatureClass, ByVal plineFt As IFeature, _
        ByVal pSegTbl As ITable, ByVal dctSeq As Dictionary(Of Long, Long), ByVal modelYEAR As Long)
        Dim pSegCs As ICursor
        Dim pTrLine As IPolyline4
        Dim sLineID As String
        Dim sLineNo As String
        Dim pJctCol As IPointCollection4, pTPCol As IPointCollection4
        Dim dctE As New Dictionary(Of String, Long), dctTP As New Dictionary(Of Long, PSRCModelIO.jctAtt), dctN As New Dictionary(Of Long, Long)

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
        If dctSeq Is Nothing Then dctSeq = New Dictionary(Of Long, Long)
        dctSeq.Clear()

        dctN = New Dictionary(Of Long, Long)
        pSegCs = pSegTbl.Insert(True)
        'loop through all the vertices in the transit line
        Try

            Dim x As Integer
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
                    If dctN.ContainsKey(CType(CurJct, Long)) Then
                        'junctCount = junctCount - 1
                        '***********Need to create a counter that keeps track of transit points bette
                        x = dctN.Item(CType(CurJct, Long))
                        x = x + 1
                        dctN.Remove(CurJct)
                        dctN.Add(CType(CurJct, Long), x)
                        x = 0

                        'dctN.Item(CType(CurJct, Long)) = dctN.Item(CType(CurJct, Long) + 1)
                        ' dctN.ItemCurJct, Long)) = dctN.Item(CType(CurJct, Long) + 1)

                    Else
                        dctN.Add(CType(CurJct, Long), 1)
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
                                dctSeq.Add(dctSeq.Count, CType(CurJct, Long))
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
        If edgeCount < (dctE.Count - 3) Then
            PrintLine(2, sLineID)
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
    Public Sub create_TransitFile5(ByVal pathnameN As String, ByVal filenameN As String)

        '[080109] SEC: Created this to deal with Dwell Times. Older versions applied a dwell time from a TransitPoint to all
        'subsequent nodes until the next TransitPoint. This resulted in way more stops than there should. Also, DWT's are listed for the
        'Inode but the actual stop takes place at the Jnode portion of the link. Therefore, if you have a transit point with a positive dwell time
        'followed by several non transit-point nodes the dwell time needs to be applied to the node before the next transit point. This is all taken
        'care of in the code below.
        'Example:
        'TP
        '1 dwt=.25
        '5 dwt = .25
        '8 dwt =.25
        '10 lay = 0

        'actual link to link read out with non tp nodes
        '1 dwt=0
        '2 dwt=0
        '4 dwt=.25
        '5 dwt=.25
        '8 dwt=0
        '9 dwt= .25
        '10 lay = 0




        'Now , DWT Is only
        'taken and applied to nodes that are also transit points.

        '[051106] hyu: this routine is to replace "create_TransitFile_0"
        'The algorithm of traversing transit segments and writting nodes/weave nodes are different from the old one.
        'It dramatically reduced the running time from hours to minutes.

        'Dim pStatusBar As IStatusBar
        'pStatusBar = m_App.StatusBar

        'on error GoTo eh
        WriteLogLine("")
        WriteLogLine("========================================")
        WriteLogLine("create_TransitFile started " & Now())
        WriteLogLine("========================================")
        WriteLogLine("")

        'this must be called by createScenarioShapefile
        Dim runtype(5) As String
        runtype(0) = "AM"
        runtype(1) = "MD"
        runtype(2) = "PM"
        runtype(3) = "EV"
        runtype(4) = "NI"

        Dim thepath As String
        Dim tempfield As String
        Dim tempString As String
        Dim attrib(4) As String
        Try

            thepath = pathnameN + "\" + "am_" + filenameN
            FileOpen(1, thepath, OpenMode.Output, OpenAccess.Default, OpenShare.Shared)
            'Open thepath For Output As 1

            thepath = pathnameN + "\" + "md_" + filenameN
            FileOpen(2, thepath, OpenMode.Output, OpenAccess.Default, OpenShare.Shared)
            'Open thepath For Output As 2
            thepath = pathnameN + "\" + "pm_" + filenameN
            FileOpen(3, thepath, OpenMode.Output, OpenAccess.Default, OpenShare.Shared)
            'Open thepath For Output As 3
            thepath = pathnameN + "\" + "ev_" + filenameN
            FileOpen(4, thepath, OpenMode.Output, OpenAccess.Default, OpenShare.Shared)
            'Open thepath For Output As 4
            thepath = pathnameN + "\" + "ni_" + filenameN
            FileOpen(5, thepath, OpenMode.Output, OpenAccess.Default, OpenShare.Shared)

            'Open thepath For Output As 5

            PrintLine(1, "c Exported Transit Routes from ArcMap /Emme/2 Interface for Am: ") '+ CStr(MyDate)
            PrintLine(1, "c line   inode   jnode   dwfac   dwt   timtr   ttf    @ehdwy    result")
            PrintLine(2, "c Exported Transit Routes from ArcMap /Emme/2 Interface for Midday: ") '+ CStr(MyDate)
            PrintLine(2, "c line   inode   jnode   dwfac   dwt   timtr   ttf    @ehdwy    result")
            PrintLine(3, "c Exported Transit Routes from ArcMap /Emme/2 Interface for PM: ") '+ CStr(MyDate)
            PrintLine(3, "c line   inode   jnode   dwfac   dwt   timtr   ttf    @ehdwy    result")
            PrintLine(4, "c Exported Transit Routes from ArcMap /Emme/2 Interface for Evening: ") '+ CStr(MyDate)
            PrintLine(4, "c line   inode   jnode   dwfac   dwt   timtr   ttf    @ehdwy    result")
            PrintLine(5, "c Exported Transit Routes from ArcMap /Emme/2 Interface for Night: ") '+ CStr(MyDate)
            PrintLine(5, "c line   inode   jnode   dwfac   dwt   timtr   ttf    @ehdwy    result")
            PrintLine(2, "t lines init")
            PrintLine(1, "t lines init")
            PrintLine(3, "t lines init")
            PrintLine(4, "t lines init")
            PrintLine(5, "t lines init")
            Dim pFeatLayerE As IFeatureLayer, pTLineLayer As IFeatureLayer
            Dim pWS As IWorkspace
            Dim tblTSeg As ITable
            'Dim Pfltransitpoints As IFeatureLayer
            Dim pModeAtts As ITable
            Dim dctMAtt As New Dictionary(Of Long, Row), dctEdges As New Dictionary(Of Long, Feature)
            Dim dctAM_IJ_HOV As New Dictionary(Of Long, Long)
            Dim dctAM_JI_HOV As New Dictionary(Of Long, Long)
            Dim dctMD_IJ_HOV As New Dictionary(Of Long, Long)
            Dim dctMD_JI_HOV As New Dictionary(Of Long, Long)
            Dim dctNodeString As New Dictionary(Of Long, String)
            Dim pTblPrjEdge As ITable
            Dim pTblMode As ITable
            pTblPrjEdge = get_TableClass(m_layers(24)) 'use if future project
            pTblMode = get_TableClass(m_layers(2))
          'sec 072609
            'Pfltransitpoints = get_FeatureLayer("sde.SDE.TransitPoints")

            pWS = get_Workspace()
            '    Set pTLineLayer = get_FeatureLayer(m_layers(13))  'transitlines
            Dim pTLineFCls As IFeatureClass
            pTLineFCls = getInterimTransitLines()
            tblTSeg = get_TableClass(m_layers(14))  'tbltransitsegments
            pModeAtts = get_TableClass(m_layers(2))
            getEdgeAtts(m_edgeShp, pTblMode, pTblPrjEdge, g_PSRCEdgeID, dctMAtt)


            LoadHOVDict(dctMAtt, dctAM_IJ_HOV, dctAM_JI_HOV, dctMD_IJ_HOV, dctMD_JI_HOV)

            Dim pMARow As IRow
            Dim pEdgeID As Long
            Dim pEdgeIDFilter As IQueryFilter
            'Dim clsModeAtts As New clsModeAttributes(pMARow)
            'Dim pMACursor As ICursor



            Dim pQFtlines As IQueryFilter, pQF As IQueryFilter, pQF2 As IQueryFilter
            Dim pFeature As IFeature, pFeatTRoute As IFeature, pFeatE As IFeature
            Dim pFeatCursor As IFeatureCursor
            Dim pFCtroute As IFeatureCursor
            Dim pTC As ICursor
            Dim pRow As IRow
            Dim i As Long

            pQFtlines = New QueryFilter
            '    pQFtlines.WhereClause = "InServiceDate <= " + CStr(inserviceyear)
            '    Set pFCtroute = pTLineLayer.Search(pQFtlines, False)
            pFCtroute = pTLineFCls.Search(Nothing, False)
            pFeatTRoute = pFCtroute.NextFeature

            Dim index As Long
            Dim idIndex As Long, sIndex As Long
            '    Dim iNode As Long, jNode As Long

            'sort transit segment table by LineID and SegOrder, then loop through the sorted rows
            Dim pSort As ITableSort
            Dim dctTransitLine As New Dictionary(Of String, String)

            'dctTransitLine = New Dictionary
            pSort = New TableSort
            pQF = New QueryFilter

            If strLayerPrefix = "SDE" Then
                pQF.WhereClause = "SUBSTRING(CAST(LineID AS CHARACTER), 2, 2) = '" & Right(CStr(inserviceyear), 2) & "'" & "Or SUBSTRING(CAST(LineID AS CHARACTER), 2, 2) = '" & Right(CStr(inserviceyear + 1), 2) & "'"
            Else
                pQF.WhereClause = "Mid([LineID], 2, 2) = '" & Right(CStr(inserviceyear), 2) & "'"
            End If

            With pSort
                .Fields = "LineID, SegOrder"
                .Ascending("LineID") = True
                .Ascending("SegOrder") = True
                .QueryFilter = Nothing ' pQF
                .Table = tblTSeg
            End With
            pSort.Sort(Nothing)
            pTC = pSort.Rows

            'start statusbar
            Dim count As Long
            Dim sLineID As String
            '    count = pTLineLayer.FeatureClass.featurecount(pQFtlines)
            count = pTLineFCls.FeatureCount(Nothing)
            'pStatusBar.ShowProgressBar("Preparing for Transit...", 0, count, 1, True)
            'jaf--let's keep the user updated
            'frmNetLayer.lblStatus = "GlobalMod.create_Netfile: Starting Transit Build..."
            'pan frmNetLayer.Refresh

            'loop through searched transit Lines table, and get all transit info into a dictionary
            '    idIndex = pTLineLayer.FeatureClass.FindField("LineID")
            idIndex = pTLineFCls.FindField("LineID")
            sIndex = tblTSeg.FindField("SegOrder")
            Dim x As Long
            Dim y As Long

            Dim pTransitMode As String
            Dim intOperator As Integer
            Dim prevDwtStop As String = Nothing
            'Dim intTPCounter As Integer
            Do Until pFeatTRoute Is Nothing
                '[040307] hyu: per Jeff's email on [05/16/2006]
                'For each insertion of a new transit line in the buildfile, the application should
                'extract the first three digits of TransitLine.LineID to report as the time-of-day and year code
                'in a new comment line placed just above the buildfile line that inserts the transit line.
                'All remaining TransitLine.LineID digits to the right of the first three
                'should be written as the line ID in the route insertion line.

                'pTransitMode = pFeatTRoute.value(pFeatTRoute.Fields.FindField("Mode"))





                sLineID = CStr(pFeatTRoute.Value(idIndex))



                If Len(sLineID) > 4 Then
                    tempString = "c '" + Left(sLineID, 2) + "'" + vbNewLine
                    tempString = tempString + "a '" + Mid(sLineID, 3, Len(sLineID) - 2) + "'"
                Else
                    tempString = "a '" + sLineID + "'"
                End If
                'tempString = "a '" + CStr(pFeatTRoute.value(idIndex)) + "'"

                index = pFeatTRoute.Fields.FindField("Mode")
                tempString = tempString + " " + CStr(pFeatTRoute.Value(index))
                'index = pFeatTRoute.Fields.FindField("VehicleType")
                index = pFeatTRoute.Fields.FindField(Left("VehicleType", 10))
                If Not IsDBNull(pFeatTRoute.Value(index)) Then
                    tempString = tempString + " " + CStr(pFeatTRoute.Value(index))
                Else
                    tempString = tempString + " 1" 'need to get better default
                End If
                index = pFeatTRoute.Fields.FindField("Headway")
                If Not IsDBNull(pFeatTRoute.Value(index)) Then
                    tempString = tempString + " " + CStr(pFeatTRoute.Value(index))
                Else
                    tempString = tempString + " 1" 'need to get better default
                End If
                index = pFeatTRoute.Fields.FindField("Speed")
                If Not IsDBNull(pFeatTRoute.Value(index)) Then
                    tempString = tempString + " " + CStr(pFeatTRoute.Value(index))
                Else
                    tempString = tempString + " 55" 'need to get better default
                End If
                'index = pFeatTRoute.Fields.FindField("Description")
                index = pFeatTRoute.Fields.FindField(Left("Description", 10))
                If Not IsDBNull(pFeatTRoute.Value(index)) Then
                    tempString = tempString + " '" + CStr(pFeatTRoute.Value(index)) + "'"
                Else
                    tempString = tempString + " 'null'"
                End If
                'index = pFeatTRoute.Fields.FindField("Geography")
                'If Not IsNull(pFeatTRoute.value(index)) Then
                'tempString = tempString + " " + CStr(pFeatTRoute.value(index))
                'Else
                'tempString = tempString + " 1"
                'End If
                'index = pFeatTRoute.Fields.FindField("Oneway")
                'If Not IsNull(pFeatTRoute.value(index)) Then
                'tempString = tempString + " " + CStr(pFeatTRoute.value(index))
                'Else
                'tempString = tempString + " 1"
                'End If
                index = pFeatTRoute.Fields.FindField("Processing")
                If Not IsDBNull(pFeatTRoute.Value(index)) Then
                    tempString = tempString + " " + CStr(pFeatTRoute.Value(index))
                Else
                    tempString = tempString + " 0"
                End If

                index = pFeatTRoute.Fields.FindField("UL2")
                If Not IsDBNull(pFeatTRoute.Value(index)) Then
                    tempString = tempString + " " + CStr(pFeatTRoute.Value(index))
                Else
                    tempString = tempString + " 1"
                End If

                index = pFeatTRoute.Fields.FindField("Operator")
                If Not IsDBNull(pFeatTRoute.Value(index)) Then
                    tempString = tempString + " " + CStr(pFeatTRoute.Value(index))
                Else
                    tempString = tempString + " 1"
                End If

                index = pFeatTRoute.Fields.FindField("TimePeriod")
                If Not IsDBNull(pFeatTRoute.Value(index)) Then
                    tempString = CStr(pFeatTRoute.Value(index)) + tempString
                Else
                    tempString = "0" + tempString
                End If

                If Not dctTransitLine.ContainsKey(CStr(pFeatTRoute.Value(idIndex))) Then dctTransitLine.Add(CStr(pFeatTRoute.Value(idIndex)), tempString)

                'pStatusBar.StepProgressBar()
                pFeatTRoute = pFCtroute.NextFeature
            Loop

            'now get starting segment
            '    Dim skip As Integer, cindex As Integer
            Dim preNode As Long, preLine As Long, curNode As Long, nextNode As Long, preWNode As Long, curWNode As Long, nextWNode As Long, preWNode2 As Long
            Dim lSegOrder As Long, lLineID As Long, LineInfo As String, lUseGP As Long, lTOD As Long
            Dim fldSegOrder As Long, fldLineId As Long, fldINode As Long, fldJNode As Long, fldUseGP As Long
            Dim fldPath As Long, fldDwtStop As Long, fldtimeFuncID As Long, fldLayover As Long, fldUser1 As Long, fldUser2 As Long, fldUser3 As Long
            Dim sPath As String, sDwtStop As String, stimeFuncID As String, sLayover As String, sUser1 As String, sUser2 As String, sUser3 As String
            Dim PreWeaveString As String, lastNodeString As String
            Dim bPreWeave As Boolean
            Dim sDwtStopJ As String
            Dim tempLastLineString As String
            With pTC
                fldINode = .FindField(g_INode)
                fldJNode = .FindField(g_JNode)
                fldSegOrder = .FindField("SegOrder")
                fldLineId = .FindField("LineID")
                fldPath = .FindField("Path")
                fldDwtStop = .FindField("DwtStop")
                fldtimeFuncID = .FindField("timeFuncID")
                fldLayover = .FindField("LayOver")
                fldUser1 = .FindField("User1")
                fldUser2 = .FindField("User2")
                fldUser3 = .FindField("User3")
                fldUseGP = .FindField("UseGPOnly")
            End With



            pRow = pTC.NextRow


            If pRow Is Nothing Then
                WriteLogLine("No Transit Line selected. End Creating Transit buildfiles")
                'pStatusBar.HideProgressBar()
                Exit Sub
            End If
            'preLine = pRow.value(pRow.Fields.FindField("LineID"))
            preLine = 0






            'get all nodes in the intermediate layer
            Dim dctNodes As New Dictionary(Of String, String)

            'dctNodes = New Dictionary
            getAllNodes2(dctNodes)
            ' pStatusBar.ShowProgressBar("Creating Transit...", 0, tblTSeg.rowcount(pQF), 1, True)
            Dim intTPCounter As Long
            Dim dctTransitPoints As Dictionary(Of Long, Long)
            Dim dctDwellTimes As Dictionary(Of Long, String)
            Dim dctStopDistance As Dictionary(Of Long, Long)

            'Dim x As Long
            x = 0
            Dim intNodeCounter As Long = 1
            Do Until pRow Is Nothing
                'check to see if this is a new line, if so get it's mode
                If preLine = 0 Or preLine <> pRow.Value(pRow.Fields.FindField("LineID")) Then
                    Dim pFilt As IQueryFilter
                    Dim pFeat As IFeature
                    Dim pFCur As IFeatureCursor
                    pFilt = New QueryFilter
                    pFilt.WhereClause = "LineID = " & pRow.Value(pRow.Fields.FindField("LineID"))
                    pFCur = pTLineFCls.Search(pFilt, True)
                    pFeat = pFCur.NextFeature
                    preLine = pRow.Value(pRow.Fields.FindField("LineID"))

                    pTransitMode = pFeat.Value(pFeat.Fields.FindField("Mode"))

                End If


                'pStatusBar.StepProgressBar()

                lLineID = pRow.Value(fldLineId)
                If x = 0 Then
                    pQFtlines = New QueryFilter
                    pQFtlines.WhereClause = "LineID = " + CStr(lLineID)
                    'Set pFCtroute = pTLineLayer.Search(pQFtlines, False)
                    pFCtroute = pTLineFCls.Search(pQFtlines, False)
                    pFeatTRoute = pFCtroute.NextFeature
                    pTransitMode = pFeatTRoute.Value(pFeatTRoute.Fields.FindField("Mode"))
                    intOperator = pFeatTRoute.Value(pFeatTRoute.Fields.FindField("Operator"))
                End If


                'sec 073009
                'get all the TransitPoints that belong to the route, store them in a dictionary
                'store their DWTs in another dictionary
                If x <> pRow.Value(fldLineId) Then
                    dctTransitPoints = New Dictionary(Of Long, Long)
                    dctDwellTimes = New Dictionary(Of Long, String)
                    dctStopDistance = New Dictionary(Of Long, Long)
                    GetTransitPointsByLineID2(dctTransitPoints, dctDwellTimes, Pfltransitpoints.FeatureClass, lLineID)
                    intTPCounter = 1
                    x = pRow.Value(fldLineId)
                    GetStopDistances2(lLineID, dctStopDistance, tblTSeg, dctTransitPoints)
                End If
                Do Until dctTransitLine.ContainsKey(CStr(lLineID))
                    'if transit info doesn't exist, then skip all the records of this line.
                    pRow = pTC.NextRow
                    'pStatusBar.StepProgressBar()

                    If pRow Is Nothing Then
                        'Close()

                        Exit Sub
                    End If
                    lLineID = pRow.Value(fldLineId)
                Loop

                lSegOrder = pRow.Value(fldSegOrder)

                '[040207]hyu: per Jeff's email on [05/16/06]: path=no or path=yes where TransitLines.Path=0 signifies no and TransitLines.Path=1 signifies yes.
                sPath = IIf(pRow.Value(fldPath) = 1, " path=yes", IIf(pRow.Value(fldPath) = 0, " path=no", ""))

                'sDwtStop = " dwt=" + CStr(pRow.value(fldDwtStop))
                'check to see if the Inode of the segment is a transit point for the transit route:


                Dim strDwell As String
                Dim intPrevTN As Integer
                Dim lngStopDistance As Long

                Dim prevNodeString As String

                'see if the current node is a transit point
                Try


                    If dctTransitPoints.Item(intTPCounter) = pRow.Value(fldINode) Then
                        If dctStopDistance.ContainsKey(intTPCounter + 1) Then


                            lngStopDistance = dctStopDistance.Item(intTPCounter) - dctStopDistance.Item(intTPCounter + 1)
                            strDwell = dctDwellTimes.Item(intTPCounter)
                            intPrevTN = intTPCounter
                            intTPCounter = intTPCounter + 1
                        Else
                            WriteLogLine("Line Failure at: " & lLineID & " " & intTPCounter + 1)
                            intTPCounter = intTPCounter + 1
                        End If
                    End If

                    If dctTransitPoints.ContainsKey(intPrevTN + 1) Then
                        If dctTransitPoints.Item(intPrevTN + 1) = pRow.Value(fldJNode) Then
                            sDwtStop = " " & strDwell
                        Else
                            sDwtStop = " dwt=#.00"
                        End If
                    Else
                        WriteLogLine("Line Failure at: " & lLineID & " " & intTPCounter + 1)
                    End If

                    'sDwtStop = " " & GetDwellTime2(lLineID, intTPCounter)
                    'if so, grab the DWT from the DWT dictionary
                Catch ex As Exception
                    MessageBox.Show(ex.ToString)

                End Try



                'deal with the last node
                If dctTransitPoints.Item(dctTransitPoints.Count) = pRow.Value(fldJNode) Then
                    'sDwtStopJ = " " & CStr(GetDwellTime2(lLineID, dctTransitPoints.count))
                    sDwtStopJ = " " & dctDwellTimes.Item(dctDwellTimes.Count)
                End If

                'sDwtStopJ = " " & CStr(GetDwellTime(Pfltransitpoints.FeatureClass, lLineID, pRow.value(fldJNode)))
                stimeFuncID = " ttf=" + CStr(pRow.Value(fldtimeFuncID))
                sLayover = IIf(pRow.Value(fldLayover) > 0, " lay=" + CStr(pRow.Value(fldLayover)), "")
                sUser1 = " us1=" + CStr(IIf(IsDBNull(pRow.Value(fldUser1)), 0, "0"))
                sUser2 = " us2=" + CStr(IIf(IsDBNull(pRow.Value(fldUser2)), 0, pRow.Value(fldUser2)))
                sUser3 = " us3=" + CStr(IIf(IsDBNull(pRow.Value(fldUser3)), 0, pRow.Value(fldUser3)))
                tempString = sDwtStop + stimeFuncID + sLayover + sUser1 + sUser2 + sUser3
                tempLastLineString = sDwtStopJ + stimeFuncID + sLayover + sUser1 + sUser2 + sUser3
                If lSegOrder = 1 Then
                    'start a new line, then get the transit line info

                    LineInfo = dctTransitLine.Item(CStr(lLineID))
                    lTOD = Left(LineInfo, 1)
                    LineInfo = Right(LineInfo, Len(LineInfo) - 1)

                    i = 0

                    'get the first node
                    curNode = pRow.Value(fldINode)
                    nextNode = pRow.Value(fldJNode)
                    preNode = 0 'curNode
                    preWNode = 0
                    curWNode = 0
                    PreWeaveString = ""
                    lastNodeString = ""
                    bPreWeave = False

                    writeTransitNode(lTOD, LineInfo)

                    If sPath <> "" Then writeTransitNode(lTOD, sPath)
                End If

                curNode = pRow.Value(fldINode)
                nextNode = pRow.Value(fldJNode)

                If curNode <> preNode Then
                    If Not dctNodes.ContainsKey(CStr(curNode)) Then
                        If fVerboseLog Then WriteLogLine("Line " & lLineID & " Node " & curNode & " SegOrder=" & lSegOrder & " dissolved")

                    Else    'dctNodes.Exists(CStr(curNode))
                        'form an edge of preNode-curNode
                        'check whether it should use the GP/TR/HOV lane
                        'if the nodes dosn't exist, skip it.
                        lUseGP = pRow.Value(fldUseGP)

                        If lUseGP = 0 Then
                            If fVerboseLog Then WriteLogLine("Line " & lLineID & " Node " & curNode & " SegOrder=" & lSegOrder & " UseGPOnly=0")

                            pQF2 = New QueryFilter
                            pQF2.WhereClause = "(" + g_INode + "=" + CStr(curNode) + " AND " + g_JNode + "=" + CStr(nextNode) + ") OR (" _
                                + g_JNode + "=" + CStr(curNode) + " AND " + g_INode + "=" + CStr(nextNode) + ")"
                            pFeatCursor = m_edgeShp.Search(pQF2, False)
                            pFeature = pFeatCursor.NextFeature
                            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(pFeatCursor)

                            If Not pFeature Is Nothing Then
                                'Do TTF stuff here
                                If intOperator = 4 Then
                                    'do nothing
                                Else
                                    'pTransitMode = pFeature.Fields.FindField("NewFacilityType")
                                    lngStopDistance = Math.Abs(lngStopDistance)
                                    Dim strSegmentMode As String
                                    strSegmentMode = pFeature.Value(pFeature.Fields.FindField("Modes"))
                                    If pTransitMode = "r" Or pTransitMode = "c" Or pTransitMode = "f" Then
                                        stimeFuncID = 5
                                    ElseIf strSegmentMode = "bp" Or strSegmentMode = "bwlp" Or strSegmentMode = "brp" Or strSegmentMode = "bwp" Then
                                        stimeFuncID = 4
                                    ElseIf lngStopDistance > 7920 Then
                                        stimeFuncID = 14
                                    ElseIf (pFeature.Value(pFeature.Fields.FindField("NewFacilityType")) = 1 Or pFeature.Value(pFeature.Fields.FindField("NewFacilityType")) = 2) And lngStopDistance > 2640 Then
                                        stimeFuncID = 13
                                    ElseIf lngStopDistance > 2640 Then
                                        stimeFuncID = 12
                                    Else : stimeFuncID = 11

                                    End If
                                    stimeFuncID = " ttf=" + stimeFuncID
                                End If


                                tempString = sDwtStop + stimeFuncID + sLayover + sUser1 + sUser2 + sUser3
                                tempLastLineString = sDwtStopJ + stimeFuncID + sLayover + sUser1 + sUser2 + sUser3


                                'Does the Transit Segment goe in the IJ or JI direciton
                                'Does the Transit Segment goe in the IJ or JI direciton
                                If pFeature.Value(m_edgeShp.FindField("INode")) = curNode Then 'JI
                                    If Not IsDBNull(pFeature.Value(m_edgeShp.FindField("TR_I"))) Then

                                        'If pFeature.Value(m_edgeShp.FindField("TR_I")) > 0 Then
                                        'weave nodes from "TR_I" and "TR_J"
                                        curWNode = pFeature.Value(m_edgeShp.FindField("TR_I")) + m_Offset
                                        nextWNode = pFeature.Value(m_edgeShp.FindField("TR_J")) + m_Offset

                                        If fVerboseLog Then WriteLogLine("potential weave nodes (TR_J/I) " & preWNode & ", " & curWNode)
                                        'End If
                                    Else
                                        'weave nodes may from "HOV_I/J"
                                        Select Case pFeature.Value(m_edgeShp.FindField("FacilityType"))
                                            Case 11, 12, 3, 13, 9
                                                curWNode = 0
                                                nextWNode = 0
                                            Case Else
                                                pEdgeID = pFeature.Value(m_edgeShp.FindField("PSRCEdgeID"))
                                                'pEdgeIDFilter = New QueryFilter
                                                'pEdgeIDFilter.WhereClause = "PSRCEdgeID = " & pEdgeID
                                                'pMACursor = pModeAtts.Search(pEdgeIDFilter, True)
                                                '
                                                'pMARow = pMACursor.NextRow
                                                'System.Runtime.InteropServices.Marshal.ReleaseComObject(pMACursor)
                                                '111309- Previous code assumed that an edge flagged HOV in ScenarioEdge is HOV for both directions at all times of day
                                                'the following makes sure the at the edge is HOV for TOD and direction
                                                'If IsHOV(pMARow, Left(CStr(lLineID), 1), 1) Then
                                                If Left(CStr(lLineID), 1) = "1" And dctAM_IJ_HOV.ContainsKey(pEdgeID) Then
                                                    ' clsModeAtts.modeAttributeRow = pMARow
                                                    'MsgBox (clsModeAtts.IJLANESHOVAM)
                                                    curWNode = IIf(IsDBNull(pFeature.Value(m_edgeShp.FindField("HOV_I"))), 0, pFeature.Value(m_edgeShp.FindField("HOV_I")) + m_Offset)
                                                    nextWNode = IIf(IsDBNull(pFeature.Value(m_edgeShp.FindField("HOV_J"))), 0, pFeature.Value(m_edgeShp.FindField("HOV_J")) + m_Offset)
                                                ElseIf Left(CStr(lLineID), 1) = "2" And dctMD_IJ_HOV.ContainsKey(pEdgeID) Then
                                                    curWNode = IIf(IsDBNull(pFeature.Value(m_edgeShp.FindField("HOV_I"))), 0, pFeature.Value(m_edgeShp.FindField("HOV_I")) + m_Offset)
                                                    nextWNode = IIf(IsDBNull(pFeature.Value(m_edgeShp.FindField("HOV_J"))), 0, pFeature.Value(m_edgeShp.FindField("HOV_J")) + m_Offset)
                                                End If
                                        End Select

                                        If fVerboseLog Then WriteLogLine("potential weave nodes (HOV_J/I) " & curWNode & ", " & nextWNode)
                                    End If

                                Else 'Transit segment goes in the JI direction
                                    If Not IsDBNull(pFeature.Value(m_edgeShp.FindField("TR_J"))) Then
                                        'If pFeature.Value(m_edgeShp.FindField("TR_J")) > 0 Then
                                        'weave nodes from "TR_I" and "TR_J"
                                        curWNode = pFeature.Value(m_edgeShp.FindField("TR_J")) + m_Offset
                                        nextWNode = pFeature.Value(m_edgeShp.FindField("TR_I")) + m_Offset

                                        'End If
                                        'End If
                                        If fVerboseLog Then WriteLogLine("potential weave nodes (TR_I/J) " & curWNode & ", " & nextWNode)

                                    Else
                                        'weave nodes may from "HOV_I/J"
                                        Select Case pFeature.Value(m_edgeShp.FindField("FacilityType"))
                                            Case 11, 12, 3, 13, 9
                                                curWNode = 0
                                                nextWNode = 0
                                            Case Else
                                                pEdgeID = pFeature.Value(m_edgeShp.FindField("PSRCEdgeID"))
                                                'pEdgeIDFilter = New QueryFilter
                                                'pEdgeIDFilter.WhereClause = "PSRCEdgeID = " & pEdgeID
                                                'pMACursor = pModeAtts.Search(pEdgeIDFilter, False)
                                                'pMARow = pMACursor.NextRow
                                                ' System.Runtime.InteropServices.Marshal.FinalReleaseComObject(pMACursor)
                                                ' pMACursor = Nothing
                                                'If IsHOV(pMARow, Left(CStr(lLineID), 1), 2) Then
                                                'curWNode = IIf(IsDBNull(pFeature.Value(m_edgeShp.FindField("HOV_J"))), 0, pFeature.Value(m_edgeShp.FindField("HOV_J")) + m_Offset)
                                                'nextWNode = IIf(IsDBNull(pFeature.Value(m_edgeShp.FindField("HOV_I"))), 0, pFeature.Value(m_edgeShp.FindField("HOV_I")) + m_Offset)
                                                'End If
                                                If Left(CStr(lLineID), 1) = "1" And dctAM_JI_HOV.ContainsKey(pEdgeID) Then
                                                    ' clsModeAtts.modeAttributeRow = pMARow
                                                    'MsgBox (clsModeAtts.IJLANESHOVAM)
                                                    curWNode = IIf(IsDBNull(pFeature.Value(m_edgeShp.FindField("HOV_J"))), 0, pFeature.Value(m_edgeShp.FindField("HOV_J")) + m_Offset)
                                                    nextWNode = IIf(IsDBNull(pFeature.Value(m_edgeShp.FindField("HOV_I"))), 0, pFeature.Value(m_edgeShp.FindField("HOV_I")) + m_Offset)
                                                ElseIf Left(CStr(lLineID), 1) = "2" And dctMD_JI_HOV.ContainsKey(pEdgeID) Then
                                                    curWNode = IIf(IsDBNull(pFeature.Value(m_edgeShp.FindField("HOV_J"))), 0, pFeature.Value(m_edgeShp.FindField("HOV_J")) + m_Offset)
                                                    nextWNode = IIf(IsDBNull(pFeature.Value(m_edgeShp.FindField("HOV_I"))), 0, pFeature.Value(m_edgeShp.FindField("HOV_I")) + m_Offset)
                                                End If

                                        End Select

                                        If fVerboseLog Then WriteLogLine("potential weave nodes (HOV_I/J) " & preWNode & ", " & curWNode)
                                    End If
                                End If 'pfeature.value(m_edgeShp.FindField("INode")) = curNode



                                If bPreWeave And preWNode = curWNode And curWNode <> 0 Then
                                    'we are on hov system, write out the next node because current one has alreayd been written. We
                                    'are one step ahead here. 
                                    'Use next weave node, but have to make sure dwell time is right!
                                    'if next node is a stop we need to assign it to the previous node since we are one node ahead here!

                                    'writeTransitNode(lTOD, " " + CStr(nextWNode) + tempString)
                                    If dctTransitPoints.Item(intPrevTN + 1) = pRow.Value(fldJNode) Then
                                        'sDwtStop = " " & strDwell
                                        'sDwtStop is set from above and should indicate a stop. We need to apply this to the previous node. 

                                        
                                        prevNodeString = dctNodeString.Item(intNodeCounter - 1)
                                        prevNodeString = prevNodeString.Replace(LTrim(prevDwtStop), LTrim(sDwtStop))
                                        dctNodeString.Item(intNodeCounter - 1) = prevNodeString


                                        'Stop should not take place here. The dictionary will be updated next iteration if it is indeed a stop. 
                                        tempString = " dwt=#.00" + stimeFuncID + sLayover + sUser1 + sUser2 + sUser3

                                    End If
                                    dctNodeString.Add(intNodeCounter, " " + CStr(nextWNode) + tempString)
                                    intNodeCounter = intNodeCounter + 1
                                    bPreWeave = True
                                Else
                                    'On GP lanes, did we just get off HOV?
                                    'If bPreWeave And pRow.Value(fldINode) = dctTransitPoints(intPrevTN) Then
                                    'if yes, have to update last node with proper stop info:
                                    'NodeString = dctNodeString.Item(intNodeCounter - 1)
                                    'prevNodeString = System.Text.RegularExpressions.Regex.Replace(prevNodeString, "dwt=*....", LTrim(dctDwellTimes.Item(intTPCounter - 1)))


                                    'prevNodeString = prevNodeString.Replace(LTrim(prevDwtStop), LTrim(dctDwellTimes.Item(intTPCounter - 1)))
                                    'dctNodeString.Item(intNodeCounter - 1) = prevNodeString
                                    'End If

                                    'see if transit line is entering HOV/Managed lanes, if so, write out I & J Nodes of HOV:
                                    If curWNode > 0 And nextWNode > 0 Then
                                        'this is the gp node. If the next node is a stop, need to put the dwll time on the INode of the HOV, not here. 
                                        tempString = " dwt=#.00" + stimeFuncID + sLayover + sUser1 + sUser2 + sUser3
                                        'writeTransitNode(lTOD, " " + CStr(dctNodes.Item(CStr(curNode))) + tempString)
                                        dctNodeString.Add(intNodeCounter, " " + CStr(dctNodes.Item(CStr(curNode))) + tempString)
                                        intNodeCounter = intNodeCounter + 1

                                        'HOV INode of the first HOV link. Put dwell time here if the next node is a stop. 
                                        tempString = sDwtStop + stimeFuncID + sLayover + sUser1 + sUser2 + sUser3
                                        'writeTransitNode(lTOD, " " + CStr(curWNode) + tempString)
                                        dctNodeString.Add(intNodeCounter, " " + CStr(curWNode) + tempString)
                                        intNodeCounter = intNodeCounter + 1

                                        'HOV JNode of the first HOV link
                                        'if next node on the next link is indeed a stop, this dictionary will get updated. 
                                        tempString = " dwt=#.00" + stimeFuncID + sLayover + sUser1 + sUser2 + sUser3
                                        'writeTransitNode(lTOD, " " + CStr(nextWNode) + tempString)
                                        dctNodeString.Add(intNodeCounter, " " + CStr(nextWNode) + tempString)
                                        intNodeCounter = intNodeCounter + 1
                                        bPreWeave = True
                                    Else
                                        bPreWeave = False
                                        'was on a GP and still is 
                                        'writeTransitNode(lTOD, " " + CStr(dctNodes.Item(CStr(curNode))) + tempString)
                                        dctNodeString.Add(intNodeCounter, " " + CStr(dctNodes.Item(CStr(curNode))) + tempString)
                                        intNodeCounter = intNodeCounter + 1
                                    End If
                                End If
                                    'sec 072909- fixing DWTs
                                    'lastNodeString = " " + CStr(dctNodes.Item(CStr(nextNode))) + tempString
                                    lastNodeString = " " + CStr(dctNodes.Item(CStr(nextNode))) + tempLastLineString

                                    preNode = curNode
                                    preWNode = nextWNode
                                    curWNode = 0
                                    nextWNode = 0

                            Else 'pfeature Is Nothing
                                    If i = 0 Then
                                        WriteLogLine("Data Error: Segment " + CStr(lSegOrder) + " on transit line " + CStr(lLineID) + " underlying TransRefEdge not in service")
                                        WriteLogLine("Skip the rest of line " & lLineID)
                                        i = 1
                                    End If
                            End If  'Not pfeature Is Nothing
                        Else    'lUseGP <> 0
                            'GP only

                            'SEC 072909- fixing DWTs
                            'lastNodeString = " " + CStr(dctNodes.Item(CStr(nextNode))) + tempString
                            lastNodeString = " " + CStr(dctNodes.Item(CStr(nextNode))) + tempLastLineString
                            'writeTransitNode(lTOD, " " + CStr(dctNodes.Item(CStr(curNode))) + tempString)
                            dctNodeString.Add(intNodeCounter, " " + CStr(dctNodes.Item(CStr(curNode))) + tempString)
                            intNodeCounter = intNodeCounter + 1
                            bPreWeave = False
                            'in case we need to update the previous string

                            preNode = curNode
                            curNode = nextNode
                            nextNode = 0
                            curWNode = 0
                            preWNode = 0
                            nextWNode = 0
                        End If  'lUseGP = 0
                    End If  'Not dctNodes.Exists(CStr(curNode))
                End If  'curNode <> preNode
                pRow = pTC.NextRow

                If pRow Is Nothing Then


                    'now is at the end, write the last node
                    'writeTransitNode(lTOD, lastNodeString)
                    dctNodeString.Add(intNodeCounter, lastNodeString)

                    For y = 1 To dctNodeString.Count
                        writeTransitNode(lTOD, dctNodeString.Item(y))
                    Next y
                    dctNodeString.Clear()
                    intNodeCounter = 1
                Else
                    If pRow.Value(fldLineId) <> lLineID Then
                        'a new transit line starts. so the unwritten last node of the previous line should be written now
                        'writeTransitNode(lTOD, lastNodeString)
                        dctNodeString.Add(intNodeCounter, lastNodeString)
                        For y = 1 To dctNodeString.Count
                            writeTransitNode(lTOD, dctNodeString.Item(y))
                        Next y
                        dctNodeString.Clear()

                        intNodeCounter = 1
                    End If
                End If
                prevDwtStop = sDwtStop
            Loop

            'pStatusBar.HideProgressBar()
            FileClose(1, 2, 3, 4, 5)
            pFeatLayerE = Nothing
            '   pTLineLayer = Nothing
            pTLineFCls = Nothing
            pFeature = Nothing
            WriteLogLine("FINISHED create_TransitFile at " & Now())
            dctTransitLine.Clear()
            dctNodes.Clear()
            dctTransitLine = Nothing
            dctNodes = Nothing




            'Exit Sub
            'eh:



        Catch ex As Exception
            MessageBox.Show(ex.ToString)

            CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.create_TransitFile")
            MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.create_TransitFile")
            'Print #9, Err.Description, vbExclamation, "TransitFile"
            FileClose(1, 2, 3, 4, 5)
            'Close 9

        End Try

    End Sub
   
    Public Sub create_TransitFile_NewSchema2(ByVal pathnameN As String, ByVal filenameN As String)

        '[080109] SEC: Created this to deal with Dwell Times. Older versions applied a dwell time from a TransitPoint to all
        'subsequent nodes until the next TransitPoint. This resulted in way more stops than there should. Also, DWT's are listed for the
        'Inode but the actual stop takes place at the Jnode portion of the link. Therefore, if you have a transit point with a positive dwell time
        'followed by several non transit-point nodes the dwell time needs to be applied to the node before the next transit point. This is all taken
        'care of in the code below.
        'Example:
        'TP
        '1 dwt=.25
        '5 dwt = .25
        '8 dwt =.25
        '10 lay = 0

        'actual link to link read out with non tp nodes
        '1 dwt=0
        '2 dwt=0
        '4 dwt=.25
        '5 dwt=.25
        '8 dwt=0
        '9 dwt= .25
        '10 lay = 0




        'Now , DWT Is only
        'taken and applied to nodes that are also transit points.

        '[051106] hyu: this routine is to replace "create_TransitFile_0"
        'The algorithm of traversing transit segments and writting nodes/weave nodes are different from the old one.
        'It dramatically reduced the running time from hours to minutes.

        'Dim pStatusBar As IStatusBar
        'pStatusBar = m_App.StatusBar

        'on error GoTo eh
        WriteLogLine("")
        WriteLogLine("========================================")
        WriteLogLine("create_TransitFile started " & Now())
        WriteLogLine("========================================")
        WriteLogLine("")

        'this must be called by createScenarioShapefile
        Dim runtype(5) As String
        runtype(0) = "AM"
        runtype(1) = "MD"
        runtype(2) = "PM"
        runtype(3) = "EV"
        runtype(4) = "NI"

        Dim thepath As String
        Dim tempfield As String
        Dim tempString As String
        Dim attrib(4) As String
        Try

            thepath = pathnameN + "\" + "am_" + filenameN
            FileOpen(1, thepath, OpenMode.Output, OpenAccess.Default, OpenShare.Shared)
            'Open thepath For Output As 1

            thepath = pathnameN + "\" + "md_" + filenameN
            FileOpen(2, thepath, OpenMode.Output, OpenAccess.Default, OpenShare.Shared)
            'Open thepath For Output As 2
            thepath = pathnameN + "\" + "pm_" + filenameN
            FileOpen(3, thepath, OpenMode.Output, OpenAccess.Default, OpenShare.Shared)
            'Open thepath For Output As 3
            thepath = pathnameN + "\" + "ev_" + filenameN
            FileOpen(4, thepath, OpenMode.Output, OpenAccess.Default, OpenShare.Shared)
            'Open thepath For Output As 4
            thepath = pathnameN + "\" + "ni_" + filenameN
            FileOpen(5, thepath, OpenMode.Output, OpenAccess.Default, OpenShare.Shared)

            'Open thepath For Output As 5

            PrintLine(1, "c Exported Transit Routes from ArcMap /Emme/2 Interface for Am: ") '+ CStr(MyDate)
            PrintLine(1, "c line   inode   jnode   dwfac   dwt   timtr   ttf    @ehdwy    result")
            PrintLine(2, "c Exported Transit Routes from ArcMap /Emme/2 Interface for Midday: ") '+ CStr(MyDate)
            PrintLine(2, "c line   inode   jnode   dwfac   dwt   timtr   ttf    @ehdwy    result")
            PrintLine(3, "c Exported Transit Routes from ArcMap /Emme/2 Interface for PM: ") '+ CStr(MyDate)
            PrintLine(3, "c line   inode   jnode   dwfac   dwt   timtr   ttf    @ehdwy    result")
            PrintLine(4, "c Exported Transit Routes from ArcMap /Emme/2 Interface for Evening: ") '+ CStr(MyDate)
            PrintLine(4, "c line   inode   jnode   dwfac   dwt   timtr   ttf    @ehdwy    result")
            PrintLine(5, "c Exported Transit Routes from ArcMap /Emme/2 Interface for Night: ") '+ CStr(MyDate)
            PrintLine(5, "c line   inode   jnode   dwfac   dwt   timtr   ttf    @ehdwy    result")
            PrintLine(2, "t lines init")
            PrintLine(1, "t lines init")
            PrintLine(3, "t lines init")
            PrintLine(4, "t lines init")
            PrintLine(5, "t lines init")
            Dim pFeatLayerE As IFeatureLayer, pTLineLayer As IFeatureLayer
            Dim pWS As IWorkspace
            Dim tblTSeg As ITable
            'Dim Pfltransitpoints As IFeatureLayer
            Dim pModeAtts As ITable
            Dim dctMAtt As New Dictionary(Of Long, Row), dctEdges As New Dictionary(Of Long, Feature)
            Dim dctAM_IJ_HOV As New Dictionary(Of Long, Long)
            Dim dctAM_JI_HOV As New Dictionary(Of Long, Long)
            Dim dctMD_IJ_HOV As New Dictionary(Of Long, Long)
            Dim dctMD_JI_HOV As New Dictionary(Of Long, Long)
            Dim dctPM_IJ_HOV As New Dictionary(Of Long, Long)
            Dim dctPM_JI_HOV As New Dictionary(Of Long, Long)
            Dim dctEV_IJ_HOV As New Dictionary(Of Long, Long)
            Dim dctEV_JI_HOV As New Dictionary(Of Long, Long)
            Dim dctNI_IJ_HOV As New Dictionary(Of Long, Long)
            Dim dctNI_JI_HOV As New Dictionary(Of Long, Long)
            Dim dctNodeString As New Dictionary(Of Long, String)
            Dim pTblPrjEdge As ITable
            Dim pTblMode As ITable
            pTblPrjEdge = get_TableClass(m_layers(24)) 'use if future project
            pTblMode = get_TableClass(m_layers(2))
            'sec 072609
            'Pfltransitpoints = get_FeatureLayer("sde.SDE.TransitPoints")

            pWS = get_Workspace()
            '    Set pTLineLayer = get_FeatureLayer(m_layers(13))  'transitlines
            Dim pTLineFCls As IFeatureClass
            pTLineFCls = getInterimTransitLines()
            tblTSeg = get_TableClass(m_layers(14))  'tbltransitsegments
            pModeAtts = get_TableClass(m_layers(2))
            getEdgeAtts(m_edgeShp, pTblMode, pTblPrjEdge, g_PSRCEdgeID, dctMAtt)


            LoadHOVDict(dctMAtt, dctAM_IJ_HOV, dctAM_JI_HOV, dctMD_IJ_HOV, dctMD_JI_HOV, dctPM_IJ_HOV, dctPM_JI_HOV, dctEV_IJ_HOV, dctEV_JI_HOV, dctNI_IJ_HOV, dctNI_JI_HOV)

            Dim pMARow As IRow
            Dim pEdgeID As Long
            Dim pEdgeIDFilter As IQueryFilter
            Dim pQFtlines As IQueryFilter, pQF As IQueryFilter, pQF2 As IQueryFilter
            Dim pFeature As IFeature, pFeatTRoute As IFeature, pFeatE As IFeature
            Dim pFeatCursor As IFeatureCursor
            Dim pFCtroute As IFeatureCursor
            Dim pTC As ICursor
            Dim pRow As IRow
            Dim i As Long

            pQFtlines = New QueryFilter
            pFCtroute = pTLineFCls.Search(Nothing, False)
            pFeatTRoute = pFCtroute.NextFeature

            Dim idIndex As Long, sIndex As Long
            'sort transit segment table by LineID and SegOrder, then loop through the sorted rows
            Dim pSort As ITableSort
            Dim dctTransitLine As New Dictionary(Of Long, Feature)

            pSort = New TableSort
            pQF = New QueryFilter

            If strLayerPrefix = "SDE" Then
                pQF.WhereClause = "SUBSTRING(CAST(LineID AS CHARACTER), 2, 2) = '" & Right(CStr(inserviceyear), 2) & "'" & "Or SUBSTRING(CAST(LineID AS CHARACTER), 2, 2) = '" & Right(CStr(inserviceyear + 1), 2) & "'"
            Else
                pQF.WhereClause = "Mid([LineID], 2, 2) = '" & Right(CStr(inserviceyear), 2) & "'"
            End If

            Dim count As Long
            Dim sLineID As String

            count = pTLineFCls.FeatureCount(Nothing)
            idIndex = pTLineFCls.FindField("LineID")
            sIndex = tblTSeg.FindField("SegOrder")
            Dim x As Long
            Dim y As Long
            

            Dim pTransitMode As String
            Dim intOperator As Integer

            'Dim intTPCounter As Integer
            Do Until pFeatTRoute Is Nothing
                'loop through each line and add the feature to a dictionar

                sLineID = pFeatTRoute.Value(idIndex)
                If Not dctTransitLine.ContainsKey(sLineID) Then dctTransitLine.Add(sLineID, pFeatTRoute)
                pFeatTRoute = pFCtroute.NextFeature

            Loop

            'now get starting segment
            '    Dim skip As Integer, cindex As Integer
            Dim preNode As Long, preLine As Long, curNode As Long, nextNode As Long, preWNode As Long, curWNode As Long, nextWNode As Long, preWNode2 As Long
            Dim lSegOrder As Long, lLineID As Long, LineInfo As String, lUseGP As Long, lTOD As Long
            Dim fldSegOrder As Long, fldLineId As Long, fldINode As Long, fldJNode As Long, fldUseGP As Long
            Dim fldPath As Long, fldDwtStop As Long, fldtimeFuncID As Long, fldLayover As Long, fldUser1 As Long, fldUser2 As Long, fldUser3 As Long
            Dim sPath As String, sDwtStop As String, stimeFuncID As String, sLayover As String, sUser1 As String, sUser2 As String, sUser3 As String
            Dim PreWeaveString As String, lastNodeString As String
            Dim bPreWeave As Boolean
            Dim sDwtStopJ As String
            Dim tempLastLineString As String
            With tblTSeg
                fldINode = .FindField(g_INode)
                fldJNode = .FindField(g_JNode)
                fldSegOrder = .FindField("SegOrder")
                fldLineId = .FindField("LineID")
                fldPath = .FindField("Path")
                fldDwtStop = .FindField("DwtStop")
                fldtimeFuncID = .FindField("timeFuncID")
                fldLayover = .FindField("LayOver")
                fldUser1 = .FindField("User1")
                fldUser2 = .FindField("User2")
                fldUser3 = .FindField("User3")
                fldUseGP = .FindField("UseGPOnly")
            End With


            'get all nodes in the intermediate layer
            Dim dctNodes As New Dictionary(Of String, String)

            'dctNodes = New Dictionary
            getAllNodes2(dctNodes)
            ' pStatusBar.ShowProgressBar("Creating Transit...", 0, tblTSeg.rowcount(pQF), 1, True)
            Dim intTPCounter As Long
            Dim dctTransitPoints As Dictionary(Of Long, Long)
            Dim dctDwellTimes As Dictionary(Of Long, String)
            Dim dctStopDistance As Dictionary(Of Long, Long)
            Dim curTransitLine As clsTransitLineAtts



            'Dim x As Long
            

            For Each kvp As KeyValuePair(Of Long, ESRI.ArcGIS.Geodatabase.Feature) In dctTransitLine
                

                Dim lineID As Long = kvp.Key
                Dim lineFeature As ESRI.ArcGIS.Geodatabase.Feature = kvp.Value
                curTransitLine = New clsTransitLineAtts(lineFeature)
                pTransitMode = curTransitLine.Mode
                intOperator = curTransitLine.Company

                
               
                'get an array with tod headways
                Dim todList As New List(Of String)
                If curTransitLine.Headway_AM > 0 Then todList.Add("AM")
                If curTransitLine.Headway_MD > 0 Then todList.Add("MD")
                If curTransitLine.Headway_PM > 0 Then todList.Add("PM")
                If curTransitLine.Headway_EV > 0 Then todList.Add("EV")
                If curTransitLine.Headway_NI > 0 Then todList.Add("NI")

                'if ltod is All them we only need to build it once, if it is am then we need to build them for each time period seperateley
                lTOD = curTransitLine.TimePeriod

                For Each tod In todList
                    Dim lineQF As IQueryFilter
                    lineQF = New QueryFilter
                    lineQF.WhereClause = "LineID = " & Str(lineID)
                    With pSort
                        .Fields = "LineID, SegOrder"
                        .Ascending("LineID") = True
                        .Ascending("SegOrder") = True
                        .QueryFilter = lineQF
                        .Table = tblTSeg
                    End With
                    pSort.Sort(Nothing)
                    pTC = pSort.Rows
                    pRow = pTC.NextRow
                    x = 0
                    Dim intNodeCounter As Long = 1
                    'preLine = 0

                    'preLine = lineID
                    Do Until pRow Is Nothing
                        'sec 073009
                        'get all the TransitPoints that belong to the route, store them in a dictionary
                        'store their DWTs in another dictionary
                        If x <> pRow.Value(fldLineId) Then
                            dctTransitPoints = New Dictionary(Of Long, Long)
                            dctDwellTimes = New Dictionary(Of Long, String)
                            dctStopDistance = New Dictionary(Of Long, Long)
                            GetTransitPointsByLineID2(dctTransitPoints, dctDwellTimes, Pfltransitpoints.FeatureClass, lineID)
                            intTPCounter = 1
                            x = pRow.Value(fldLineId)
                            GetStopDistances2(lineID, dctStopDistance, tblTSeg, dctTransitPoints)
                        End If


                        lSegOrder = pRow.Value(fldSegOrder)

                        '[040207]hyu: per Jeff's email on [05/16/06]: path=no or path=yes where TransitLines.Path=0 signifies no and TransitLines.Path=1 signifies yes.
                        sPath = IIf(pRow.Value(fldPath) = 1, " path=yes", IIf(pRow.Value(fldPath) = 0, " path=no", ""))

                        'sDwtStop = " dwt=" + CStr(pRow.value(fldDwtStop))
                        'check to see if the Inode of the segment is a transit point for the transit route:


                        Dim strDwell As String
                        Dim intPrevTN As Integer
                        Dim lngStopDistance As Long

                        Dim prevNodeString As String



                        'see if the current node is a transit point
                        Try


                            If dctTransitPoints.Item(intTPCounter) = pRow.Value(fldINode) Then
                                If dctStopDistance.ContainsKey(intTPCounter + 1) Then


                                    lngStopDistance = dctStopDistance.Item(intTPCounter) - dctStopDistance.Item(intTPCounter + 1)
                                    strDwell = dctDwellTimes.Item(intTPCounter)
                                    intPrevTN = intTPCounter
                                    intTPCounter = intTPCounter + 1
                                Else
                                    WriteLogLine("Line Failure at: " & lineID & " " & intTPCounter + 1)
                                    intTPCounter = intTPCounter + 1
                                End If
                            End If

                            If dctTransitPoints.ContainsKey(intPrevTN + 1) Then
                                If dctTransitPoints.Item(intPrevTN + 1) = pRow.Value(fldJNode) Then
                                    sDwtStop = " " & strDwell
                                Else
                                    sDwtStop = " dwt=#.00"
                                End If
                            Else
                                WriteLogLine("Line Failure at: " & lineID & " " & intTPCounter + 1)
                            End If


                            'if so, grab the DWT from the DWT dictionary
                        Catch ex As Exception
                            MessageBox.Show(ex.ToString)

                        End Try



                        'deal with the last node
                        If dctTransitPoints.Item(dctTransitPoints.Count) = pRow.Value(fldJNode) Then
                            'sDwtStopJ = " " & CStr(GetDwellTime2(lineID, dctTransitPoints.count))
                            sDwtStopJ = " " & dctDwellTimes.Item(dctDwellTimes.Count)
                        End If

                        'sDwtStopJ = " " & CStr(GetDwellTime(Pfltransitpoints.FeatureClass, lineID, pRow.value(fldJNode)))
                        stimeFuncID = " ttf=" + CStr(pRow.Value(fldtimeFuncID))
                        sLayover = IIf(pRow.Value(fldLayover) > 0, " lay=" + CStr(pRow.Value(fldLayover)), "")
                        sUser1 = " us1=" + CStr(IIf(IsDBNull(pRow.Value(fldUser1)), 0, "0"))
                        sUser2 = " us2=" + CStr(IIf(IsDBNull(pRow.Value(fldUser2)), 0, pRow.Value(fldUser2)))
                        sUser3 = " us3=" + CStr(IIf(IsDBNull(pRow.Value(fldUser3)), 0, pRow.Value(fldUser3)))
                        tempString = sDwtStop + stimeFuncID + sLayover + sUser1 + sUser2 + sUser3
                        tempLastLineString = sDwtStopJ + stimeFuncID + sLayover + sUser1 + sUser2 + sUser3


                        If lSegOrder = 1 Then

                            'start a new line, then get the transit line info

                            i = 0

                            'get the first node
                            curNode = pRow.Value(fldINode)
                            nextNode = pRow.Value(fldJNode)
                            preNode = 0 'curNode
                            preWNode = 0
                            curWNode = 0
                            PreWeaveString = ""
                            lastNodeString = ""
                            bPreWeave = False

                            'new transit line, write out line info
                            writeTransitLineInfo(lTOD, curTransitLine, tod)

                            If sPath <> "" Then writeTransitNode2(lTOD, sPath, curTransitLine, tod)
                        End If

                        curNode = pRow.Value(fldINode)
                        nextNode = pRow.Value(fldJNode)

                        If curNode <> preNode Then
                            If Not dctNodes.ContainsKey(CStr(curNode)) Then
                                If fVerboseLog Then WriteLogLine("Line " & lineID & " Node " & curNode & " SegOrder=" & lSegOrder & " dissolved")

                            Else    'dctNodes.Exists(CStr(curNode))
                                'form an edge of preNode-curNode
                                'check whether it should use the GP/TR/HOV lane
                                'if the nodes dosn't exist, skip it.
                                lUseGP = pRow.Value(fldUseGP)

                                If lUseGP = 0 Then
                                    If fVerboseLog Then WriteLogLine("Line " & lineID & " Node " & curNode & " SegOrder=" & lSegOrder & " UseGPOnly=0")

                                    pQF2 = New QueryFilter
                                    pQF2.WhereClause = "(" + g_INode + "=" + CStr(curNode) + " AND " + g_JNode + "=" + CStr(nextNode) + ") OR (" _
                                        + g_JNode + "=" + CStr(curNode) + " AND " + g_INode + "=" + CStr(nextNode) + ")"
                                    pFeatCursor = m_edgeShp.Search(pQF2, False)
                                    pFeature = pFeatCursor.NextFeature
                                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(pFeatCursor)

                                    If Not pFeature Is Nothing Then
                                        'Do TTF stuff here
                                        
                                            'pTransitMode = pFeature.Fields.FindField("NewFacilityType")
                                            lngStopDistance = Math.Abs(lngStopDistance)
                                            Dim strSegmentMode As String
                                            strSegmentMode = pFeature.Value(pFeature.Fields.FindField("Modes"))
                                        If pTransitMode = "r" Or pTransitMode = "c" Or pTransitMode = "f" Then
                                            stimeFuncID = 5
                                        ElseIf strSegmentMode = "abp" Or strSegmentMode = "abwlp" Or strSegmentMode = "abrp" Or strSegmentMode = "abwp" Then
                                            stimeFuncID = 4
                                        ElseIf lngStopDistance > 7920 Then
                                            stimeFuncID = 14
                                        ElseIf (pFeature.Value(pFeature.Fields.FindField("NewFacilityType")) = 1 Or pFeature.Value(pFeature.Fields.FindField("NewFacilityType")) = 2) And lngStopDistance > 2640 Then
                                            stimeFuncID = 13
                                        ElseIf lngStopDistance > 2640 Then
                                            stimeFuncID = 12
                                        Else : stimeFuncID = 11

                                        End If
                                            stimeFuncID = " ttf=" + stimeFuncID



                                        tempString = sDwtStop + stimeFuncID + sLayover + sUser1 + sUser2 + sUser3
                                        tempLastLineString = sDwtStopJ + stimeFuncID + sLayover + sUser1 + sUser2 + sUser3


                                        'Does the Transit Segment goe in the IJ or JI direciton
                                        'Does the Transit Segment goe in the IJ or JI direciton
                                        If pFeature.Value(m_edgeShp.FindField("INode")) = curNode Then 'IJ

                                            If Not IsDBNull(pFeature.Value(m_edgeShp.FindField("TR_I"))) Then

                                                'If pFeature.Value(m_edgeShp.FindField("TR_I")) > 0 Then
                                                'weave nodes from "TR_I" and "TR_J"
                                                curWNode = pFeature.Value(m_edgeShp.FindField("TR_I")) + m_Offset
                                                nextWNode = pFeature.Value(m_edgeShp.FindField("TR_J")) + m_Offset

                                                If fVerboseLog Then WriteLogLine("potential weave nodes (TR_J/I) " & preWNode & ", " & curWNode)
                                                'End If
                                            Else
                                                'weave nodes may from "HOV_I/J"
                                                Select Case pFeature.Value(m_edgeShp.FindField("FacilityType"))
                                                    Case 11, 12, 3, 13, 9
                                                        curWNode = 0
                                                        nextWNode = 0
                                                    Case Else
                                                        pEdgeID = pFeature.Value(m_edgeShp.FindField("PSRCEdgeID"))

                                                        Select Case tod
                                                            Case Is = "AM"
                                                                If curTransitLine.Headway_AM > 0 Then
                                                                    If dctAM_IJ_HOV.ContainsKey(pEdgeID) Then

                                                                        curWNode = IIf(IsDBNull(pFeature.Value(m_edgeShp.FindField("HOV_I"))), 0, pFeature.Value(m_edgeShp.FindField("HOV_I")) + m_Offset)
                                                                        nextWNode = IIf(IsDBNull(pFeature.Value(m_edgeShp.FindField("HOV_J"))), 0, pFeature.Value(m_edgeShp.FindField("HOV_J")) + m_Offset)
                                                                    End If
                                                                End If
                                                            Case Is = "MD"
                                                                If curTransitLine.Headway_MD > 0 Then
                                                                    If dctMD_IJ_HOV.ContainsKey(pEdgeID) Then
                                                                        curWNode = IIf(IsDBNull(pFeature.Value(m_edgeShp.FindField("HOV_I"))), 0, pFeature.Value(m_edgeShp.FindField("HOV_I")) + m_Offset)
                                                                        nextWNode = IIf(IsDBNull(pFeature.Value(m_edgeShp.FindField("HOV_J"))), 0, pFeature.Value(m_edgeShp.FindField("HOV_J")) + m_Offset)
                                                                    End If
                                                                End If
                                                            Case Is = "PM"
                                                                If curTransitLine.Headway_PM > 0 Then
                                                                    If dctPM_IJ_HOV.ContainsKey(pEdgeID) Then
                                                                        curWNode = IIf(IsDBNull(pFeature.Value(m_edgeShp.FindField("HOV_I"))), 0, pFeature.Value(m_edgeShp.FindField("HOV_I")) + m_Offset)
                                                                        nextWNode = IIf(IsDBNull(pFeature.Value(m_edgeShp.FindField("HOV_J"))), 0, pFeature.Value(m_edgeShp.FindField("HOV_J")) + m_Offset)
                                                                    End If
                                                                End If
                                                            Case Is = "EV"
                                                                If curTransitLine.Headway_EV > 0 Then
                                                                    If dctEV_IJ_HOV.ContainsKey(pEdgeID) Then
                                                                        curWNode = IIf(IsDBNull(pFeature.Value(m_edgeShp.FindField("HOV_I"))), 0, pFeature.Value(m_edgeShp.FindField("HOV_I")) + m_Offset)
                                                                        nextWNode = IIf(IsDBNull(pFeature.Value(m_edgeShp.FindField("HOV_J"))), 0, pFeature.Value(m_edgeShp.FindField("HOV_J")) + m_Offset)
                                                                    End If
                                                                End If
                                                            Case Is = "NI"
                                                                If curTransitLine.Headway_NI > 0 Then
                                                                    If dctNI_IJ_HOV.ContainsKey(pEdgeID) Then
                                                                        curWNode = IIf(IsDBNull(pFeature.Value(m_edgeShp.FindField("HOV_I"))), 0, pFeature.Value(m_edgeShp.FindField("HOV_I")) + m_Offset)
                                                                        nextWNode = IIf(IsDBNull(pFeature.Value(m_edgeShp.FindField("HOV_J"))), 0, pFeature.Value(m_edgeShp.FindField("HOV_J")) + m_Offset)
                                                                    End If
                                                                End If
                                                        End Select
                                                End Select

                                                If fVerboseLog Then WriteLogLine("potential weave nodes (HOV_J/I) " & curWNode & ", " & nextWNode)
                                            End If

                                        Else 'Transit segment goes in the JI direction


                                            If Not IsDBNull(pFeature.Value(m_edgeShp.FindField("TR_J"))) Then
                                                'If pFeature.Value(m_edgeShp.FindField("TR_J")) > 0 Then
                                                'weave nodes from "TR_I" and "TR_J"
                                                curWNode = pFeature.Value(m_edgeShp.FindField("TR_J")) + m_Offset
                                                nextWNode = pFeature.Value(m_edgeShp.FindField("TR_I")) + m_Offset

                                                'End If
                                                'End If
                                                If fVerboseLog Then WriteLogLine("potential weave nodes (TR_I/J) " & curWNode & ", " & nextWNode)

                                            Else
                                                'weave nodes may from "HOV_I/J"
                                                Select Case pFeature.Value(m_edgeShp.FindField("FacilityType"))
                                                    Case 11, 12, 3, 13, 9
                                                        curWNode = 0
                                                        nextWNode = 0
                                                    Case Else
                                                        pEdgeID = pFeature.Value(m_edgeShp.FindField("PSRCEdgeID"))

                                                        Select Case tod
                                                            Case Is = "AM"
                                                                If curTransitLine.Headway_AM > 0 Then
                                                                    If dctAM_JI_HOV.ContainsKey(pEdgeID) Then
                                                                        ' clsModeAtts.modeAttributeRow = pMARow
                                                                        'MsgBox (clsModeAtts.IJLANESHOVAM)
                                                                        curWNode = IIf(IsDBNull(pFeature.Value(m_edgeShp.FindField("HOV_J"))), 0, pFeature.Value(m_edgeShp.FindField("HOV_J")) + m_Offset)
                                                                        nextWNode = IIf(IsDBNull(pFeature.Value(m_edgeShp.FindField("HOV_I"))), 0, pFeature.Value(m_edgeShp.FindField("HOV_I")) + m_Offset)
                                                                    End If
                                                                End If
                                                            Case Is = "MD"
                                                                If curTransitLine.Headway_MD > 0 Then
                                                                    If dctMD_JI_HOV.ContainsKey(pEdgeID) Then
                                                                        curWNode = IIf(IsDBNull(pFeature.Value(m_edgeShp.FindField("HOV_J"))), 0, pFeature.Value(m_edgeShp.FindField("HOV_J")) + m_Offset)
                                                                        nextWNode = IIf(IsDBNull(pFeature.Value(m_edgeShp.FindField("HOV_I"))), 0, pFeature.Value(m_edgeShp.FindField("HOV_I")) + m_Offset)
                                                                    End If
                                                                End If
                                                            Case Is = "PM"
                                                                If curTransitLine.Headway_PM > 0 Then
                                                                    If dctPM_JI_HOV.ContainsKey(pEdgeID) Then
                                                                        curWNode = IIf(IsDBNull(pFeature.Value(m_edgeShp.FindField("HOV_J"))), 0, pFeature.Value(m_edgeShp.FindField("HOV_J")) + m_Offset)
                                                                        nextWNode = IIf(IsDBNull(pFeature.Value(m_edgeShp.FindField("HOV_I"))), 0, pFeature.Value(m_edgeShp.FindField("HOV_I")) + m_Offset)
                                                                    End If
                                                                End If
                                                            Case Is = "EV"
                                                                If curTransitLine.Headway_EV > 0 Then
                                                                    If dctEV_JI_HOV.ContainsKey(pEdgeID) Then
                                                                        curWNode = IIf(IsDBNull(pFeature.Value(m_edgeShp.FindField("HOV_J"))), 0, pFeature.Value(m_edgeShp.FindField("HOV_J")) + m_Offset)
                                                                        nextWNode = IIf(IsDBNull(pFeature.Value(m_edgeShp.FindField("HOV_I"))), 0, pFeature.Value(m_edgeShp.FindField("HOV_I")) + m_Offset)
                                                                    End If
                                                                End If
                                                            Case Is = "NI"
                                                                If curTransitLine.Headway_NI > 0 Then
                                                                    If dctNI_JI_HOV.ContainsKey(pEdgeID) Then
                                                                        curWNode = IIf(IsDBNull(pFeature.Value(m_edgeShp.FindField("HOV_J"))), 0, pFeature.Value(m_edgeShp.FindField("HOV_J")) + m_Offset)
                                                                        nextWNode = IIf(IsDBNull(pFeature.Value(m_edgeShp.FindField("HOV_I"))), 0, pFeature.Value(m_edgeShp.FindField("HOV_I")) + m_Offset)
                                                                    End If
                                                                End If
                                                        End Select
                                                End Select

                                                If fVerboseLog Then WriteLogLine("potential weave nodes (HOV_I/J) " & preWNode & ", " & curWNode)
                                            End If
                                        End If 'pfeature.value(m_edgeShp.FindField("INode")) = curNode



                                        If bPreWeave And preWNode = curWNode And curWNode <> 0 Then
                                            'we are on hov system, write out the next node because current one has alreayd been written. We
                                            'are one step ahead here. 
                                            'Use next weave node, but have to make sure dwell time is right!
                                            'if next node is a stop we need to assign it to the previous node since we are one node ahead here!

                                            'writeTransitNode(lTOD, " " + CStr(nextWNode) + tempString)

                                            If dctTransitPoints.Item(intPrevTN + 1) = pRow.Value(fldJNode) Then
                                                'sDwtStop = " " & strDwell
                                                'sDwtStop is set from above and should indicate a stop. We need to apply this to the previous node. 


                                                prevNodeString = dctNodeString.Item(intNodeCounter - 1)
                                                'prevNodeString = prevNodeString.Replace(LTrim(prevDwtStop), LTrim(sDwtStop))
                                                prevNodeString = prevNodeString.Replace(LTrim("dwt=#.00"), LTrim(sDwtStop))
                                                dctNodeString.Item(intNodeCounter - 1) = prevNodeString


                                                'Stop should not take place here. The dictionary will be updated next iteration if it is indeed a stop. 
                                                tempString = " dwt=#.00" + stimeFuncID + sLayover + sUser1 + sUser2 + sUser3

                                            End If




                                            dctNodeString.Add(intNodeCounter, " " + CStr(nextWNode) + tempString)
                                            intNodeCounter = intNodeCounter + 1
                                            bPreWeave = True
                                        Else
                                            'On GP lanes, did we just get off HOV?
                                            'If bPreWeave And pRow.Value(fldINode) = dctTransitPoints(intPrevTN) Then
                                            'if yes, have to update last node with proper stop info:
                                            'NodeString = dctNodeString.Item(intNodeCounter - 1)
                                            'prevNodeString = System.Text.RegularExpressions.Regex.Replace(prevNodeString, "dwt=*....", LTrim(dctDwellTimes.Item(intTPCounter - 1)))

                                            'see if transit line is entering HOV/Managed lanes, if so, write out I & J Nodes of HOV:
                                            If curWNode > 0 And nextWNode > 0 Then
                                                'this is the gp node. If the next node is a stop, need to put the dwll time on the INode of the HOV, not here. 
                                                tempString = " dwt=#.00" + stimeFuncID + sLayover + sUser1 + sUser2 + sUser3
                                                'writeTransitNode(lTOD, " " + CStr(dctNodes.Item(CStr(curNode))) + tempString)
                                                dctNodeString.Add(intNodeCounter, " " + CStr(dctNodes.Item(CStr(curNode))) + tempString)
                                                intNodeCounter = intNodeCounter + 1

                                                'HOV INode of the first HOV link. Put dwell time here if the next node is a stop. 
                                                tempString = sDwtStop + stimeFuncID + sLayover + sUser1 + sUser2 + sUser3
                                                'writeTransitNode(lTOD, " " + CStr(curWNode) + tempString)
                                                dctNodeString.Add(intNodeCounter, " " + CStr(curWNode) + tempString)
                                                intNodeCounter = intNodeCounter + 1

                                                'HOV JNode of the first HOV link
                                                'if next node on the next link is indeed a stop, this dictionary will get updated. 
                                                tempString = " dwt=#.00" + stimeFuncID + sLayover + sUser1 + sUser2 + sUser3
                                                'writeTransitNode(lTOD, " " + CStr(nextWNode) + tempString)
                                                dctNodeString.Add(intNodeCounter, " " + CStr(nextWNode) + tempString)
                                                intNodeCounter = intNodeCounter + 1
                                                bPreWeave = True
                                            Else
                                                bPreWeave = False
                                                'was on a GP and still is 
                                                'writeTransitNode(lTOD, " " + CStr(dctNodes.Item(CStr(curNode))) + tempString)
                                                dctNodeString.Add(intNodeCounter, " " + CStr(dctNodes.Item(CStr(curNode))) + tempString)
                                                intNodeCounter = intNodeCounter + 1
                                            End If
                                        End If
                                        'sec 072909- fixing DWTs
                                        'lastNodeString = " " + CStr(dctNodes.Item(CStr(nextNode))) + tempString
                                        lastNodeString = " " + CStr(dctNodes.Item(CStr(nextNode))) + tempLastLineString

                                        preNode = curNode
                                        preWNode = nextWNode
                                        curWNode = 0
                                        nextWNode = 0

                                    Else 'pfeature Is Nothing
                                        If i = 0 Then
                                            WriteLogLine("Data Error: Segment " + CStr(lSegOrder) + " on transit line " + CStr(lineID) + " underlying TransRefEdge not in service")
                                            WriteLogLine("Skip the rest of line " & lineID)
                                            i = 1
                                        End If
                                    End If  'Not pfeature Is Nothing
                                Else    'lUseGP <> 0
                                    'GP only

                                    'SEC 072909- fixing DWTs
                                    'lastNodeString = " " + CStr(dctNodes.Item(CStr(nextNode))) + tempString
                                    lastNodeString = " " + CStr(dctNodes.Item(CStr(nextNode))) + tempLastLineString
                                    'writeTransitNode(lTOD, " " + CStr(dctNodes.Item(CStr(curNode))) + tempString)
                                    dctNodeString.Add(intNodeCounter, " " + CStr(dctNodes.Item(CStr(curNode))) + tempString)
                                    intNodeCounter = intNodeCounter + 1
                                    bPreWeave = False
                                    'in case we need to update the previous string

                                    preNode = curNode
                                    curNode = nextNode


                                    nextNode = 0
                                    curWNode = 0
                                    preWNode = 0
                                    nextWNode = 0
                            End If  'lUseGP = 0
                            End If  'Not dctNodes.Exists(CStr(curNode))
                        End If  'curNode <> preNode
                        pRow = pTC.NextRow

                        If pRow Is Nothing Then


                            'now is at the end, write the last node
                            'writeTransitNode(lTOD, lastNodeString)
                            dctNodeString.Add(intNodeCounter, lastNodeString)

                            For y = 1 To dctNodeString.Count
                                writeTransitNode2(lTOD, dctNodeString.Item(y), curTransitLine, tod)
                            Next y
                            dctNodeString.Clear()
                            intNodeCounter = 1
                            'Node Sequence for all TODs is written out, we can exit the for loop:
                            If lTOD = 0 Then
                                Exit For
                            End If

                        End If

                    Loop
                Next
            Next




            'pStatusBar.HideProgressBar()
            FileClose(1, 2, 3, 4, 5)
            pFeatLayerE = Nothing
            '   pTLineLayer = Nothing
            pTLineFCls = Nothing
            pFeature = Nothing
            WriteLogLine("FINISHED create_TransitFile at " & Now())
            dctTransitLine.Clear()
            dctNodes.Clear()
            dctTransitLine = Nothing
            dctNodes = Nothing




            'Exit Sub
            'eh:



        Catch ex As Exception
            MessageBox.Show(ex.ToString)

            CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.create_TransitFile")
            MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.create_TransitFile")
            'Print #9, Err.Description, vbExclamation, "TransitFile"
            FileClose(1, 2, 3, 4, 5)
            'Close 9

        End Try

    End Sub
    Public Function GetTransitPointsByLineID2(ByVal dct As Dictionary(Of Long, Long), ByVal dct2 As Dictionary(Of Long, String), ByVal TransitPointFC As IFeatureClass, ByVal LineID As Long)
        '[080309] SEC: stores all the transit points in a dictionary with a counter as the key. Could not use pointorder as
        'key because there are sequential transit nodes that are on the same node/junction. This was causing a problem and
        'are not included int the dictionary.

        Dim pFCursor As IFeatureCursor
        Dim pFilter As IQueryFilter
        Dim pFeature As IFeature
        'Dim pDictionary As Dictionary
        ' pDictionary = New Dictionary(of Long, Long)
        Dim indexOrder As Long
        Dim indexJunctID As Long
        Dim lngPrevJunctID As Long
        Dim indexDwell As Long
        indexDwell = TransitPointFC.FindField("txtDWT")
        indexJunctID = TransitPointFC.FindField("PSRCJunctID")
        indexOrder = TransitPointFC.FindField("PointOrder")
        'indexOrder = TransitPointFC.FindField("PointOrder")
        Dim Counter As Long
        Dim psort As ITableSort
        psort = New TableSort
        Counter = 1
        pFilter = New QueryFilter
        pFilter.WhereClause = "LineID = " & LineID
        With psort
            .Fields = "LineID, PointOrder"
            .Ascending("LineID") = True
            .Ascending("PointOrder") = True
            .QueryFilter = pFilter ' pQF
            .Table = TransitPointFC
        End With
        psort.Sort(Nothing)
        pFCursor = psort.Rows
        Dim prevDwell As String



        ' pFCursor = TransitPointFC.Search(pFilter, False)
        pFeature = pFCursor.NextFeature
        Do Until pFeature Is Nothing
            If lngPrevJunctID = pFeature.Value(indexJunctID) Then
                'this node is on top of the last node, dont use it
                If prevDwell <> pFeature.Value(indexDwell) And pFeature.Value(indexDwell) = "dwt=.25" Then
                    dct2.Remove(Counter - 1)
                    dct2.Add((Counter - 1), pFeature.Value(indexDwell))
                    pFeature = pFCursor.NextFeature
                Else
                    pFeature = pFCursor.NextFeature
                End If
            Else
                dct.Add(CType(Counter, Long), pFeature.Value(CType(indexJunctID, Long)))
                dct2.Add(CType(Counter, Long), pFeature.Value(indexDwell))
                prevDwell = pFeature.Value(indexDwell)
                lngPrevJunctID = pFeature.Value(indexJunctID)
                Counter = Counter + 1
                pFeature = pFCursor.NextFeature
            End If

        Loop



        ' GetTransitPointsByLineID = pDictionary


    End Function
    Public Function getAllNodes2(ByVal dct As Dictionary(Of String, String))
        Dim pCs As ICursor
        Dim pRow As IRow
        Dim fldEID As Long, fldPsrcID As Long
        Dim lPsrcID As Long, lEID As Long

        'get nodes from ScenarioJunction layer
        pCs = m_junctShp.Search(Nothing, True)
        pRow = pCs.NextRow
        fldEID = m_junctShp.FindField("Scen_Node")
        fldPsrcID = m_junctShp.FindField(g_PSRCJctID)

        Do Until pRow Is Nothing
            lPsrcID = pRow.Value(fldPsrcID)
            lEID = pRow.Value(fldEID)
            dct.Add(CStr(lPsrcID), lEID)
            pRow = pCs.NextRow
        Loop

        'get weave nodes from ScenarioEdge layer
        Dim pFilt As IQueryFilter
        Dim fldTRI As Long, fldTRJ As Long, fldHOVI As Long, fldHOVJ As Long, fldTKI As Long, fldTKJ As Long
        Dim lTRI As Long, lTRJ As Long, lHOVI As Long, lHOVJ As Long, lTKI As Long, lTKJ As Long

        With m_edgeShp
            fldTRI = .FindField("TR_I")
            fldTRJ = .FindField("TR_J")
            fldHOVI = .FindField("HOV_I")
            fldHOVJ = .FindField("HOV_J")
            fldTKI = .FindField("TK_I")
            fldTKJ = .FindField("TK_J")
        End With

        pFilt = New QueryFilter
        pFilt.WhereClause = "TR_I + TR_J + HOV_I + HOV_J + TK_I + TK_J >0 "

        pCs = m_edgeShp.Search(pFilt, True)
        pRow = pCs.NextRow
        Do Until pRow Is Nothing
            lTRI = pRow.Value(fldTRI)
            lTRJ = pRow.Value(fldTRJ)
            lHOVI = pRow.Value(fldHOVI)
            lHOVJ = pRow.Value(fldHOVJ)
            lTKI = pRow.Value(fldTKI)
            lTKJ = pRow.Value(fldTKJ)

            If lTRI > 0 And Not dct.ContainsKey(CStr(lTRI)) Then dct.Add(CStr(lTRI), lTRI)
            If lTRJ > 0 And Not dct.ContainsKey(CStr(lTRJ)) Then dct.Add(CStr(lTRJ), lTRJ)
            If lHOVI > 0 And Not dct.ContainsKey(CStr(lHOVI)) Then dct.Add(CStr(lHOVI), lHOVI)
            If lHOVJ > 0 And Not dct.ContainsKey(CStr(lHOVJ)) Then dct.Add(CStr(lHOVJ), lHOVJ)
            If lTKI > 0 And Not dct.ContainsKey(CStr(lTKI)) Then dct.Add(CStr(lTKI), lTKI)
            If lTKJ > 0 And Not dct.ContainsKey(CStr(lTKJ)) Then dct.Add(CStr(lTKJ), lTKJ)

            pRow = pCs.NextRow
        Loop

        pCs = Nothing
        pFilt = Nothing

    End Function
    Public Function GetStopDistances2(ByVal LineID As Long, ByVal dct As Dictionary(Of Long, Long), ByVal TransitSegments As ITable, ByVal dctTP As Dictionary(Of Long, Long))
        Dim pCursor As ICursor
        Dim pFilter As IQueryFilter
        Dim pRow As IRow
        Dim indexDist As Long
        Dim indexINode As Long
        Dim indexJNode As Long
        indexDist = TransitSegments.FindField("Change")
        indexINode = TransitSegments.FindField("INode")
        indexJNode = TransitSegments.FindField("JNode")
        Dim lngDist As Long
        Dim INode As Long
        Dim JNode As Long
        Dim pSort As ITableSort
        pSort = New TableSort
        Dim intTPCounter As Long
        Dim x As Long
        x = 1

        pFilter = New QueryFilter
        pFilter.WhereClause = "LineID = " & LineID
        With pSort
            .Fields = "SegOrder"
            '.Ascending("LineID") = True
            .Ascending("SegOrder") = True
            .QueryFilter = pFilter ' pQF
            .Table = TransitSegments
        End With
        pSort.Sort(Nothing)

        pCursor = pSort.Rows

        intTPCounter = 1
        Dim test As Boolean
        test = False
        pRow = pCursor.NextRow
        lngDist = pRow.Value(indexDist)
        Do Until pRow Is Nothing
            INode = pRow.Value(indexINode)
            JNode = pRow.Value(indexJNode)

            'Check to see if we are at the last transit point Transit Point
            'if so, we need to do some different accounting
            'have to include the length of the last link
            'If intTPCounter = dctTP.count - 1 Then
            'If INode = dctTP.Item(intTPCounter) Then
            'dct.Add intTPCounter, lngDist
            'lngDist = pRow.value(indexDist)

            'ElseIf JNode = dctTP.Item(intTPCounter + 1) Then
            'lngDist = lngDist + pRow.value(indexDist)
            'dct.Add intTPCounter, lngDist
            'Else
            'lngDist = lngDist + pRow.value(indexDist)
            'End If

            Try



                If intTPCounter = 1 Then
                    dct.Add(CType(intTPCounter, Long), 0)
                    intTPCounter = intTPCounter + 1
                    If JNode = dctTP.Item(CType(intTPCounter, Long)) Then
                        lngDist = pRow.Value(indexDist)
                        dct.Add(CType(intTPCounter, Long), CType(lngDist, Long))
                        intTPCounter = intTPCounter + 1
                        'test = True
                    Else
                        lngDist = pRow.Value(indexDist)
                    End If

                ElseIf dctTP.ContainsKey(CType(intTPCounter, Long)) Then
                    If JNode = dctTP.Item(CType(intTPCounter, Long)) Then



                        lngDist = lngDist + pRow.Value(indexDist)
                        dct.Add(CType(intTPCounter, Long), CType(lngDist, Long))
                        intTPCounter = intTPCounter + 1


                    Else
                        lngDist = lngDist + pRow.Value(indexDist)


                    End If
                Else
                    WriteLogLine("Transit Error: Line " & LineID & " Stop Number " & intTPCounter)
                End If








                pRow = pCursor.NextRow
            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
        Loop





    End Function
    Public Function LoadHOVDict(ByVal modeAtts_dct As Dictionary(Of Long, Row), ByVal dctAM_IJHOV As Dictionary(Of Long, Long), ByVal dctAM_JIHOV As Dictionary(Of Long, Long), ByVal dctMD_IJHOV As Dictionary(Of Long, Long), ByVal dctMD_JIHOV As Dictionary(Of Long, Long), Optional ByVal dctPM_IJHOV As Dictionary(Of Long, Long) = Nothing, Optional ByVal dctPM_JIHOV As Dictionary(Of Long, Long) = Nothing,
                                Optional ByVal dctEV_IJHOV As Dictionary(Of Long, Long) = Nothing, Optional ByVal dctEV_JIHOV As Dictionary(Of Long, Long) = Nothing, Optional ByVal dctNI_IJHOV As Dictionary(Of Long, Long) = Nothing, Optional ByVal dctNI_JIHOV As Dictionary(Of Long, Long) = Nothing)
        For Each kvp As KeyValuePair(Of Long, Row) In modeAtts_dct
            Dim rowModeAtt As New clsModeAttributes(kvp.Value)
            If rowModeAtt.IJLANESHOVAM > 0 Then
                dctAM_IJHOV.Add(rowModeAtt.PSRCEDGEID, rowModeAtt.PSRCEDGEID)
            End If
            If rowModeAtt.JILANESHOVAM > 0 Then
                dctAM_JIHOV.Add(rowModeAtt.PSRCEDGEID, rowModeAtt.PSRCEDGEID)
            End If
            If rowModeAtt.IJLANESHOVMD > 0 Then
                dctMD_IJHOV.Add(rowModeAtt.PSRCEDGEID, rowModeAtt.PSRCEDGEID)
            End If
            If rowModeAtt.JILANESHOVMD > 0 Then
                dctMD_JIHOV.Add(rowModeAtt.PSRCEDGEID, rowModeAtt.PSRCEDGEID)
            End If

            'PM
            If Not dctPM_IJHOV Is Nothing Then
                If rowModeAtt.IJLANESHOVPM > 0 Then
                    dctPM_IJHOV.Add(rowModeAtt.PSRCEDGEID, rowModeAtt.PSRCEDGEID)
                End If
            End If
            If Not dctPM_JIHOV Is Nothing Then
                If rowModeAtt.JILANESHOVPM > 0 Then
                    dctPM_JIHOV.Add(rowModeAtt.PSRCEDGEID, rowModeAtt.PSRCEDGEID)
                End If
            End If

            'Evening
            If Not dctEV_IJHOV Is Nothing Then
                If rowModeAtt.IJLANESHOVEV > 0 Then
                    dctEV_IJHOV.Add(rowModeAtt.PSRCEDGEID, rowModeAtt.PSRCEDGEID)
                End If
            End If
            If Not dctEV_JIHOV Is Nothing Then
                If rowModeAtt.JILANESHOVEV > 0 Then
                    dctEV_JIHOV.Add(rowModeAtt.PSRCEDGEID, rowModeAtt.PSRCEDGEID)
                End If
            End If

            'Night
            If Not dctNI_IJHOV Is Nothing Then
                If rowModeAtt.IJLANESHOVNI > 0 Then
                    dctNI_IJHOV.Add(rowModeAtt.PSRCEDGEID, rowModeAtt.PSRCEDGEID)
                End If
            End If
            If Not dctNI_JIHOV Is Nothing Then
                If rowModeAtt.JILANESHOVNI > 0 Then
                    dctNI_JIHOV.Add(rowModeAtt.PSRCEDGEID, rowModeAtt.PSRCEDGEID)
                End If
            End If
        Next



    End Function

    Public Function GetTransitLineInfo(ByVal _clsTransitLineAtts As clsTransitLineAtts, ByVal headway_value As Double)
        Dim tempString As String
        Dim sLineID As String
        sLineID = CType(_clsTransitLineAtts.LineID, String)



        If Len(sLineID) > 4 Then
            tempString = "c '" + Left(_clsTransitLineAtts.LineID, 2) + "'" + vbNewLine
            tempString = tempString + "a '" + Mid(sLineID, 3, Len(sLineID) - 2) + "'"
        Else
            tempString = "a '" + sLineID + "'"
        End If
        'tempString = "a '" + CStr(pFeatTRoute.value(idIndex)) + "'"
        tempString = tempString + " " + _clsTransitLineAtts.Mode
        'index = pFeatTRoute.Fields.FindField("VehicleType")
        'Index = pFeatTRoute.Fields.FindField(Left("VehicleType", 10))

        tempString = tempString + " " + CStr(_clsTransitLineAtts.VehicleType)

        tempString = tempString + " " + CStr(headway_value)

        tempString = tempString + " " + CStr(_clsTransitLineAtts.Speed)

        tempString = tempString + " '" + CStr(_clsTransitLineAtts.Description) + "'"

        tempString = tempString + " " + CStr(_clsTransitLineAtts.Processing)

        tempString = tempString + " " + CStr(_clsTransitLineAtts.UL2)

        tempString = tempString + " " + CStr(_clsTransitLineAtts.Company)

        'tempString = CStr(_clsTransitLineAtts.TimePeriod) + tempString

        'index = pFeatTRoute.Fields.FindField("Geography")
        'If Not IsNull(pFeatTRoute.value(index)) Then
        'tempString = tempString + " " + CStr(pFeatTRoute.value(index))
        'Else
        'tempString = tempString + " 1"
        'End If
        'index = pFeatTRoute.Fields.FindField("Oneway")
        'If Not IsNull(pFeatTRoute.value(index)) Then
        'tempString = tempString + " " + CStr(pFeatTRoute.value(index))
        'Else
        'tempString = tempString + " 1"
        'End If
        Return tempString

    End Function
    Public Sub writeTransitLineInfo(ByVal lTOD As Long, ByVal _clsTransitLineAtts As clsTransitLineAtts, currentTimeOfDay As String)

        Dim sLine As String


        If lTOD = 0 Then

            If _clsTransitLineAtts.Headway_AM > 0 Then
                sLine = GetTransitLineInfo(_clsTransitLineAtts, _clsTransitLineAtts.Headway_AM)
                PrintLine(1, sLine)
            End If

            If _clsTransitLineAtts.Headway_MD > 0 Then
                sLine = GetTransitLineInfo(_clsTransitLineAtts, _clsTransitLineAtts.Headway_MD)
                PrintLine(2, sLine)
            End If
            If _clsTransitLineAtts.Headway_PM > 0 Then
                sLine = GetTransitLineInfo(_clsTransitLineAtts, _clsTransitLineAtts.Headway_PM)
                PrintLine(3, sLine)
            End If
            If _clsTransitLineAtts.Headway_EV > 0 Then
                sLine = GetTransitLineInfo(_clsTransitLineAtts, _clsTransitLineAtts.Headway_EV)
                PrintLine(4, sLine)
            End If
            If _clsTransitLineAtts.Headway_NI > 0 Then
                sLine = GetTransitLineInfo(_clsTransitLineAtts, _clsTransitLineAtts.Headway_NI)
                PrintLine(5, sLine)
            End If
        Else
            'we cannot build each time of day at once. 
            If currentTimeOfDay = "AM" Then
                sLine = GetTransitLineInfo(_clsTransitLineAtts, _clsTransitLineAtts.Headway_AM)
                PrintLine(1, sLine)
            ElseIf currentTimeOfDay = "MD" Then
                sLine = GetTransitLineInfo(_clsTransitLineAtts, _clsTransitLineAtts.Headway_MD)
                PrintLine(2, sLine)
            ElseIf currentTimeOfDay = "PM" Then
                sLine = GetTransitLineInfo(_clsTransitLineAtts, _clsTransitLineAtts.Headway_PM)
                PrintLine(3, sLine)
            ElseIf currentTimeOfDay = "EV" Then
                sLine = GetTransitLineInfo(_clsTransitLineAtts, _clsTransitLineAtts.Headway_EV)
                PrintLine(4, sLine)
            ElseIf currentTimeOfDay = "NI" Then
                sLine = GetTransitLineInfo(_clsTransitLineAtts, _clsTransitLineAtts.Headway_NI)
                PrintLine(5, sLine)
            End If
        End If

    End Sub
    Public Function writeTransitNode2(ByVal lTOD As Long, ByVal sLine As String, ByVal _clsTransitLineAtts As clsTransitLineAtts, currentTimeOfDay As String)
        Dim t As Integer
        If lTOD = 0 Then
            If _clsTransitLineAtts.Headway_AM > 0 Then
                PrintLine(1, sLine)
            End If
            If _clsTransitLineAtts.Headway_MD > 0 Then
                PrintLine(2, sLine)
            End If
            If _clsTransitLineAtts.Headway_PM > 0 Then
                PrintLine(3, sLine)
            End If
            If _clsTransitLineAtts.Headway_EV > 0 Then
                PrintLine(4, sLine)
            End If
            If _clsTransitLineAtts.Headway_NI > 0 Then
                PrintLine(5, sLine)
            End If
        Else
            'can do all tod's at once:
            If currentTimeOfDay = "AM" Then
                PrintLine(1, sLine)
            ElseIf currentTimeOfDay = "MD" Then
                PrintLine(2, sLine)
            ElseIf currentTimeOfDay = "PM" Then
                PrintLine(3, sLine)
            ElseIf currentTimeOfDay = "EV" Then
                PrintLine(4, sLine)
            ElseIf currentTimeOfDay = "NI" Then
                PrintLine(5, sLine)
            End If
        End If
    End Function

End Module


