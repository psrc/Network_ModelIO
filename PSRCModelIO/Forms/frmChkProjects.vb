
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
Imports ESRI.ArcGIS.ADF
Imports ESRI.ArcGIS.DataSourcesFile
Public Class frmChkProjects
    Private _passedFilePath As String
    Private _passedIMap As IMap
    Private _passedIApp As IApplication
    Private _passed_fXchk As Boolean
    Private _passedScenarioDesc As String
    Private _passedScenarioTitle As String
    Private _passedScenarioID As String
    Public colScenarioProjects As Collection
    Public colScenarioEvents As Collection
    Public my_clsCreateScenarioShapfiles As New clsCreateScenarioShapfiles
    Public WithEvents g_frmScenarioProjects As New frmScenarioProjects
    Public WithEvents g_frmCreateScenarioFromSelProjects As New frmCreateScenarioFromSelProjects
    Public _scenarioName As String
    Private _passedOldTAZ As Boolean

    Private Sub btnStopCmd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSkipProjects.Click

        colScenarioEvents = New Collection
        colScenarioProjects = New Collection
        'unload(Me)
        '[061207]hyu: call create scenarioshapefile
        Dim pathName As String
        'Me.Close()

        'pathname = "D:\projects\psrc\test"
        g_FWS = getPGDws(Me.PassedIMap)

        pathName = Me.PassedFilePath
        'pathName = Left(pathName, InStrRev(pathName, "\") - 1)

        Debug.Print("Start create scenario shapefiles: " & Now())
        'create_ScenarioShapefiles(pathName, "t1Edge", pathName, "t1Jct", pathName, "t1")




        my_clsCreateScenarioShapfiles.pApp = PassedIApp
        my_clsCreateScenarioShapfiles.pMap = PassedIMap
        my_clsCreateScenarioShapfiles.fXchk = passed_fXchk
        my_clsCreateScenarioShapfiles.pathnameE = pathName
        my_clsCreateScenarioShapfiles.filenameE = "t1Edge"
        my_clsCreateScenarioShapfiles.pathnameJ = pathName
        my_clsCreateScenarioShapfiles.filenameJ = "t1Jct"
        my_clsCreateScenarioShapfiles.pathnameN = pathName
        my_clsCreateScenarioShapfiles.filenameN = "t1"
        my_clsCreateScenarioShapfiles.prjselectedCol = colScenarioProjects
        my_clsCreateScenarioShapfiles.evtselectedCol = colScenarioEvents
        my_clsCreateScenarioShapfiles.scenarioTitle = passedScenarioTitle
        my_clsCreateScenarioShapfiles.scenarioTitle = passedScenarioDescription

        my_clsCreateScenarioShapfiles.scenarioTitle = m_ScenarioId
        my_clsCreateScenarioShapfiles.oldTAZ = passedOldTAZ
        my_clsCreateScenarioShapfiles.cSS()



        Debug.Print("end create scenario shapefiles: " & Now())

    End Sub

    Private Sub frmChkProjects_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
    Public Property [PassedFilePath]() As String
        Get
            Return _passedFilePath
        End Get
        Set(ByVal value As String)
            _passedFilePath = value

        End Set
    End Property
    Public Property [PassedIMap]() As IMap
        Get
            Return _passedIMap
        End Get
        Set(ByVal value As IMap)
            _passedIMap = value
        End Set
    End Property

    Public Property [PassedIApp]() As IApplication
        Get
            Return _passedIApp
        End Get
        Set(ByVal value As IApplication)
            _passedIApp = value
        End Set
    End Property
    Public Property [passed_fXchk]() As Boolean
        Get
            Return _passed_fXchk
        End Get
        Set(ByVal value As Boolean)
            _passed_fXchk = value
        End Set
    End Property
    Public Property [ScenarioName]() As String
        Get
            Return _scenarioName
        End Get
        Set(ByVal value As String)
            _scenarioName = value
        End Set
    End Property
    Public Property [passedScenarioTitle]() As String
        Get
            Return _passedScenarioTitle
        End Get
        Set(ByVal value As String)
            _passedScenarioTitle = value
        End Set
    End Property
    Public Property [passedScenarioDescription]() As String
        Get
            Return _passedScenarioDesc
        End Get
        Set(ByVal value As String)
            _passedScenarioDesc = value
        End Set
    End Property
    Public Property [passedOldTAZ]() As Boolean
        Get
            Return _passedOldTAZ
        End Get
        Set(ByVal value As Boolean)
            _passedOldTAZ = value
        End Set
    End Property
    Public Sub create_ScenarioShapefiles(ByVal pathnameE As String, ByVal filenameE As String, ByVal pathnameJ As String, ByVal filenameJ As String, ByVal pathnameN As String, ByVal filenameN As String)
        'pan I think On Error doesn't work in VBA
        'On Error GoTo ErrChk
        'called by frmNetLayer to create intermediate shapefiles for inputting to emme/2
        'these shapefiles will be used by output process to bring back the model run results to lay on top of the TransRef edges
        '  Dim pathname As String
        '  pathname = "c:\MODULEcreate_shp.txt"
        '  Open pathname For Output As #2
        '  Print #2, "Model Input create  shapefiles"
        '  Print #2, pathnameE
        '  pathname = "c:\createshp_ERROR.txt"
        '  Open pathname For Output As #1
        '  Print #1, "Create_ScenarioShapefiles error log"
        'jaf [051805] switched to standard logging modules


        '[061207]hyu: all the passed arguments are either not used in this procedure or their values are reset.
        'those arguments are for creating shapefile in a shapefile workspace not feature classes on a SDE workspace.
        WriteLogLine("")
        WriteLogLine("===============================================")
        WriteLogLine("create_ScenarioShapefiles started " & Now())
        WriteLogLine("===============================================")
        WriteLogLine("")
        Dim MyfrmNetLayer As New frmNetLayer

        '  WriteLogLine "pathnameE=" & pathnameE

        MyfrmNetLayer.lblStatus.Text = ("GlobalMod.create_ScenarioShapefiles: Starting")
        MyfrmNetLayer.Refresh()

        '  '[031606] hyu: comment this block to save time for testing purpose
        'MsgBox ("Create_ScenarioRow")
        Debug.Print("create ScenarioRow: " & Now())
        create_ScenarioRow()

        Debug.Print("create ProjectSceRow: " & Now())
        create_ProjectScenRows()

        'MsgBox ("Create_ProjectScenRows")
        Dim intIndexJunctType As Integer      'index to field JunctionType

        Dim pStatusBar As IStatusBar
        Dim pProgbar As IStepProgressor
        Dim pFeatLayerE As IFeatureLayer, pFeatLayerJ As IFeatureLayer
        Dim pTblLine As ITable

        Dim pFCS As IFeatureCursor
        Dim pWSedit As IWorkspaceEdit
        Dim pFeatSelectJ As IFeatureSelection
        Dim pointcnt As Long
        Dim pVersionedObject As IVersionedObject
        '[040406] pan duplicate
        'Dim pWSEdit As IWorkspaceEdit

        Dim pIndexes As IIndexes
        Dim penumindex As IEnumIndex
        Dim i As Long
        Dim lIndex As Long
        Dim pPolygon As IPolygon


        'pStatusBar = m_App.StatusBar
        'pProgbar = pStatusBar.ProgressBar

        'MsgBox ("Before get_MasterNetwork0")
        pFeatLayerE = get_MasterNetwork(0, _passedIMap)
        pFeatLayerJ = get_MasterNetwork(1, _passedIMap)

        'Dim ptblPoint As ITable no longer using tblPoint table
        'create the workspace to save the shapefiles to

        ' Dim pFWS As IFeatureWorkspace, pFWS2 As IFeatureWorkspace
        ' Dim pWSF As IWorkspaceFactory
        '  pWSF = New ShapefileWorkspaceFactory
        '  pFWS = pWSF.OpenFromFile(pathnameE, 0)
        'pan This will delete existing intermediate edges
        filenameE = m_layers(22)
        Debug.Print("Check Ds: " & filenameE & " - " & Now())

        '[031506] hyu: new codes to replace the following block which is commented out

        If Check_if_exist(g_FWS, filenameE) Then
            m_edgeShp = g_FWS.OpenFeatureClass(filenameE)
            'addField(m_edgeShp, "Dissolve", esriFieldType.esriFieldTypeInteger, , 0)
        Else
            '[031506] sec:  m_edgeShp(ScenarioEdge) = MasterNetwork (TransRefEdges)
            Debug.Print("create ScenarioEdge: " & Now())
            m_edgeShp = CreateScenarioEdge(pFeatLayerE.FeatureClass, g_FWS, filenameE)
        End If

        Debug.Print("search edge: " & Now())
        m_EdgeSSet.Search(Nothing, False, pFCS)
        Debug.Print("export edges: " & Now())

        pWSedit = g_FWS
        pWSedit.StartEditing(False)
        pWSedit.StartEditOperation()
        ExportFeature(pFCS, m_edgeShp)
        pWSedit.StopEditOperation()
        pWSedit.StopEditing(True)

        filenameJ = m_layers(21) 'ScenarioJunct
        pWSedit.StartEditing(False)
        Debug.Print("Check DS: " & filenameJ & " - " & Now())
        If Check_if_exist(g_FWS, filenameJ) Then
            m_junctShp = g_FWS.OpenFeatureClass(filenameJ)
        Else
            Debug.Print("create ScenarioJunct: " & Now())
            m_junctShp = CreateScenarioJunct(pFeatLayerJ.FeatureClass, g_FWS, filenameJ)
        End If
        '[040406] pan missing StartEditing
        pWSedit.StartEditing(False)
        pWSedit.StartEditOperation()
        selectJunctByAtt(m_edgeShp, pFeatLayerJ, m_junctShp)
        pWSedit.StopEditOperation()
        pWSedit.StopEditing(True)
        WriteLogLine("Select juncions m_junctshp.featurecount: " & m_junctShp.FeatureCount(Nothing))

        pFCS = Nothing

        '  pointcnt = m_JunctSSet.count
        Debug.Print("add index: " & Now())

        If (strLayerPrefix = "SDE") Then
            pVersionedObject = m_edgeShp
            If pVersionedObject.IsRegisteredAsVersioned = False Then
                'Registering a stand alone feature class as versioned
                pVersionedObject.RegisterAsVersioned(True)
            End If
        End If
        'add intermediate edge attribute indexes
        If m_edgeShp.Indexes.FindIndexesByFieldName("PSRCEDGEID").Next Is Nothing Then AddMyIndex(m_edgeShp, "PSRCEDGEID", "Idx_1")
        If m_edgeShp.Indexes.FindIndexesByFieldName("INODE").Next Is Nothing Then AddMyIndex(m_edgeShp, "INODE", "Idx_2")
        If m_edgeShp.Indexes.FindIndexesByFieldName("JNODE").Next Is Nothing Then AddMyIndex(m_edgeShp, "JNODE", "Idx_3")


        If (strLayerPrefix = "SDE") Then
            pVersionedObject = Nothing
            pVersionedObject = m_junctShp
            If pVersionedObject.IsRegisteredAsVersioned = False Then
                WriteLogLine("Registering a stand alone feature class as versioned")
                pVersionedObject.RegisterAsVersioned(True)
            End If
        End If
        'add scenariojunct attribute indexes
        'check if this had been created already

        pIndexes = m_junctShp.Indexes
        If pIndexes.FindIndexesByFieldName("PSRCJunctID").Next Is Nothing Then
            AddMyIndex(m_junctShp, "PSRCJunctID", "Idx_1")
        End If
        Debug.Print("calculate fields: " & Now())

        'Debug.Print "Add fields ends: " & Now()
        WriteLogLine("m_junctShp.Fields.FieldCount= " & m_junctShp.Fields.FieldCount)
        WriteLogLine(" ")
        WriteLogLine("m_edgeShp.Fields.FieldCount= " & m_edgeShp.Fields.FieldCount)
        'now to fill the values of these fields
        Dim tempID As Long
        Dim largestedge As Long
        LargestJunct = 0
        largestedge = 0

        Dim lSFld As Long, Tindex As Long
        Dim pGeometry As IGeometry2
        Dim pFeatCursor As IFeatureCursor
        Dim pFeatCursor2 As IFeatureCursor
        Dim pFeat As IFeature
        Dim pFeat2 As IFeature
        Dim pQF As IQueryFilter
        Dim useEmmeNode As Boolean
        Dim theexp As String
        Dim e2id As String

        'pan g_FWS sets the SDE workspace
        pWorkspaceI = g_FWS 'must be the intermediate workspace
        '[031606] hyu: calculate fields
        pWSedit.StartEditing(False)
        pWSedit.StartEditOperation()
        UpdateFields(m_edgeShp, m_junctShp)
        Debug.Print("end calc fields: " & Now())

        pWSedit.StopEditOperation()

        GetLargestID(m_junctShp, "PSRCJunctID", LargestJunct)

StartHere:

        pFeatCursor = Nothing
        pFeat2 = Nothing
        pFeat = Nothing
        pQF = Nothing

        WriteLogLine("finished w/fields")
        'now need to add the features into these newly created empty intermediate shapefiles
        Dim count As Long
        count = m_EdgeSSet.count '+ m_JunctSSet.count
        'pProgbar.position = 0
        'pStatusBar.ShowProgressBar("Creating Shapefiles...", 0, count, 1, True)

        'don't need the code below anymore with ExportFeatureLayer call
        'jaf--loop through all ref EDGES in pFeatCursor

        'jaf--let's keep the user updated
        g_frmNetLayer.lblStatus.Text = "GlobalMod.create_ScenarioShapefiles: finished ref edges"
        g_frmNetLayer.Refresh()

        'jaf--let's keep the user updated
        g_frmNetLayer.lblStatus.Text = "GlobalMod.create_ScenarioShapefiles: Finished creating Edge featureclass"
        g_frmNetLayer.Refresh()

        'pProgbar.position = 0
        'pStatusBar.ShowProgressBar("Creating Featureclasses...", 0, pointcnt, 1, True)

        Dim poldfeat As IFeature
        Dim Tindex2 As Long
        Dim test As Long
        test = 1
        'pStatusBar.HideProgressBar()
        'MsgBox "Finished creating Junction Layer", , "GlobalMod.create_ScenarioShapefiles"
        'jaf--let's keep the user updated
        g_frmNetLayer.lblStatus.Text = "GlobalMod.create_ScenarioShapefiles: Finished creating Junction featureclass"
        g_frmNetLayer.Refresh()

        '******************************************************************************
        'jaf--begin handling projects
        '******************************************************************************

        _passedIMap.ClearSelection()
        Dim PMx As IMxDocument
        Dim pFlds As IFields
        Dim pPrjFeatLayer As IFeatureLayer
        Dim pTopoOp As ITopologicalOperator2
        Dim pTC As ICursor
        Dim pRow As IRow
        Dim pFilter As ISpatialFilter
        Dim pFCj As IFeatureCursor, pFC As IFeatureCursor
        Dim rIndex As Long, idIndex As Long, index As Long
        Dim tDate As Date, t2Date As Date
        Dim tempString As String
        'Dim pPolygon As IPolygon

        PMx = _passedIApp.Document
        PMx.ActiveView.Refresh()
        ' ptblPoint = get_TableClass("tblPointProjects", pWorkspace)
        pTblLine = get_TableClass(m_layers(3))          'tblLineProjects
        pPrjFeatLayer = get_FeatureLayer2(m_layers(8))   'PRojectRoutes
        pFeatCursor = Nothing

        WriteLogLine(vbCrLf & "Starting add projects outcomes (in create_ScenarioShapefiles) at " & Now())
        'MsgBox "Adding projects outcomes", , "GlobalMod.create_ScenarioShapefiles"
        'jaf--let's keep the user updated
        g_frmNetLayer.lblStatus.Text = "GlobalMod.create_ScenarioShapefiles: Starting add projects"
        g_frmNetLayer.Refresh()

        'pProgbar.position = 0
        'pStatusBar.ShowProgressBar("Adding Projects...", 0, (prjselectedCol.Count + evtselectedCol.Count), 1, True)

        Dim j As Long, routeId As Long
        Dim pQFprjrte As IQueryFilter
        Dim pFCprjrte As IFeatureCursor
        Dim pPrjFeature As IFeature

        Dim tempprj As ClassPrjSelect
        Dim junctFeat As IFeature, pJFeat As IFeature
        Dim nIndex As Long, node As Long

        'projInode = pPrjFeature.value(projInodeIndex)
        Dim projInode As Long
        Dim projJnode As Long
        Dim projInodeIndex As Integer
        Dim projJnodeIndex As Integer
        Dim pTblPrjEdgeAtt As ITable
        Dim pRowPrjAtt As IRow
        Dim dirPrj As Integer 'flags the relationship between the project route's digitizing diriction _
        'and attributes direction.  It is updated in the updateProjectEdgeAttributes rouine.
        Dim iWithEvents As Integer

        pTblPrjEdgeAtt = get_TableClass(m_layers(24))

        '[061907] jaf: how many projects?
        WriteLogLine("Total projects selected for inclusion=" & prjselectedCol.Count)

        SortProjectsByYear()
        deleteAllRows(pWSedit, pTblPrjEdgeAtt)

        pWSedit.StartEditing(False)
        pWSedit.StartEditOperation()

        'NOTE- point projects will only be found in events table
        If (prjselectedCol.count > 0) Then 'get projects in list
            'jaf--projects selected in frmMapped are ALL in prjselectedCol collection
            '...regardless of origin or point/line

            For j = 1 To prjselectedCol.count 'loop though projects selected to either add them or update attributes

                dirPrj = 0
                pQF = New QueryFilter
                tempprj = prjselectedCol.Item(j)
                tempString = tempprj.PrjId 'get one
                If fVerboseLog Then WriteLogLine("Adding " + tempString)
                'jaf--let's keep the user updated
                g_frmNetLayer.lblStatus.Text = "GlobalMod.create_ScenarioShapefiles: project " & tempString
                g_frmNetLayer.Refresh()

                'jaf-- query filter to projID_Ver of project item j in collection
                'pQF.WhereClause = "projID_Ver = '" + tempString + "'"
                pQF.WhereClause = "projRteID = " + tempString
                '            If (pTblLine.rowcount(pQF) > 0) Then
                'jaf--there is a row in tblLineProjects for the current project in collection
                '... a table cursor pTC to tblLineProjects
                '                If fVerboseLog Then WriteLogLine "found project "
                'jaf--let's keep the user updated
                g_frmNetLayer.lblStatus.Text = "GlobalMod.create_ScenarioShapefiles: project " & tempString & " is a line project"
                g_frmNetLayer.Refresh()

                rIndex = pTblLine.FindField("projRteID")
                pTC = pTblLine.Search(pQF, False)
                'jaf--Melinda's comment in next line indicates many rows per project, which is no longer true
                pRow = pTC.NextRow ' can be more than one of these
                'jaf--get the route id for the current project and find the route feature in ProjectRoutes
                routeId = pRow.value(rIndex)

                pQFprjrte = New QueryFilter
                pQFprjrte.WhereClause = "projRteID = " & CStr(routeId)

                pFCprjrte = pPrjFeatLayer.FeatureClass.Search(pQFprjrte, False)
                pPrjFeature = pFCprjrte.NextFeature 'found it in ProjectRoutes
                iWithEvents = pPrjFeature.value(pPrjFeature.Fields.FindField("withEvents"))

                'Stefan- These are need for the GetDirectionFromProjects Sub
                projInodeIndex = pFCprjrte.Fields.FindField("INode")
                projInode = pPrjFeature.value(projInodeIndex)
                projJnodeIndex = pFCprjrte.Fields.FindField("JNode")
                projJnode = pPrjFeature.value(projJnodeIndex)

                'idIndex = pRow.Fields.FindField("PSRCEdgeID")
                Dim nodecnt As Long

                'jaf--spatial filter for the project route overlay operation
                m_Map.MapUnits = esriUnits.esriPoints

                pPolygon = New Polygon
                pTopoOp = pPrjFeature.Shape 'make a buffer object
                pPolygon = pTopoOp.Buffer(8)  'map units only which are in feet and need points here

                pFilter = New SpatialFilter
                With pFilter
                    .Geometry = pPolygon
                    .GeometryField = pFeatLayerE.FeatureClass.ShapeFieldName
                    '[061907] jaf: intersect returns too many edges when used with a buffer, let's try contains again
                    .SpatialRel = esriSpatialRelEnum.esriSpatialRelContains
                    '[122205] pan--.SpatialRelIntersects only valid for SDE layers
                    'should be faster operation
                    '.SpatialRel = esriSpatialRelIntersects
                End With

                'jaf--are there edges under the route?
                If (pFeatLayerE.FeatureClass.featurecount(pFilter) > 0) Then 'find underlying edges used by this project
                    'jaf--there are ref edges under current project route
                    'jaf-- a feature cursor pFeatCursor and loop through the  of edges
                    If fVerboseLog Then WriteLogLine("transRefEdges found under ProjRteID=" & tempString & " count=" & pFeatLayerE.FeatureClass.FeatureCount(pFilter))
                    'jaf--let's keep the user updated
                    g_frmNetLayer.lblStatus.Text = "GlobalMod.create_ScenarioShapefiles: project " & tempString & " has ref features"
                    g_frmNetLayer.Refresh()

                    '[072007] hyu: get the attributes of the project
                    pRowPrjAtt = pRow ' getProjectAtt(pPrjFeature)

                    pFeatCursor = Nothing
                    pFeatCursor = pFeatLayerE.FeatureClass.Search(pFilter, False)

                    '[072007] hyu: do not call the writeIntermedEdgesUnderProjects function
                    '                    '[070607] jaf new procedure to flag intermediate edge direction relative to overlying project route
                    '                    Call writeIntermedEdgesUnderProjects(pFeatLayerE.FeatureClass, m_edgeShp, pFilter, projInode, projJnode)
                    '                    '[070607] jaf subsequent code remains unchanged

                    poldfeat = pFeatCursor.NextFeature
                    Dim legn As Long
                    idIndex = poldfeat.Fields.FindField("PSRCEdgeID")
                    m_Map.MapUnits = esriUnits.esriFeet 'reset map units
                    Do Until poldfeat Is Nothing
                        pQF = New QueryFilter
                        '[062007] jaf: search should be made on PSRCEdgeID NOT the E2_ID
                        'pQF.WhereClause = "PSRC_E2ID = " + CStr(poldfeat.value(idIndex))
                        pQF.WhereClause = "PSRCEdgeID = " + CStr(poldfeat.value(idIndex))

                        '[072007] hyu: update edge attributes by project attributes
                        pFC = m_edgeShp.Search(pQF, True)
                        pFeat = pFC.NextFeature
                        'If (m_edgeShp.featurecount(pQF) < 1) Then 'not in the intermediate featureclass yet
                        If pFeat Is Nothing Then
                            'jaf--this check tests if the current ref edge is already in the intermediate shapefile

                            'jaf--let's keep the user updated
                            g_frmNetLayer.lblStatus.Text = "GlobalMod.create_ScenarioShapefiles: project " & tempString & " ref features need adding to featureclass"
                            g_frmNetLayer.Refresh()
                            If fVerboseLog Then WriteLogLine(pQF.WhereClause & " added to intermediate edges because of project " & tempString)
                            pQF = New QueryFilter
                            pFeat = Nothing
                            pFeat = m_edgeShp.CreateFeature
                            pFlds = m_edgeShp.Fields
                            pFeat.Shape = poldfeat.ShapeCopy
                            pFeat.Store()
                            'If fVerboseLog Then WriteLogLine "updated new edge"
                            For i = 0 To pFlds.FieldCount - 1
                                If (pFlds.field(i).Editable) Then
                                    lSFld = poldfeat.Fields.FindField(pFlds.field(i).Name)
                                    If (lSFld <> -1) Then
                                        If Not poldfeat.Value(lSFld) Is Nothing Then
                                            If fVerboseLog Then WriteLogLine("fieldname " & i & "=" & pFlds.Field(i).Name)
                                            pFeat.Value(i) = poldfeat.Value(lSFld)
                                        End If
                                    End If
                                    If UCase(pFlds.field(i).Name) = "PSRC_E2ID" Then
                                        Tindex = poldfeat.Fields.FindField("PSRCEdgeID")
                                        pFeat.value(i) = poldfeat.value(Tindex)
                                    End If
                                    If UCase(pFlds.field(i).Name) = UCase("ScenarioID") Then
                                        pFeat.value(i) = m_ScenarioId
                                    End If
                                    If UCase(pFlds.field(i).Name) = UCase("Scen_Link") Then
                                        Tindex = poldfeat.Fields.FindField("PSRCEdgeID")
                                        pFeat.value(i) = poldfeat.value(Tindex) + 1
                                    End If
                                End If
                                pFeat.Store()
                            Next i

                            'If fVerboseLog Then WriteLogLine "Writing update, shptype & prjrte flds in ScenarioEdge for "
                            'write the update, shptype and prjrte fields in intermediate edge layer

                            '[100207] hyu: the Updated1 field indicates where to look for the edge's attributes.
                            'When there's ONLY a point event on the edge, the edge's attributes are not update.
                            If iWithEvents <> 2 Then
                                index = pFeat.Fields.FindField("Updated1")
                                pFeat.value(index) = "Yes"
                            End If
                            If fVerboseLog Then WriteLogLine("here")
                            index = pFeat.Fields.FindField("prjRte")
                            '[061907] jaf: lets use routeID variable for consistency
                            'pFeat.value(index) = pRow.value(rIndex)
                            pFeat.value(index) = routeId

                            index = pFeat.Fields.FindField("shptype")
                            pFeat.value(index) = "Line"

                            Tindex = poldfeat.Fields.FindField("FunctionalClass")
                            lSFld = pFeat.Fields.FindField("FunctionalClass")
                            pFeat.value(lSFld) = poldfeat.value(Tindex)

                            If fVerboseLog Then WriteLogLine(pQF.WhereClause & " ADDED to ScenarioEdge & tagged to use attribs from ProjRteID=" & routeId)

                            'pStatusBar.StepProgressBar()
                            pFeat.Store()

                            'jaf--if the ref edge isn't in the shapefile, chances are...
                            '...its nodes aren't either.  following blocks check for the nodes
                            'now check if its from and to nodes in shapefile

                            '***********************************************************
                            'jaf-question: does this nodes insertion allow for project outcomes to alter node characteristics?
                            '***********************************************************


                            For nodecnt = 0 To 1
                                If nodecnt = 0 Then
                                    nIndex = poldfeat.Fields.FindField("INode")
                                Else
                                    nIndex = poldfeat.Fields.FindField("JNode")
                                End If

                                If Not poldfeat.Value(nIndex) Is Nothing Then
                                    'jaf--I think this means the ref edge HAS an Inode
                                    node = poldfeat.Value(nIndex)
                                    pQF = New QueryFilter
                                    pQF.WhereClause = "PSRCJunctID = " + CStr(node)

                                    pFCj = m_junctShp.Search(pQF, True)
                                    If (m_junctShp.FeatureCount(pQF) = 0) Then 'it is not yet in intermediate shapefile
                                        'jaf--Inode for current ref edge is not in intermediate shapefile
                                        '...so add it
                                        If fVerboseLog Then WriteLogLine("not yet in intermediatej")
                                        useEmmeNode = False
                                        'jaf--let's keep the user updated
                                        g_frmNetLayer.lblStatus.Text = "GlobalMod.create_ScenarioShapefiles: project " & tempString & " needs Inode added to featureclass"
                                        g_frmNetLayer.Refresh()

                                        pQF = New QueryFilter
                                        pQF.WhereClause = "PSRCJunctID = " + CStr(node)
                                        pFCj = pFeatLayerJ.Search(pQF, False)
                                        If fVerboseLog Then WriteLogLine("featurecount=" & pFeatLayerJ.FeatureClass.FeatureCount(pQF))
                                        junctFeat = pFCj.NextFeature
                                        If fVerboseLog Then WriteLogLine("junct feat OID=" & junctFeat.OID)
                                        pJFeat = m_junctShp.CreateFeature
                                        pointcnt = pointcnt + 1
                                        pGeometry = junctFeat.ShapeCopy
                                        pJFeat.Shape = pGeometry
                                        pJFeat.Store()

                                        If fVerboseLog Then WriteLogLine(" geometry")
                                        '******************************************************************************
                                        'jaf--this checks handling of centroids inserted because of projects
                                        '******************************************************************************
                                        'jaf--let's check to be sure we haven't put a project on a centroid or a link
                                        '     connecting to a centroid
                                        intIndexJunctType = junctFeat.Fields.FindField("JunctionType")
                                        If intIndexJunctType < 0 Then
                                            'FATAL ERROR--JunctionType field does NOT exist in TransRefJunctions
                                            WriteLogLine("DATA ERROR: JunctionType field does NOT exist in TransRefJunctions--GlobalMod.create_ScenarioShapefiles (proj insert Inode)")
                                            'Exit Sub
                                        End If
                                        If Not junctFeat.Value(intIndexJunctType) Is Nothing Then
                                            If junctFeat.Value(intIndexJunctType) = 6 Then
                                                'NONFATAL ERROR--projects shouldn't be on links connecting to a Centroid!
                                                WriteLogLine("DATA ERROR: Project ProjRteID=" & CStr(routeId) & " is on a centroid or a link with Inode on a centroid--GlobalMod.CreateScenarioShapefiles")
                                            End If
                                        End If

                                        'jaf--let's keep the user updated
                                        g_frmNetLayer.lblStatus.Text = "GlobalMod.create_ScenarioShapefiles: project " & tempString & " is past Inode centroid check"
                                        g_frmNetLayer.Refresh()

                                        'Tindex = pjfeat.Fields.FindField("PSRC_E2ID")
                                        'Tindex2 = junctFeat.Fields.FindField("PSRCJunctID")
                                        Tindex2 = junctFeat.Fields.FindField("EMME2nodeID")
                                        If Not junctFeat.Value(Tindex2) Is Nothing Then
                                            If junctFeat.Value(Tindex2) > 0 And junctFeat.Value(Tindex2) <= m_Offset Then  ' using this instead of PSRCJunctID
                                                tempID = junctFeat.Value(Tindex2)
                                                useEmmeNode = True
                                            Else
                                                Tindex2 = junctFeat.Fields.FindField("PSRCJunctID")
                                                If Not junctFeat.Value(Tindex2) Is Nothing Or junctFeat.Value(Tindex2) < 1 Then
                                                    WriteLogLine("DATA ERROR: No CentroidID value in TransRefJunctions for PSRCJunctID =" & pFeat.Value(Tindex) & "--GlobalMod.create_ScenarioShapefiles")
                                                    'Exit Sub
                                                End If
                                                tempID = junctFeat.Value(Tindex2) + m_Offset 'user defined offset
                                                useEmmeNode = False
                                            End If
                                        Else
                                            Tindex2 = junctFeat.Fields.FindField("PSRCJunctID") ' using this instead of PSRCJunctID

                                            'If junctFeat.value(Tindex2) = 10063 Then
                                            'MsgBox "found one"
                                            'End If
                                            If Not junctFeat.Value(Tindex2) Is Nothing Or junctFeat.Value(Tindex2) < 1 Then
                                                WriteLogLine("DATA ERROR: No CentroidID value in TransRefJunctions for PSRCJunctID =" & pFeat.Value(Tindex) & "--GlobalMod.create_ScenarioShapefiles")
                                                'Exit Sub
                                            End If
                                            tempID = junctFeat.Value(Tindex2) + m_Offset 'user defined offset
                                            useEmmeNode = False
                                            'should i now store this in tblJunctFacility or this only for parknrides?
                                        End If
                                        '[071306] pan:  largestjunct is assigned new highest value from ScenarioJunct including 1200
                                        'from m_Offset from above--is this right?
                                        '[061207]hyu: this is not right. largestJunct doesn't include the default offset.


                                        Tindex = pJFeat.Fields.FindField("Scen_Node")
                                        '                                    If (tempID > LargestJunct) Then
                                        '                                       MsgBox "largestjunct " & CStr(LargestJunct) & " to tempID " & CStr(tempID)
                                        '                                       LargestJunct = tempID
                                        '                                    End If
                                        pJFeat.Value(Tindex) = tempID
                                        If fVerboseLog Then WriteLogLine("scennode " + CStr(tempID))
                                        Tindex = pJFeat.Fields.FindField("ScenarioID")
                                        pJFeat.Value(Tindex) = m_ScenarioId
                                        If fVerboseLog Then WriteLogLine("here")
                                        pJFeat.Value(Tindex) = junctFeat.Value(Tindex2)
                                        Tindex = pJFeat.Fields.FindField("JunctionType")
                                        Tindex2 = junctFeat.Fields.FindField("JunctionType")
                                        pJFeat.Value(Tindex) = junctFeat.Value(Tindex2)
                                        'Tindex = pjfeat.Fields.FindField("PctGreen")
                                        'Tindex2 = junctFeat.Fields.FindField("PctGreen")
                                        If fVerboseLog Then WriteLogLine("here")
                                        'pjfeat.value(Tindex) = junctFeat.value(Tindex2)
                                        Tindex = pJFeat.Fields.FindField("P_RStalls")
                                        Tindex2 = junctFeat.Fields.FindField("P_RStalls")
                                        pJFeat.Value(Tindex) = junctFeat.Value(Tindex2)
                                        If fVerboseLog Then WriteLogLine("here")
                                        Tindex = pJFeat.Fields.FindField("Modes")
                                        Tindex2 = junctFeat.Fields.FindField("Modes")
                                        pJFeat.Value(Tindex) = junctFeat.Value(Tindex2)
                                        Tindex = pJFeat.Fields.FindField("PSRCJunctID")
                                        Tindex2 = junctFeat.Fields.FindField("PSRCJunctID")
                                        pJFeat.Value(Tindex) = junctFeat.Value(Tindex2)

                                        pJFeat.Store()

                                        pQF = New QueryFilter
                                        If useEmmeNode = True Then
                                            If fVerboseLog Then WriteLogLine("find others that need modifying")
                                            pQF.WhereClause = "INode = " + CStr(pJFeat.Value(Tindex)) + " Or JNode =" + CStr(pJFeat.Value(Tindex))
                                            pFeatCursor2 = m_edgeShp.Search(pQF, False)
                                            pFeat2 = pFeatCursor2.NextFeature
                                            Do Until pFeat2 Is Nothing

                                                lIndex = pFeat2.Fields.FindField("INode")
                                                If pFeat2.Value(lIndex) = pJFeat.Value(Tindex) Then
                                                    pFeat2.Value(lIndex) = tempID
                                                    lIndex = pFeat2.Fields.FindField("UseEmmeN")

                                                    If pFeat2.Value(lIndex) > 0 Then pFeat2.Value(lIndex) = 3
                                                    If pFeat2.Value(lIndex) = 0 Then pFeat2.Value(lIndex) = 1
                                                End If
                                                lIndex = pFeat2.Fields.FindField("JNode")
                                                If pFeat2.Value(lIndex) = pJFeat.Value(Tindex) Then
                                                    pFeat2.Value(lIndex) = tempID
                                                    lIndex = pFeat2.Fields.FindField("UseEmmeN")

                                                    If pFeat2.Value(lIndex) > 0 Then pFeat2.Value(lIndex) = 3
                                                    If pFeat2.Value(lIndex) = 0 Then pFeat2.Value(lIndex) = 2
                                                End If
                                                pFeat2.Store()
                                                pFeat2 = pFeatCursor2.NextFeature
                                            Loop
                                            pFeat2 = Nothing
                                            pFeatCursor2 = Nothing
                                        End If
                                    End If 'jaf--end Inode for current ref edge is not in intermediate shapefile
                                End If  'jaf--end of I think this means the ref edge HAS an Inode
                            Next nodecnt

                        Else
                            'edge IS already in ScenarioEdge: need to flag it to use Project attribs not modeAttributes
                            If fVerboseLog Then WriteLogLine(pQF.WhereClause & " ALREADY IN ScenarioEdge;  tagged to use attribs from ProjRteID=" & routeId)
                            'jaf--let's keep the user updated
                            g_frmNetLayer.lblStatus.Text = "GlobalMod.create_ScenarioShapefiles: project " & tempString & " has ref edges in shp, will flag update"
                            g_frmNetLayer.Refresh()

                            'MsgBox "write over attr fields"
                            ' pFeat = pFC.NextFeature
                            'find update, shptype and prjrte

                            '[100207] hyu: the Updated1 field indicates where to look for the edge's attributes.
                            'When there's ONLY a point event on the edge, the edge's attributes are not update.
                            If iWithEvents <> 2 Then
                                index = pFeat.Fields.FindField("Updated1")
                                pFeat.value(index) = "Yes"
                            End If
                            index = pFeat.Fields.FindField("prjRte")
                            pFeat.value(index) = routeId
                            index = pFeat.Fields.FindField("shptype")
                            pFeat.value(index) = "Line"

                            pFeat.Store()
                            'MsgBox "updated"
                        End If 'edge is ALREADY in ScenarioEdge (m_edgeShp.featurecount(pQF) < 1)

                        '[072007] hyu: update edge attributes [100207] by project attributes when not ONLY with point events
                        If iWithEvents <> 2 Then UpdateProjectEdgeAttributes(m_ScenarioId, pPrjFeature, pRowPrjAtt, pFeat, pFeatLayerJ.FeatureClass, dirPrj)

                        pStatusBar.StepProgressBar()
                        poldfeat = pFeatCursor.NextFeature
                    Loop

                    'update the P_RStall value of the related junction when there is a point event on this project route.
                    If iWithEvents = 2 Then updateProjectJunctionAttributes(pPrjFeature)
                Else
                    'jaf--there are NO ref edges under current project route
                    WriteLogLine("NO REF EDGES found under ProjRteID =" & CStr(routeId) & ".  GlobalMod.create_ScenarioShapefiles")
                End If

                'ElseIf (ptblPoint.rowcount(pQFt) > 0) Then
                'jaf--there is NO row in tblLineProjects for the current project in collection...
                '...but there IS a row in tblPointProjects...
                '...so  a table cursor pTC to the rows in tblPointProjects with current ProjRteID
                'MsgBox "found it in point"

                'jaf--let's keep the user updated
                'frmNetlayer.lblStatus = "GlobalMod.create_ScenarioShapefiles: project " & tempstring & " is a point project"
                'pan frmNetLayer.Refresh
                'deleted code 01-19-05
                '            End If
            Next j
            'jaf--let's keep the user updated
            g_frmNetLayer.lblStatus.Text = "GlobalMod.create_ScenarioShapefiles: finished project routes"
            g_frmNetLayer.Refresh()

        End If

        'free memory
        pQFprjrte = Nothing
        pFCprjrte = Nothing
        pPrjFeature = Nothing
        pFCj = Nothing
        pPrjFeatLayer = Nothing
        pTC = Nothing
        pRow = Nothing
        pFC = Nothing
        pQF = Nothing
        pJFeat = Nothing
        junctFeat = Nothing
        pFeat2 = Nothing
        pFeatCursor2 = Nothing
        pFeat = Nothing
        poldfeat = Nothing
        pGeometry = Nothing
        pTopoOp = Nothing
        pFilter = Nothing

        '******************************************************************************
        'jaf--commence handling of project events
        '******************************************************************************
        'jaf--evtselectedCol is the collection of events selected by user in frmMapped
        If (evtselectedCol.count > 0) Then
            Dim pPolyline As IPolyline
            Dim evtPoint As IFeatureClass
            Dim evtLine As IFeatureClass
            Dim FromPt As IPoint, ToPt As IPoint, AtPt As IPoint
            Dim pQFevt As IQueryFilter
            Dim pFCevt As IFeatureCursor
            'Dim pFilter As ISpatialFilter
            Dim pFeatEvt As IFeature
            Dim psrcIndex As Long, psrcValue As Long
            Dim pGeomEvt As IGeometry
            Dim pSSetevt As ISelectionSet
            Dim evtType As Integer ' 0 if point and 1 if line

            Dim firstnewJ As Long
            Dim secnewJ As Long
            Dim EprjRteID As String
            Dim pGeom As IGeometry
            Dim e As Integer, adde As Long
            Dim pJunctFeat As IFeature
            WriteLogLine("Now adding Events (create_ScenarioShapefiles)")
            'jaf--let's keep the user updated
            g_frmNetLayer.lblStatus.Text = "GlobalMod.create_ScenarioShapefiles: Starting event addition"
            g_frmNetLayer.Refresh()
            evtPoint = get_RouteEventSource(m_layers(7), m_App)
            evtLine = get_RouteEventSource(m_layers(5), m_App)

            For j = 1 To evtselectedCol.count 'loop though project events selected to either add events or update attributes
                tempprj = evtselectedCol.Item(j)
                tempString = tempprj.PrjId
                rIndex = evtLine.FindField("projRteID")
                pQFevt = New QueryFilter
                'pQFevt.WhereClause = "projID_Ver = '" + tempString + "' And InServiceDate <= " + CStr(inserviceyear)
                pQFevt.WhereClause = "projRteID=" + tempString

                If fVerboseLog Then WriteLogLine("line featurecount=" & evtLine.FeatureCount(pQFevt))
                If (evtLine.featurecount(pQFevt) > 0) Then
                    pFCevt = evtLine.Search(pQFevt, False)
                    pFeatEvt = pFCevt.NextFeature
                    pGeomEvt = pFeatEvt.ShapeCopy

                    pPolyline = pGeomEvt
                    FromPt = pPolyline.FromPoint
                    ToPt = pPolyline.ToPoint

                    evtType = 1
                    EprjRteID = pFeatEvt.value(rIndex)
                    If fVerboseLog Then WriteLogLine("here evtline")
                ElseIf (evtPoint.featurecount(pQFevt) > 0) Then
                    pFCevt = evtPoint.Search(pQFevt, False)
                    pFeatEvt = pFCevt.NextFeature
                    pGeomEvt = pFeatEvt.ShapeCopy
                    AtPt = pGeomEvt
                    evtType = 0
                    EprjRteID = pFeatEvt.value(rIndex)
                    If fVerboseLog Then WriteLogLine("here evtpoint")
                End If

                adde = 0

                If fVerboseLog Then WriteLogLine("evtType=" & evtType)

                For e = 0 To evtType
                    If fVerboseLog Then WriteLogLine("e=" & e)

                    If evtType = 1 And e = 0 Then AtPt = FromPt
                    If evtType = 1 And e = 1 Then AtPt = ToPt
                    pFilter = New SpatialFilter
                    With pFilter
                        .Geometry = AtPt
                        .GeometryField = evtPoint.ShapeFieldName
                        .SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects
                    End With
                    pFC = Nothing
                    'first see if in m_junctShp- might be at measure 0
                    WriteLogLine("junct captured=" & CStr(m_junctShp.FeatureCount(pFilter)))

                    If (m_junctShp.featurecount(pFilter) < 1) Then
                        'need to see what edge falls- first check if in m_edge

                        pFeatCursor = m_edgeShp.Search(pFilter, True) 'transrefedge search
                        WriteLogLine("edges captured=" & CStr(m_edgeShp.featurecount(pFilter)))
                        If m_edgeShp.featurecount(pFilter) < 1 Then 'need to add feature
                            If fVerboseLog Then WriteLogLine("need to add feature")
                            pFC = pFeatLayerE.FeatureClass.Search(pFilter, True) 'transrefedge search
                            If fVerboseLog Then WriteLogLine("intermediate edge featurecount=" & pFeatLayerE.FeatureClass.FeatureCount(pFilter))
                            poldfeat = pFC.NextFeature
                            'psrcValue = poldfeat.value(psrcIndex)
                            pFeat = m_edgeShp.CreateFeature
                            pFlds = m_edgeShp.Fields

                            pFeat.Shape = poldfeat.ShapeCopy
                            pFeat.Store()

                            If fVerboseLog Then WriteLogLine("created shape for additional feature")

                            For i = 0 To pFlds.FieldCount - 1
                                lSFld = poldfeat.Fields.FindField(pFlds.field(i).Name)
                                If (lSFld <> -1) Then
                                    If (IsDBNull(poldfeat.Value(lSFld)) = False) Then
                                        pFeat.Value(i) = poldfeat.Value(lSFld)
                                    End If
                                End If
                                If (pFlds.field(i).Name = "PSRC_E2ID") Then
                                    Tindex = poldfeat.Fields.FindField("PSRCEdgeID")
                                    pFeat.value(i) = poldfeat.value(Tindex)
                                End If
                                If (pFlds.field(i).Name = "ScenarioID") Then
                                    pFeat.value(i) = m_ScenarioId
                                End If
                                If (pFlds.field(i).Name = "Scen_Link") Then
                                    Tindex = poldfeat.Fields.FindField("PSRCEdgeID")
                                    pFeat.value(i) = poldfeat.value(Tindex)
                                End If
                                pFeat.Store()

                            Next i

                            If fVerboseLog Then WriteLogLine(" fields for additional feature")

                            'next build check to see if junctions in here
                            For nodecnt = 0 To 1
                                If nodecnt = 0 Then
                                    nIndex = poldfeat.Fields.FindField("INode")
                                Else
                                    nIndex = poldfeat.Fields.FindField("JNode")
                                End If

                                If fVerboseLog Then WriteLogLine("check points")

                                If (IsDBNull(poldfeat.Value(nIndex)) = False) Then
                                    'jaf--I think this means the ref edge HAS an Inode
                                    node = poldfeat.Value(nIndex)
                                    pQF = New QueryFilter
                                    pQF.WhereClause = "PSRCJunctID = " + CStr(node)

                                    pFCj = m_junctShp.Search(pQF, True)
                                    If (m_junctShp.FeatureCount(pQF) = 0) Then 'it is not yet in intermediate shapefile
                                        'jaf--Inode for current ref edge is not in intermediate shapefile
                                        '...so add it
                                        If fVerboseLog Then WriteLogLine("not yet in intermediatej")
                                        useEmmeNode = False
                                        'jaf--let's keep the user updated
                                        g_frmNetLayer.lblStatus.Text = "GlobalMod.create_ScenarioShapefiles: project " & tempString & " needs Inode added to shp"
                                        g_frmNetLayer.Refresh()

                                        pQF = New QueryFilter
                                        pQF.WhereClause = "PSRCJunctID = " + CStr(node)
                                        pFCj = pFeatLayerJ.Search(pQF, False)
                                        junctFeat = pFCj.NextFeature
                                        pJFeat = m_junctShp.CreateFeature
                                        pointcnt = pointcnt + 1
                                        pGeometry = junctFeat.ShapeCopy
                                        pJFeat.Shape = pGeometry

                                        '******************************************************************************
                                        'jaf--this checks handling of centroids inserted because of projects
                                        '******************************************************************************
                                        'jaf--let's check to be sure we haven't put a project on a centroid or a link
                                        '     connecting to a centroid
                                        intIndexJunctType = junctFeat.Fields.FindField("JunctionType")
                                        If intIndexJunctType < 0 Then
                                            'FATAL ERROR--JunctionType field does NOT exist in TransRefJunctions
                                            WriteLogLine("DATA ERROR: JunctionType field does NOT exist in TransRefJunctions--GlobalMod.create_ScenarioShapefiles (proj insert Inode)")
                                            Exit Sub
                                        End If
                                        If Not IsDBNull(junctFeat.Value(intIndexJunctType)) Then
                                            If junctFeat.Value(intIndexJunctType) = 6 Then
                                                'NONFATAL ERROR--projects shouldn't be on links connecting to a Centroid!
                                                WriteLogLine("DATA ERROR: Project ProjRteID=" & CStr(routeId) & " is on a centroid or a link with Inode on a centroid--GlobalMod.CreateScenarioShapefiles")
                                            End If
                                        End If

                                        'jaf--let's keep the user updated
                                        g_frmNetLayer.lblStatus.Text = "GlobalMod.create_ScenarioShapefiles: project " & tempString & " is past Inode centroid check"
                                        g_frmNetLayer.Refresh()

                                        'Tindex = pjfeat.Fields.FindField("PSRC_E2ID")
                                        'Tindex2 = junctFeat.Fields.FindField("PSRCJunctID")
                                        Tindex2 = junctFeat.Fields.FindField("EMME2nodeID")
                                        If Not IsDBNull(junctFeat.Value(Tindex2)) Then
                                            If junctFeat.Value(Tindex2) > 0 And junctFeat.Value(Tindex2) <= m_Offset Then
                                                tempID = junctFeat.Value(Tindex2)
                                                useEmmeNode = True
                                            Else
                                                Tindex2 = junctFeat.Fields.FindField("PSRCJunctID") ' using this instead of PSRCJunctID
                                                If IsDBNull(junctFeat.Value(Tindex2)) Or junctFeat.Value(Tindex2) < 1 Then
                                                    WriteLogLine("DATA ERROR: No CentroidID value in TransRefJunctions for PSRCJunctID =" & pFeat.Value(Tindex) & "--GlobalMod.create_ScenarioShapefiles")
                                                End If
                                                tempID = junctFeat.Value(Tindex2) + m_Offset 'user defined offset
                                                useEmmeNode = False
                                            End If
                                        Else
                                            Tindex2 = junctFeat.Fields.FindField("PSRCJunctID") ' using this instead of PSRCJunctID


                                            If IsDBNull(junctFeat.Value(Tindex2)) Or junctFeat.Value(Tindex2) < 1 Then
                                                WriteLogLine("DATA ERROR: No CentroidID value in TransRefJunctions for PSRCJunctID =" & pFeat.Value(Tindex) & "--GlobalMod.create_ScenarioShapefiles")
                                                'Exit Sub
                                            End If
                                            tempID = junctFeat.Value(Tindex2) + m_Offset 'user defined offset
                                            useEmmeNode = False
                                            'should i now store this in tblJunctFacility or this only for parknrides?
                                        End If
                                        Tindex = pJFeat.Fields.FindField("Scen_Node")
                                        'tempID = junctFeat.value(Tindex2)
                                        If (tempID > LargestJunct) Then
                                            LargestJunct = tempID
                                            'MsgBox "largest" + CStr(largestjunct)
                                        End If
                                        pJFeat.Value(Tindex) = tempID
                                        Tindex = pJFeat.Fields.FindField("ScenarioID")
                                        pJFeat.Value(Tindex) = m_ScenarioId

                                        pJFeat.Value(Tindex) = junctFeat.Value(Tindex2)
                                        Tindex = pJFeat.Fields.FindField("JunctionType")
                                        Tindex2 = junctFeat.Fields.FindField("JunctionType")
                                        pJFeat.Value(Tindex) = junctFeat.Value(Tindex2)

                                        Tindex = pJFeat.Fields.FindField("P_RStalls")
                                        Tindex2 = junctFeat.Fields.FindField("P_RStalls")
                                        pJFeat.Value(Tindex) = junctFeat.Value(Tindex2)
                                        Tindex = pJFeat.Fields.FindField("Modes")
                                        Tindex2 = junctFeat.Fields.FindField("Modes")
                                        pJFeat.Value(Tindex) = junctFeat.Value(Tindex2)
                                        Tindex = pFeat.Fields.FindField("PSRCJunctID")
                                        Tindex2 = junctFeat.Fields.FindField("PSRCJunctID")
                                        pJFeat.Value(Tindex) = junctFeat.Value(Tindex2)
                                        pJFeat.Store()

                                        pQF = New QueryFilter
                                        If useEmmeNode = True Then
                                            If fVerboseLog Then WriteLogLine("find others that need modifying")
                                            pQF.WhereClause = "INode = " + CStr(pJFeat.Value(Tindex)) + " Or JNode =" + CStr(pJFeat.Value(Tindex))
                                            pFeatCursor2 = m_edgeShp.Search(pQF, False)
                                            pFeat2 = pFeatCursor2.NextFeature

                                            Do Until pFeat2 Is Nothing
                                                lIndex = pFeat2.Fields.FindField("INode")
                                                If pFeat2.Value(lIndex) = pJFeat.Value(Tindex) Then
                                                    pFeat2.Value(lIndex) = tempID
                                                    lIndex = pFeat2.Fields.FindField("UseEmmeN")
                                                    If pFeat2.Value(lIndex) > 0 Then pFeat2.Value(lIndex) = 3
                                                    If pFeat2.Value(lIndex) = 0 Then pFeat2.Value(lIndex) = 1
                                                End If
                                                lIndex = pFeat2.Fields.FindField("JNode")
                                                If pFeat2.Value(lIndex) = pJFeat.Value(Tindex) Then
                                                    pFeat2.Value(lIndex) = tempID
                                                    lIndex = pFeat2.Fields.FindField("UseEmmeN")
                                                    If pFeat2.Value(lIndex) > 0 Then pFeat2.Value(lIndex) = 3
                                                    If pFeat2.Value(lIndex) = 0 Then pFeat2.Value(lIndex) = 2
                                                End If
                                                pFeat2.Store()
                                                pFeat2 = pFeatCursor2.NextFeature
                                            Loop
                                            pFeat2 = Nothing
                                            pFeatCursor2 = Nothing
                                        End If
                                    End If 'jaf--end Inode for current ref edge is not in intermediate shapefile
                                End If  'jaf--end of I think this means the ref edge HAS an Inode
                            Next nodecnt

                        Else
                            If fVerboseLog Then WriteLogLine("do NOT need to add feature")
                            pFeat = pFeatCursor.NextFeature
                        End If

                        If fVerboseLog Then WriteLogLine("should be splitting")

                        adde = 1
                        'addNewJunction EprjRteID, evtType, AtPt

                        If fVerboseLog Then WriteLogLine("before: m_junctShp.featurecount=" & m_junctShp.FeatureCount(Nothing))
                        'Dim pJunctFeat As IFeature 'to use to add new junctions created from events
                        pJunctFeat = m_junctShp.CreateFeature


                        pGeom = AtPt
                        pJunctFeat.Shape = AtPt
                        pJunctFeat.Store()


                        index = pJunctFeat.Fields.FindField("Scen_Node")
                        LargestJunct = LargestJunct + 1
                        ' MsgBox "in add new junction"
                        pJunctFeat.value(index) = LargestJunct
                        index = pJunctFeat.Fields.FindField("PSRCJunctID")
                        pJunctFeat.value(index) = LargestJunct
                        pJunctFeat.Store()

                        If fVerboseLog Then WriteLogLine("after: m_junctShp.featurecount=" & m_junctShp.FeatureCount(Nothing))
                        'pWorkspaceEdit.StopEditOperation
                        'pWorkspaceEdit.StopEditing True
                        'WriteLogLine m_junctShp.featurecount(Nothing)
                        'pWorkspaceEdit.StartEditing False
                        'pWorkspaceEdit.StartEditOperation
                        SplitFeature(pFeat, AtPt, EprjRteID)
                        If fVerboseLog Then WriteLogLine("m_edgeShp.featurecount=" & m_edgeShp.FeatureCount(Nothing))
                        pFeat.Store() 'this new edge
                    End If

                    If evtType = 1 And e = 1 Then
                        m_Map.MapUnits = esriUnits.esriPoints
                        pFilter = New SpatialFilter
                        pTopoOp = pFeatEvt.Shape
                        pPolygon = New Polygon
                        pPolygon = pTopoOp.Buffer(100)
                        If fVerboseLog Then WriteLogLine(CStr(pPolygon.Envelope.Height) + " ' " + CStr(pPolygon.Envelope.Width))
                        With pFilter
                            .Geometry = pFeatEvt.Shape
                            .GeometryField = m_edgeShp.ShapeFieldName
                            .SpatialRel = esriSpatialRelEnum.esriSpatialRelWithin
                        End With
                        pFeatCursor = Nothing
                        pFeatCursor = m_edgeShp.Search(pFilter, True) 'transrefedge search
                        WriteLogLine("edges captured " + CStr(m_edgeShp.FeatureCount(pFilter)))
                        If m_edgeShp.featurecount(pFilter) = 0 Then
                            With pFilter
                                .Geometry = pFeatEvt.Shape
                                .GeometryField = m_edgeShp.ShapeFieldName
                                .SpatialRel = esriSpatialRelEnum.esriSpatialRelOverlaps
                            End With
                            WriteLogLine("edges captured " + CStr(m_edgeShp.FeatureCount(pFilter)))
                        End If
                        'find update, shptype and prjrte
                        poldfeat = Nothing
                        poldfeat = pFeatCursor.NextFeature
                        Do Until poldfeat Is Nothing
                            index = poldfeat.Fields.FindField("Updated1")
                            poldfeat.value(index) = "Yes"
                            index = poldfeat.Fields.FindField("prjRte")
                            poldfeat.value(index) = EprjRteID
                            index = poldfeat.Fields.FindField("shptype")
                            poldfeat.value(index) = "Event"
                            poldfeat.Store()
                            poldfeat = pFeatCursor.NextFeature
                        Loop

                    End If
                    pJunctFeat = Nothing
                    pGeom = Nothing
                    pStatusBar.StepProgressBar()
                Next e
            Next j  'end loop though projects selected to either add them or update attributes

            pPolyline = Nothing
            evtPoint = Nothing
            evtLine = Nothing
            pQFevt = Nothing
            pFCevt = Nothing
            FromPt = Nothing
            ToPt = Nothing
            AtPt = Nothing
            pFeatEvt = Nothing
            pGeomEvt = Nothing

        End If
        WriteLogLine("finished adding projects and events at " & Now())

        'pWorkspaceEdit.StopEditOperation
        'pWorkspaceEdit.StopEditing True

        pFilter = Nothing
        'free space up for next routines
        pFeatLayerJ = Nothing
        pFeatLayerE = Nothing
        pTblLine = Nothing
        ' ptblPoint = Nothing
        pFeat = Nothing
        pJFeat = Nothing
        pFeat2 = Nothing
        poldfeat = Nothing
        pGeometry = Nothing
        pRow = Nothing
        pFeatCursor = Nothing
        pFC = Nothing
        pFeatCursor2 = Nothing
        pFCj = Nothing
        '     pFieldsE = Nothing
        '     pFieldsJ = Nothing

        'MsgBox "Finished creating intermediate shapefiles", , "GlobalMod.create_ScenarioShapefiles"
        'jaf--let's keep the user updated
        g_frmNetLayer.lblStatus.Text = "GlobalMod.create_ScenarioShapefiles: Finished creating intermediate featureclasses"
        g_frmNetLayer.Refresh()

        _passedIMap.ClearSelection()

        PMx = _passedIApp.Document
        PMx.ActiveView.Refresh()
        'pStatusBar.HideProgressBar()

        'jaf--check if EMME/2 node limit is exceeded [061405] removed auto-start if pointcount
        'If (pointcnt > 15000 Or pseudothin = True) Then
        If pseudothin Then
            'jaf--let's keep the user updated
            g_frmNetLayer.lblStatus.Text = "GlobalMod.create_ScenarioShapefiles: Starting pseudo dissolve"
            g_frmNetLayer.Refresh()
            If _passed_fXchk Then
                Dim objDissolve As New clsDissolve(_passedIApp)
                objDissolve.DissolveEdges()
                'dissolveTransEdges3()  'does NOT thin features used by tolls, transit, and turns
            Else
                dissolveTransEdgesNoX()   'thins ALL pseudonodes REGARDLESS of toll, transit, turns
            End If
        End If

        If g_frmModelQuery.chkDisolveOnly.Checked Then
            'user chose to STOP after the dissolve
            WriteLogLine("STOPPING after shapefile creation per user choice")
        Else
            ' pan--keep user updated
            WriteLogLine("buildfile routines not called, run createBuildfiles separately")
            'user chose to do EVERYTHING
            ' WriteLogLine "finished shapefile creation, calling buildfile routines from within create_ScenarioShapefiles"
            'jaf--let's keep the user updated
            ' frmNetLayer.lblStatus = "GlobalMod.create_ScenarioShapefiles: Writing Park&Ride  buildfile from " & pFeatLayerIE.name
            ' frmNetLayer.Repaint
            ' createParkRideFile pathnameN
            ' frmNetLayer.lblStatus = "GlobalMod.create_ScenarioShapefiles: Writing Toll buildfile from " & pFeatLayerIE.name
            ' frmNetLayer.Repaint
            'createTollsFile pathnameN, filenameN
            ' frmNetLayer.lblStatus = "GlobalMod.create_ScenarioShapefiles: Writing turn buildfile from " & pFeatLayerIE.name
            ' frmNetLayer.Repaint
            ' create_TurnFile pathnameN, filenameN, pFeatLayerIE
            '  frmNetLayer.lblStatus = "GlobalMod.create_ScenarioShapefiles: Writing nodes & links buildfile"
            ' frmNetLayer.Repaint
            ' create_NetFile2 pathnameN, filenameN, pWorkspaceI
            If fBuildTransit Then
                ' pan--keep user informed
                WriteLogLine("Transit build not yet functioning, currently omitted")
                '   frmNetLayer.lblStatus = "GlobalMod.create_ScenarioShapefiles: Writing transit routes buildfile"
                '   frmNetLayer.Repaint
                '   create_TransitFile pathnameN, filenameN
            Else
                WriteLogLine("")
                WriteLogLine("=====================")
                WriteLogLine("TRANSIT BUILD OMITTED")
                WriteLogLine("=====================")
                WriteLogLine("")
            End If

        End If

        '     pFeatlayerIE = Nothing
        '     pFeatLayerIJ = Nothing
        pWorkspaceI = Nothing

        WriteLogLine("FINISHED create_ScenarioShapefiles at " & Now())
        'Close
        Exit Sub

ErrChk:
        'jaf [051805] retrofitted standard error log
        CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.create_ScenarioShapefiles")
        MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.create_ScenarioShapefiles")
        'Print #1, Err.Description, vbInformation, "create_ScenarioShapefiles"
        'get out of edit operation
        'If pWorkspaceEdit.IsBeingEdited Then
        '  pWorkspaceEdit.AbortEditOperation
        '  pWorkspaceEdit.StopEditing False
        'End If
        pStatusBar.HideProgressBar()
        pFilter = Nothing
        'free space up for next routines
        pFeatLayerJ = Nothing
        pFeatLayerE = Nothing
        pTblLine = Nothing
        ' ptblPoint = Nothing
        pFeat = Nothing
        pJFeat = Nothing
        pFeat2 = Nothing
        poldfeat = Nothing
        pGeometry = Nothing
        pRow = Nothing
        pFeatCursor = Nothing
        pFC = Nothing
        pFeatCursor2 = Nothing
        pFCj = Nothing
        '     pFieldsE = Nothing
        '     pFieldsJ = Nothing
        '     pFeatlayerIE = Nothing
        '     pFeatLayerIJ = Nothing
        pWorkspaceI = Nothing
    End Sub

    Public Function get_FeatureLayer2(ByVal featClsname As String) As IFeatureLayer
        'returns the feature class stated in featClsname
        'on error GoTo eh




        Dim pEnumLy As IEnumLayer


        'm_Map = pActiveView.FocusMap
        pEnumLy = _passedIMap.Layers

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
    Public Sub dissolveTransEdges3()
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

        _passedIMap.ClearSelection()
        Dim PMx As IMxDocument
        PMx = _passedIApp.Document
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

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub

    Private Sub btnOpenProj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOpenProj.Click
        'right now this is just holding an empty collection to satisfy some subs in the clsCreateScenarioShapefiles.
        'Need to check to see how events are being use/handled
        'Using pComReleaser As ComReleaser = New ComReleaser
        colScenarioEvents = New Collection

        g_FWS = getPGDws(PassedIMap)
        Dim colProjectsInScenario As New Collection
        Dim pathName As String

        If rdoProjectScenario.Checked = True Then

            'holds all the scenarios to be passed into the drop down list
            colScenarioProjects = FindProsepctiveScenarios(g_ModelYear)

            'passing some variables into the scenario form
            g_frmScenarioProjects.PassedIApp = Me.PassedIApp
            g_frmScenarioProjects.PassedIMap = Me.PassedIMap
            g_frmScenarioProjects.passed_fXchk = Me.passed_fXchk
            g_frmScenarioProjects.PassedFilePath = Me.PassedFilePath
            g_frmScenarioProjects.PassedScenarioNameCollection = colScenarioProjects
            'show the form
            g_frmScenarioProjects.ShowDialog()

            If g_frmScenarioProjects.DialogResult = Windows.Forms.DialogResult.OK Then
                _scenarioName = g_frmScenarioProjects.scenarioName
                g_frmScenarioProjects.Dispose()
                GC.Collect()
                'loads the projects(ids) form the scenario into a collection
                colProjectsInScenario = loadScenarioSelections(_scenarioName)

         
                g_FWS = getPGDws(PassedIMap)

                pathName = PassedFilePath
                my_clsCreateScenarioShapfiles.pApp = PassedIApp
                my_clsCreateScenarioShapfiles.pMap = PassedIMap
                my_clsCreateScenarioShapfiles.fXchk = passed_fXchk
                my_clsCreateScenarioShapfiles.pathnameE = pathName
                my_clsCreateScenarioShapfiles.filenameE = "t1Edge"
                my_clsCreateScenarioShapfiles.pathnameJ = pathName
                my_clsCreateScenarioShapfiles.filenameJ = "t1Jct"
                my_clsCreateScenarioShapfiles.pathnameN = pathName
                my_clsCreateScenarioShapfiles.filenameN = "t1"
                my_clsCreateScenarioShapfiles.prjselectedCol = colProjectsInScenario
                my_clsCreateScenarioShapfiles.evtselectedCol = colScenarioEvents
                my_clsCreateScenarioShapfiles.scenarioTitle = passedScenarioTitle
                my_clsCreateScenarioShapfiles.scenarioTitle = passedScenarioDescription
                my_clsCreateScenarioShapfiles.scenarioTitle = m_ScenarioId
                my_clsCreateScenarioShapfiles.oldTAZ = passedOldTAZ
                'start building scenario edges:
                my_clsCreateScenarioShapfiles.cSS()
            End If

            'Use selected projects to run and create scenario:
        ElseIf rdoCreateScenarioFromProjectSelection.Checked = True Then
           
            Dim pProjectsFL As IFeatureLayer2 = get_FeatureLayer10(m_layers(8), Me.PassedIMap)
            Dim pFSel As IFeatureSelection
            Dim pFeatureSelection As ISelection
            Dim pFeatureCursor As IFeatureCursor
            Dim pProjectFeature As IFeature
            Dim indexProjRteID As Long
            Dim pProjRteID As Long
            Dim selectedProject As ClassPrjSelect

            'get the selection set of projects (projects selected in ArcMap to build/run a new scenario
            pFSel = pProjectsFL
            pFSel.SelectionSet.Search(Nothing, False, pFeatureCursor)
            pProjectFeature = pFeatureCursor.NextFeature
            'loads the projects(ids) form the selected projects into a collection
            Do Until pProjectFeature Is Nothing
                indexProjRteID = pProjectFeature.Fields.FindField("PROJRTEID")

                selectedProject = New ClassPrjSelect
                selectedProject.PrjId = pProjectFeature.Value(indexProjRteID)
                'MsgBox "add to multiselect" + CStr(selectedprj.PrjId)
                'GlobalMod.databaseOpen = "MTP"

                colProjectsInScenario.Add(selectedProject)

                pProjectFeature = pFeatureCursor.NextFeature

            Loop

            g_FWS = getPGDws(PassedIMap)

            pathName = PassedFilePath
            my_clsCreateScenarioShapfiles.pApp = PassedIApp
            my_clsCreateScenarioShapfiles.pMap = PassedIMap
            my_clsCreateScenarioShapfiles.fXchk = passed_fXchk
            my_clsCreateScenarioShapfiles.pathnameE = pathName
            my_clsCreateScenarioShapfiles.filenameE = "t1Edge"
            my_clsCreateScenarioShapfiles.pathnameJ = pathName
            my_clsCreateScenarioShapfiles.filenameJ = "t1Jct"
            my_clsCreateScenarioShapfiles.pathnameN = pathName
            my_clsCreateScenarioShapfiles.filenameN = "t1"
            my_clsCreateScenarioShapfiles.prjselectedCol = colProjectsInScenario
            my_clsCreateScenarioShapfiles.evtselectedCol = colScenarioEvents
            my_clsCreateScenarioShapfiles.scenarioTitle = passedScenarioTitle
            my_clsCreateScenarioShapfiles.scenarioDesc = passedScenarioDescription
            my_clsCreateScenarioShapfiles.oldTAZ = passedOldTAZ
            'start building scenario edges:
            my_clsCreateScenarioShapfiles.cSS()

        End If
        
    End Sub
    Public Function FindProsepctiveScenarios(ByVal intModelYear As Integer) As Collection

        Dim pRelClass As IRelationshipClass
        pRelClass = g_FWS.OpenRelationshipClass(g_schema & "tblProjectsInScenariosTotblModelScenario")
        Dim pTblModelScenario As ITable
        Dim pTblProjectsInScenarios As ITable
        pTblModelScenario = get_TableClass(m_layers(16)) 'tblModelScenario
        pTblProjectsInScenarios = get_TableClass(m_layers(15))
        Dim pCs As ICursor

        Dim pRow As IRow

        Dim pOriginFeature As IFeature
        Dim pRelatedFeature As IFeature
        Dim pObjSet As ISet
        Dim pEnumFeature As IEnumFeature
        Dim yr As Long
        Dim lIndex As Integer
        Dim lIndex2 As Integer
        Dim pRow2 As IRow
        Dim keep As Boolean

        yr = intModelYear
        pCs = pTblModelScenario.Search(Nothing, False)
        pRow = pCs.NextRow
        FindProsepctiveScenarios = New Collection
        Do Until pRow Is Nothing
            keep = False
            pObjSet = pRelClass.GetObjectsRelatedToObject(pRow)
            pObjSet.Reset()

            If pObjSet.Count > 0 Then
                pRow2 = pObjSet.Next
                lIndex = pRow2.Fields.FindField("InServiceDate")
                Do Until pRow2 Is Nothing

                    If pRow2.Value(lIndex) <= yr Then
                        keep = True
                    Else
                        keep = False
                    End If
                    pRow2 = pObjSet.Next
                Loop
                lIndex2 = pRow.Fields.FindField("title")

                If keep = True Then
                    'FindProsepctiveScenarioProjects = New Collection
                    FindProsepctiveScenarios.Add(pRow.Value(lIndex2))

                End If
            End If


            pRow = pCs.NextRow
        Loop





    End Function
    Private Function loadScenarioSelections(ByVal scenarioName As String) As Collection
        Dim prjselectedcol As Collection

        Dim i As Long
        prjselectedcol = New Collection
        GlobalMod.evtselectedCol = New Collection
        Dim pTblModelScenario As ITable
        Dim pTblProjectsInScenarios As ITable
        pTblModelScenario = get_TableClass(m_layers(16)) 'tblModelScenario
        pTblProjectsInScenarios = get_TableClass(m_layers(15))

        Dim indexTitle As Integer
        Dim indexScenarioID As Integer
        Dim intProjRteID As Integer
        indexTitle = pTblModelScenario.FindField("Title")
        indexScenarioID = pTblModelScenario.FindField("Scenario_ID")
        intProjRteID = pTblProjectsInScenarios.FindField("projRteID")
        Dim pCs As ICursor
        Dim pFilter As IQueryFilter
        Dim pRow As IRow
        Dim intScenarioID As Integer
        'Dim selectedprj As ClassPrjSelect
        pFilter = New QueryFilter

        pFilter.WhereClause = "Title= '" & scenarioName & "'"
        pCs = pTblModelScenario.Search(pFilter, False)
        pRow = pCs.NextRow
        intScenarioID = pRow.Value(indexScenarioID)

        Dim pFilter2 As IQueryFilter
        pFilter2 = New QueryFilter
        pFilter.WhereClause = "Scenario_ID= '" & intScenarioID & "'"
        pCs = pTblProjectsInScenarios.Search(pFilter, True)
        pRow = pCs.NextRow
        Dim selectedProject As ClassPrjSelect
        Do Until pRow Is Nothing
            selectedProject = New ClassPrjSelect
            selectedProject.PrjId = pRow.Value(intProjRteID)
            'MsgBox "add to multiselect" + CStr(selectedprj.PrjId)
            'GlobalMod.databaseOpen = "MTP"
            prjselectedcol.Add(selectedProject)
            'GlobalMod.prjselectedCol.Add (pRow.value(intProjRteID))
            pRow = pCs.NextRow
        Loop

        loadScenarioSelections = prjselectedcol










    End Function

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub g_frmScenarioProjects_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles g_frmScenarioProjects.Disposed
        'Me.ShowDialog()
    End Sub
    Public Function get_FeatureLayer10(ByVal featClsname As String, pMap As IMap) As IFeatureLayer
        'returns the feature class stated in featClsname
        'on error GoTo eh




        Dim pEnumLy As IEnumLayer

       

        pEnumLy = pMap.Layers

        Dim pLy As ILayer, pFLy As IFeatureLayer
        pLy = pEnumLy.Next
        Do Until pLy Is Nothing

            If pLy.Valid Then
                If TypeOf pLy Is IFeatureLayer Then
                    pFLy = pLy

                    'If UCase(pFLy.FeatureClass.AliasName) = UCase(featClsname) Then
                    If UCase(pFLy.Name) = UCase(featClsname) Then
                        get_FeatureLayer10 = pFLy
                        Exit Do

                    End If
                End If
            End If
            pLy = pEnumLy.Next
        Loop

        If Not get_FeatureLayer10 Is Nothing Then Exit Function

        Dim pFeatLayer As IFeatureLayer
        pFeatLayer = New FeatureLayer
        Dim pfeatcls As IFeatureClass
        pfeatcls = getFeatureClass(featClsname)
        pFeatLayer.FeatureClass = pfeatcls

        get_FeatureLayer10 = pFeatLayer
        If pFeatLayer Is Nothing Then
            MsgBox("Error: did not find the feature class " + featClsname)
        End If

        Exit Function

eh:
        MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.get_FeatureLayer")

    End Function
End Class