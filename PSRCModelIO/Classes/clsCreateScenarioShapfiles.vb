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
'Imports Scripting

Imports ESRI.ArcGIS.DataSourcesFile

Public Class clsCreateScenarioShapfiles
    Public Sub New()

    End Sub
    Public pMap As IMap
    Public pApp As IApplication
    Public pathnameE As String
    Public filenameE As String
    Public pathnameJ As String
    Public filenameJ As String
    Public pathnameN As String
    Public filenameN As String
    Public fXchk As Boolean
    Public prjselectedCol As Collection
    Public evtselectedCol As Collection
    Public scenarioDesc As String
    Public scenarioTitle As String
    Public scenarioID As Long
    Public oldTAZ As Boolean









    Public Sub cSS()
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
        create_ScenarioRow2()

        Debug.Print("create ProjectSceRow: " & Now())
        create_ProjectScenRows2()

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
        pFeatLayerE = get_MasterNetwork(0, pMap)
        pFeatLayerJ = get_MasterNetwork(1, pMap)

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
        count = m_EdgeSSet.Count '+ m_JunctSSet.count
        'pProgbar.position = 0
        'pStatusBar.ShowProgressBar("Creating Shapefiles...", 0, count, 1, True)

        'don't need the code below anymore with ExportFeatureLayer call
        'jaf--loop through all ref EDGES in pFeatCursor

        'jaf--let's keep the user updated
        'g_frmNetLayer.lblStatus.Text = "GlobalMod.create_ScenarioShapefiles: finished ref edges"
        'g_frmNetLayer.Refresh()

        'jaf--let's keep the user updated
        'g_frmNetLayer.lblStatus.Text = "GlobalMod.create_ScenarioShapefiles: Finished creating Edge featureclass"
        'g_frmNetLayer.Refresh()

        'pProgbar.position = 0
        'pStatusBar.ShowProgressBar("Creating Featureclasses...", 0, pointcnt, 1, True)

        Dim poldfeat As IFeature
        Dim Tindex2 As Long
        Dim test As Long
        test = 1
        'pStatusBar.HideProgressBar()
        'MsgBox "Finished creating Junction Layer", , "GlobalMod.create_ScenarioShapefiles"
        'jaf--let's keep the user updated
        'g_frmNetLayer.lblStatus.Text = "GlobalMod.create_ScenarioShapefiles: Finished creating Junction featureclass"
        'g_frmNetLayer.Refresh()

        '******************************************************************************
        'jaf--begin handling projects
        '******************************************************************************

        pMap.ClearSelection()
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

        PMx = pApp.Document
        PMx.ActiveView.Refresh()
        ' ptblPoint = get_TableClass("tblPointProjects", pWorkspace)
        pTblLine = get_TableClass(m_layers(3))          'tblLineProjects
        pPrjFeatLayer = get_FeatureLayer2(m_layers(8))   'PRojectRoutes
        pFeatCursor = Nothing

        WriteLogLine(vbCrLf & "Starting add projects outcomes (in create_ScenarioShapefiles) at " & Now())
        'MsgBox "Adding projects outcomes", , "GlobalMod.create_ScenarioShapefiles"
        'jaf--let's keep the user updated
        'g_frmNetLayer.lblStatus.Text = "GlobalMod.create_ScenarioShapefiles: Starting add projects"
        'g_frmNetLayer.Refresh()

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
        If (prjselectedCol.Count > 0) Then 'get projects in list
            'jaf--projects selected in frmMapped are ALL in prjselectedCol collection
            '...regardless of origin or point/line



            For j = 1 To prjselectedCol.Count 'loop though projects selected to either add them or update attributes
                Try
                    GC.Collect()
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
                    routeId = pRow.Value(rIndex)

                    pQFprjrte = New QueryFilter
                    pQFprjrte.WhereClause = "projRteID = " & CStr(routeId)

                    pFCprjrte = pPrjFeatLayer.FeatureClass.Search(pQFprjrte, False)
                    pPrjFeature = pFCprjrte.NextFeature 'found it in ProjectRoutes
                    iWithEvents = pPrjFeature.Value(pPrjFeature.Fields.FindField("withEvents"))

                    'Stefan- These are need for the GetDirectionFromProjects Sub
                    projInodeIndex = pFCprjrte.Fields.FindField("INode")
                    projInode = pPrjFeature.Value(projInodeIndex)
                    projJnodeIndex = pFCprjrte.Fields.FindField("JNode")
                    projJnode = pPrjFeature.Value(projJnodeIndex)

                    'idIndex = pRow.Fields.FindField("PSRCEdgeID")
                    Dim nodecnt As Long

                    'jaf--spatial filter for the project route overlay operation
                    Try


                        pMap.MapUnits = esriUnits.esriPoints
                    Catch ex As Exception
                        MessageBox.Show(ex.ToString)
                    End Try
                    pPolygon = New Polygon
                    pTopoOp = pPrjFeature.Shape 'make a buffer object
                    pPolygon = pTopoOp.Buffer(5)  'map units only which are in feet and need points here

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
                    If (pFeatLayerE.FeatureClass.FeatureCount(pFilter) > 0) Then 'find underlying edges used by this project
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
                        pMap.MapUnits = esriUnits.esriFeet 'reset map units
                        Do Until poldfeat Is Nothing
                            pQF = New QueryFilter
                            '[062007] jaf: search should be made on PSRCEdgeID NOT the E2_ID
                            'pQF.WhereClause = "PSRC_E2ID = " + CStr(poldfeat.value(idIndex))
                            pQF.WhereClause = "PSRCEdgeID = " + CStr(poldfeat.Value(idIndex))

                            '[072007] hyu: update edge attributes by project attributes
                            pFC = m_edgeShp.Search(pQF, False)
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
                                    If (pFlds.Field(i).Editable) Then
                                        lSFld = poldfeat.Fields.FindField(pFlds.Field(i).Name)
                                        If (lSFld <> -1) Then
                                            If Not poldfeat.Value(lSFld) Is Nothing Then
                                                If fVerboseLog Then WriteLogLine("fieldname " & i & "=" & pFlds.Field(i).Name)
                                                pFeat.Value(i) = poldfeat.Value(lSFld)
                                            End If
                                        End If
                                        If UCase(pFlds.Field(i).Name) = "PSRC_E2ID" Then
                                            Tindex = poldfeat.Fields.FindField("PSRCEdgeID")
                                            pFeat.Value(i) = poldfeat.Value(Tindex)
                                        End If
                                        If UCase(pFlds.Field(i).Name) = UCase("ScenarioID") Then
                                            pFeat.Value(i) = m_ScenarioId
                                        End If
                                        If UCase(pFlds.Field(i).Name) = UCase("Scen_Link") Then
                                            Tindex = poldfeat.Fields.FindField("PSRCEdgeID")
                                            pFeat.Value(i) = poldfeat.Value(Tindex) + 1
                                        End If
                                    End If
                                    pFeat.Store()

                                Next i
                                'need to check if there is a change in Mode String and Facility Type
                                index = pRow.Fields.FindField("Modes")
                                If pRow.Value(index) <> "-1" Then
                                    pFeat.Value(pFeat.Fields.FindField("Modes")) = pRow.Value(index)
                                End If
                                index = pPrjFeature.Fields.FindField("Change_Type")
                                If pPrjFeature.Value(index) <> -1 Then
                                    pFeat.Value(pFeat.Fields.FindField("FacilityType")) = pPrjFeature.Value(index)
                                End If

                                pFeat.Store()

                                'If fVerboseLog Then WriteLogLine "Writing update, shptype & prjrte flds in ScenarioEdge for "
                                'write the update, shptype and prjrte fields in intermediate edge layer

                                '[100207] hyu: the Updated1 field indicates where to look for the edge's attributes.
                                'When there's ONLY a point event on the edge, the edge's attributes are not update.
                                If iWithEvents <> 2 Then
                                    index = pFeat.Fields.FindField("Updated1")
                                    pFeat.Value(index) = "Yes"
                                End If
                                If fVerboseLog Then WriteLogLine("here")
                                index = pFeat.Fields.FindField("prjRte")
                                '[061907] jaf: lets use routeID variable for consistency
                                'pFeat.value(index) = pRow.value(rIndex)
                                pFeat.Value(index) = routeId

                                index = pFeat.Fields.FindField("shptype")
                                pFeat.Value(index) = "Line"

                                Tindex = poldfeat.Fields.FindField("FunctionalClass")
                                lSFld = pFeat.Fields.FindField("FunctionalClass")
                                pFeat.Value(lSFld) = poldfeat.Value(Tindex)

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
                                                'See if there are other edges (facility connectors that are connected to the park and ride
                                                'and set UseEmmeN to a value so that the emme id will be used. 
                                                If fVerboseLog Then WriteLogLine("find others that need modifying")
                                                pQF.WhereClause = "INode = " + CStr(pJFeat.Value(Tindex)) + " Or JNode =" + CStr(pJFeat.Value(Tindex))
                                                pFeatCursor2 = m_edgeShp.Search(pQF, False)
                                                pFeat2 = pFeatCursor2.NextFeature
                                                Do Until pFeat2 Is Nothing

                                                    lIndex = pFeat2.Fields.FindField("INode")
                                                    If pFeat2.Value(lIndex) = pJFeat.Value(Tindex) Then
                                                        'Stefan- I dont think the IJ Nodes should change! Build portion figures this out!
                                                        'pFeat2.Value(lIndex) = tempID
                                                        lIndex = pFeat2.Fields.FindField("UseEmmeN")

                                                        If IsDBNull(pFeat2.Value(lIndex)) Then
                                                            pFeat2.Value(lIndex) = 1

                                                        ElseIf pFeat2.Value(lIndex) > 0 Then
                                                            pFeat2.Value(lIndex) = 3
                                                        ElseIf pFeat2.Value(lIndex) = 0 Then
                                                            pFeat2.Value(lIndex) = 1
                                                        End If


                                                    End If

                                                    lIndex = pFeat2.Fields.FindField("JNode")
                                                    If pFeat2.Value(lIndex) = pJFeat.Value(Tindex) Then
                                                        'Stefan- I dont think the IJ Nodes should change! Build portion figures this out!
                                                        'pFeat2.Value(lIndex) = tempID
                                                        lIndex = pFeat2.Fields.FindField("UseEmmeN")
                                                        If IsDBNull(pFeat2.Value(lIndex)) Then
                                                            pFeat2.Value(lIndex) = 2
                                                        ElseIf pFeat2.Value(lIndex) > 0 Then
                                                            pFeat2.Value(lIndex) = 3
                                                        ElseIf pFeat2.Value(lIndex) = 0 Then
                                                            pFeat2.Value(lIndex) = 2
                                                        End If
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
                                    pFeat.Value(index) = "Yes"
                                End If
                                index = pFeat.Fields.FindField("prjRte")
                                pFeat.Value(index) = routeId
                                index = pFeat.Fields.FindField("shptype")
                                pFeat.Value(index) = "Line"

                                'need to check if there is a change in Mode String and Facility Type
                                index = pRow.Fields.FindField("Modes")
                                If pRow.Value(index) <> "-1" Then
                                    pFeat.Value(pFeat.Fields.FindField("Modes")) = pRow.Value(index)
                                End If
                                index = pPrjFeature.Fields.FindField("Change_Type")
                                If pPrjFeature.Value(index) <> "-1" Then
                                    pFeat.Value(pFeat.Fields.FindField("FacilityType")) = CType(pPrjFeature.Value(index), Integer)
                                End If


                                pFeat.Store()
                                'MsgBox "updated"
                            End If 'edge is ALREADY in ScenarioEdge (m_edgeShp.featurecount(pQF) < 1)

                            '[072007] hyu: update edge attributes [100207] by project attributes when not ONLY with point events
                            If iWithEvents <> 2 Then
                                UpdateProjectEdgeAttributes(scenarioID, pPrjFeature, pRowPrjAtt, pFeat, pFeatLayerJ.FeatureClass, dirPrj)
                            End If


                            'pStatusBar.StepProgressBar()
                            poldfeat = pFeatCursor.NextFeature
                        Loop

                        'update the P_RStall value of the related junction when there is a point event on this project route.
                        If iWithEvents = 2 Then
                            updateProjectJunctionAttributes(pPrjFeature)
                        End If

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
                Catch ex As Exception
                    MessageBox.Show(ex.ToString)

                End Try   '            End If
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
        If (evtselectedCol.Count > 0) Then
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

            For j = 1 To evtselectedCol.Count 'loop though project events selected to either add events or update attributes
                tempprj = evtselectedCol.Item(j)
                tempString = tempprj.PrjId
                rIndex = evtLine.FindField("projRteID")
                pQFevt = New QueryFilter
                'pQFevt.WhereClause = "projID_Ver = '" + tempString + "' And InServiceDate <= " + CStr(inserviceyear)
                pQFevt.WhereClause = "projRteID=" + tempString

                If fVerboseLog Then WriteLogLine("line featurecount=" & evtLine.FeatureCount(pQFevt))
                If (evtLine.FeatureCount(pQFevt) > 0) Then
                    pFCevt = evtLine.Search(pQFevt, False)
                    pFeatEvt = pFCevt.NextFeature
                    pGeomEvt = pFeatEvt.ShapeCopy

                    pPolyline = pGeomEvt
                    FromPt = pPolyline.FromPoint
                    ToPt = pPolyline.ToPoint

                    evtType = 1
                    EprjRteID = pFeatEvt.Value(rIndex)
                    If fVerboseLog Then WriteLogLine("here evtline")
                ElseIf (evtPoint.FeatureCount(pQFevt) > 0) Then
                    pFCevt = evtPoint.Search(pQFevt, False)
                    pFeatEvt = pFCevt.NextFeature
                    pGeomEvt = pFeatEvt.ShapeCopy
                    AtPt = pGeomEvt
                    evtType = 0
                    EprjRteID = pFeatEvt.Value(rIndex)
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

                    If (m_junctShp.FeatureCount(pFilter) < 1) Then
                        'need to see what edge falls- first check if in m_edge

                        pFeatCursor = m_edgeShp.Search(pFilter, True) 'transrefedge search
                        WriteLogLine("edges captured=" & CStr(m_edgeShp.FeatureCount(pFilter)))
                        If m_edgeShp.FeatureCount(pFilter) < 1 Then 'need to add feature
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
                                lSFld = poldfeat.Fields.FindField(pFlds.Field(i).Name)
                                If (lSFld <> -1) Then
                                    If (IsDBNull(poldfeat.Value(lSFld)) = False) Then
                                        pFeat.Value(i) = poldfeat.Value(lSFld)
                                    End If
                                End If
                                If (pFlds.Field(i).Name = "PSRC_E2ID") Then
                                    Tindex = poldfeat.Fields.FindField("PSRCEdgeID")
                                    pFeat.Value(i) = poldfeat.Value(Tindex)
                                End If
                                If (pFlds.Field(i).Name = "ScenarioID") Then
                                    pFeat.Value(i) = m_ScenarioId
                                End If
                                If (pFlds.Field(i).Name = "Scen_Link") Then
                                    Tindex = poldfeat.Fields.FindField("PSRCEdgeID")
                                    pFeat.Value(i) = poldfeat.Value(Tindex)
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
                        pJunctFeat.Value(index) = LargestJunct
                        index = pJunctFeat.Fields.FindField("PSRCJunctID")
                        pJunctFeat.Value(index) = LargestJunct
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
                        If m_edgeShp.FeatureCount(pFilter) = 0 Then
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
                            poldfeat.Value(index) = "Yes"
                            index = poldfeat.Fields.FindField("prjRte")
                            poldfeat.Value(index) = EprjRteID
                            index = poldfeat.Fields.FindField("shptype")
                            poldfeat.Value(index) = "Event"
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

        pMap.ClearSelection()

        PMx = pApp.Document
        PMx.ActiveView.Refresh()
        'pStatusBar.HideProgressBar()

        'jaf--check if EMME/2 node limit is exceeded [061405] removed auto-start if pointcount
        'If (pointcnt > 15000 Or pseudothin = True) Then
        If pseudothin Then
            g_frmNetLayer.lblStatus.Text = "GlobalMod.create_ScenarioShapefiles: Starting pseudo dissolve"
            g_frmNetLayer.Refresh()
            If GlobalMod.FTthin = True Then
                Dim objDissolve As New clsDissolve(pApp)
                objDissolve.DissolveByFT()


                'jaf--let's keep the user updated


            ElseIf fXchk Then
                Dim objDissolve As New clsDissolve(pApp)
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
        pEnumLy = pMap.Layers



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
    Public Sub create_ScenarioRow2()
        'called by frmNetLayer to create a new entry in the tblModelScenario table
        'on error GoTo eh
        WriteLogLine("create_ScenarioRow called")
        WriteLogLine("-------------------------")
        'pan added edit environment for use with SDE


        Dim pScenTable As ITable
        Dim pRow As IRow, pNewRow As IRow
        Dim pTC As ICursor

        Dim pWorkspaceEdit As IWorkspaceEdit
        pWorkspaceEdit = g_FWS

        If (strLayerPrefix = "SDE") Then

            pWorkspaceEdit.StartEditing(False)
            pWorkspaceEdit.StartEditOperation()
        End If

        pScenTable = get_TableClass(m_layers(16))

        'MsgBox "After get_TableClass"

        pNewRow = pScenTable.CreateRow
        'MsgBox "After CreateRow in create_ScenarioRow"
        Dim index As Long
        Dim count As Long

        index = pScenTable.FindField("Scenario_ID")
        If (index = -1) Then
            WriteLogLine("Could not find field 'Scenario_ID' in tblModelScenario")
            Exit Sub
        End If
        count = 0
        'MsgBox pScenTable.rowcount(Nothing)
        If (pScenTable.RowCount(Nothing) > 0) Then
            pTC = pScenTable.Search(Nothing, False)
            pRow = pTC.NextRow
            Do Until pRow Is Nothing
                If (pRow.Value(index) > count) Then
                    count = pRow.Value(index)
                End If
                pRow = pTC.NextRow
            Loop
        End If
        m_ScenarioId = count + 1
        pNewRow.Value(index) = m_ScenarioId

        index = pScenTable.FindField("Title")
        If (index = -1) Then
            WriteLogLine("Could not find field 'Title' in tblModelScenario")
            'Exit Sub
        End If
        pNewRow.Value(index) = scenarioTitle
        index = pScenTable.FindField("Description")
        If (index = -1) Then
            WriteLogLine("Could not find field 'Description' in tblModelScenario")
            'Exit Sub
        End If
        pNewRow.Value(index) = scenarioDesc

        index = pScenTable.FindField("Date_Created")
        If (index = -1) Then
            WriteLogLine("Could not find field 'Date_Created' in tblModelScenario")
            'Exit Sub
        End If
        pNewRow.Value(index) = Date.Now
        index = pScenTable.FindField("NodeOffset")
        If (index = -1) Then
            WriteLogLine("Could not find field 'NodeOffset' in tblModelScenario")
            'Exit Sub
        End If
        pNewRow.Value(index) = m_Offset

        pNewRow.Store()
        pNewRow = Nothing
        pScenTable = Nothing

        If (strLayerPrefix = "SDE") Then
            pWorkspaceEdit.StopEditOperation()
            pWorkspaceEdit.StopEditing(True)
        End If

        WriteLogLine("finished create_ScenarioRow at " & Now())
        Exit Sub

eh:
        CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.create_ScenarioRow")
        MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.create_ScenarioRow")

        'Print #1, Err.Description, vbExclamation, "create_ScenarioRow"
        pNewRow = Nothing
        pScenTable = Nothing
    End Sub

    Public Sub create_ProjectScenRows2()
        'called by frmNetlayer to enter all the projects used in this model input prep into tblProjectsinScenario
        'MsgBox "scenario project"
        'on error GoTo ErrChk

        '[080207] hyu: update InServiceDate, projDBS, projRteID, projID, version
        '[080207] hyu: do not create a record in tblProjectsInScenarios when the project cannot be found in tblLineProjects.

        WriteLogLine("create_ProjectScenRows called")
        WriteLogLine("-----------------------------")
        Dim pInScenTable As ITable
        Dim pTblLine As ITable ' no longer use, ptblPoint As ITable
        Dim pRow As IRow, pNewRow As IRow
        Dim pTC As ICursor
        Dim pQF As IQueryFilter

        Dim pWorkspaceEdit As IWorkspaceEdit
        pWorkspaceEdit = g_FWS
        pWorkspaceEdit.StartEditing(False)
        pWorkspaceEdit.StartEditOperation()

        pInScenTable = get_TableClass(m_layers(15)) 'tblprojectsinscenarios
        pTblLine = get_TableClass(m_layers(3))
        '  ptblPoint = get_TableClass("tblPointProjects", pWs)

        Dim index As Long, dbindex As Long
        Dim count As Long, j As Long
        Dim tempprj As ClassPrjSelect
        Dim tempString As String


        For j = 1 To prjselectedCol.Count
            pQF = New QueryFilter
            tempprj = prjselectedCol.Item(j)
            tempString = tempprj.PrjId

            '[100207] hyu: query on "projRteID" instead of "projID_Ver"
            'pQF.WhereClause = "projID_Ver = '" + tempString + "'"
            pQF.WhereClause = "projRteID =" & tempString
            index = pInScenTable.FindField("Scenario_ID")
            If (index = -1) Then
                WriteLogLine("Could not find field 'Scenario_ID' in tblProjectsInScenario")
                'Exit Sub
            End If

            '         pNewRow = pInScenTable.CreateRow
            '        pNewRow.value(index) = m_ScenarioId

            If (pTblLine.RowCount(pQF) > 0) Then
                ' MsgBox "found it in line"
                pNewRow = pInScenTable.CreateRow
                pNewRow.Value(index) = m_ScenarioId

                pTC = pTblLine.Search(pQF, False)
                pRow = pTC.NextRow

                index = pInScenTable.FindField("projID_Ver")
                If (index = -1) Then
                    WriteLogLine("Could not find field 'projID_Ver' in tblProjectsInScenario")
                    'Exit Sub
                Else
                    pNewRow.Value(index) = pRow.Value(pRow.Fields.FindField("projID_Ver"))
                End If
                '[080808] SEC: Here, the index is pointing to the right field but wrong table! Should be pInScenTable not
                'pTblLine
                'index = pTblLine.FindField("projDBS")
                index = pInScenTable.FindField("projDBS")
                If (dbindex = -1) Then
                    WriteLogLine("Could not find field 'projDBS' in tblLineProjects")
                    'Exit Sub
                Else

                    pNewRow.Value(index) = pRow.Value(pRow.Fields.FindField("projDBS"))
                End If
                '             pTC = pTblLine.Search(pQF, False)
                '             pRow = pTC.NextRow

                index = pInScenTable.FindField("projID")
                If (index = -1) Then
                    WriteLogLine("Could not find field 'projID' in tblProjectsInScenario")
                    'Exit Sub
                Else
                    pNewRow.Value(index) = pRow.Value(pRow.Fields.FindField("projID"))
                End If

                index = pInScenTable.FindField("projRteID")
                If (index = -1) Then
                    WriteLogLine("Could not find field 'projRteID' in tblProjectsInScenario")
                    'Exit Sub
                Else
                    pNewRow.Value(index) = tempString
                End If

                index = pInScenTable.FindField("version")
                If (index = -1) Then
                    WriteLogLine("Could not find field 'version' in tblProjectsInScenario")
                    'Exit Sub
                Else
                    pNewRow.Value(index) = pRow.Value(pRow.Fields.FindField("version"))
                End If

                index = pInScenTable.FindField(g_InSvcDate)
                If (index = -1) Then
                    WriteLogLine("Could not find field '" & g_InSvcDate & "' in tblProjectsInScenario")
                    'Exit Sub
                Else
                    pNewRow.Value(index) = pRow.Value(pTblLine.FindField(g_InSvcDate))
                End If
                'ElseIf (ptblPoint.rowcount(pQF) > 0) Then
                'deleted code 01-19-05
                pNewRow.Store()
            Else
                '[021506] hyu: shouldn 't the pNewRow be deleted, if no project found in tbl_LineProjects?
                WriteLogLine("Could not find the project in tbl_LineProjects " + CStr(tempprj.PrjId))
            End If

        Next j

        For j = 1 To evtselectedCol.Count
            pQF = New QueryFilter
            tempprj = evtselectedCol.Item(j)
            tempString = tempprj.PrjId
            'MsgBox tempstring
            '[100207] hyu: search on "projRteID" instead of "projID_Ver"
            'pQF.WhereClause = "projID_Ver = '" + tempString + "'"
            pQF.WhereClause = "projRteID = " + tempString
            index = pInScenTable.FindField("Scenario_ID")
            pNewRow = pInScenTable.CreateRow
            pNewRow.Value(index) = m_ScenarioId
            'now just been saved in tblModelScenario
            ' index = pInScenTable.FindField("Title")
            ' pNewRow.value(index) = m_title

            If (pTblLine.RowCount(pQF) > 0) Then
                'MsgBox "found it in line"
                dbindex = pTblLine.FindField("projDBS")
                pTC = pTblLine.Search(pQF, False)
                pRow = pTC.NextRow

                '[021506] hyu: what 's the deal to have seperate condition here?
                '            If (pRow.value(dbindex) = "MTP") Then
                '             index = pInScenTable.FindField("projID_Ver")
                '             pNewRow.value(index) = tempString
                '            Else
                index = pInScenTable.FindField("projRteID") '("projID_Ver")
                pNewRow.Value(index) = tempString
                '            End If
                'ElseIf (ptblPoint.rowcount(pQF) > 0) Then
                'deleted code 01-19-05
            Else
                WriteLogLine("Could not find the project in tbl_LineProjects" + CStr(tempprj.PrjId))
            End If
            pNewRow.Store()
        Next j

        'If (strLayerPrefix = "SDE") Then
        pWorkspaceEdit.StopEditOperation()
        pWorkspaceEdit.StopEditing(True)
        'End If

        WriteLogLine("finished create_ProjectScenRows at " & Now())
        Exit Sub

ErrChk:
        CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.create_ProjectScenRows")
        MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.create_ProjectScenRows")

        'Print #1, Err.Description, vbExclamation, "create_ProjectScenRows"
        pNewRow = Nothing
        pInScenTable = Nothing
        pTblLine = Nothing
        pRow = Nothing
    End Sub
End Class

