
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

'Imports ESRI.ArcGIS.ADF



Module GlobalMod
    Public m_path As String
    Public m_App As IApplication
    Public m_Doc As IMxDocument
    Public m_Map As IMap
    Public pActiveView As IActiveView
    Public g_FWS As IFeatureWorkspace
    Public m_Editor As ESRI.ArcGIS.Editor.IEditor 'used for getting handle on Editor toolbar

    'global constants
    Public Const dbTIP As Integer = 0
    Public Const dbMTP As Integer = 1
    Public Const projTypePoint As Integer = 0
    Public Const projTypeLine As Integer = 1
    Public Const cStrAppVer As String = "0.2.0.2"
    'pan where is g_schema used?
    Public g_schema As String

    'pan this should be 22 not 21
    Public m_layers(25) As String

    'global error handling/debug variables
    Public fDebug As Boolean      'jaf--debugging flag
    Public m_RunStartTime As Date 'jaf--run start time for debug logging
    Public m_fLogOpen As Boolean  'jaf--true if log file is open for writes
    Public m_fFatalError As Boolean 'jaf-- true by module if it detects a fatal error
    '...it is responsibility of calling objects to test
    '(not yet enabled)

    'global state variables
    Public m_fEditMode            'true if frmIDFacilities has put app into edit mode
    Public m_fNetChanged          'true if user has made a change that would require a ArchiveNetwork to archive the net
    Public m_intCurrProjOrigin    'origin of current project in browser or frmIDfacilities
    '1=one of selected in frmUncodedProjects, 2=one of all in frmUncodedProjects, 3=current rec from frmProjectBrowse
    Public loaded As Boolean      'jaf 091206: not sure what this is

    'global variables about current user
    Public m_userLevel            'User's inherent security level:  0 = guest (read only), 2 = add, 4 = add+delete, 6 = add+delete+change, 9 = admin                                                       2 = edit map,  4 = edit map and data
    Public m_strUsername          'currently logged in user (for opening db)
    Public m_strUserpass          'currently logged in user's password (for opening db)

    'user selections
    Public fVerboseLog            ' true if all detailed log entries desired
    Public fBuildTransit          ' true if user wants transit buildfiles (thinning still checks for transit regardless)
    Public fXchk                  ' true if user wants pseuod thin to RETAIN features used by tolls, turns, and transit
    Public strLayerPrefix As String ' to SDE in init if SDE

    'global variables concerning external databases
    Public databaseOpen As String
    Public m_TableTIP As ITable 'handle on TIP Access database
    Public m_TableMTP As ITable 'handle on MTP Access database
    Public m_tablefile As String

    'model prep globals
    Public m_JunctSSet As ISelectionSet 'input prep's selection  of current network edges
    Public m_EdgeSSet As ISelectionSet 'input prep's selection  of current network edges
    Public inserviceDate As Date 'input prep's future inservice date
    Public inserviceyear As Long 'store the model run year
    'Private m_PointScenTable As ITable
    'Private m_LineScenTable As ITable

    Public LargestJunct As Long 'used by create_scenarioShapefile
    '[071206]- pan largestwjunct will be  from largestjunct. Shadow links were
    'being built using non-unique IDs from below 9001-10000.
    Public LargestWJunct As Long 'used by create_WeaveLink- is in range 9000-10000

    Private m_title As String 'scenario title from input frmModelQuery
    Private m_descript As String 'scenario description from input frmModelQuery
    Public m_ScenarioId As Long 'scenario ID created in Create_scenarioRow
    Public NetOutfilename As String 'filename of AM model run results
    Public NetOutfilenameMidday As String 'filename of Midday model run results
    Public NetOutfilenamePM As String 'filename of PM model run results
    Public NetOutfilenameEvening As String 'filename of Evening model run results
    Public NetOutfilenameNight As String 'filename of Night model run results
    Public TROutfilename As String 'filename of AM transit model run results
    Public TROutfilenameMidday As String 'filename of Midday transit model run results
    Public TROutfilenamePM As String 'filename of PM transit model run results
    Public TROutfilenameEvening As String 'filename of Evening transit model run results
    Public TROutfilenameNight As String 'filename of Night transit model run results

    'intermediate shapefile opened during the Input and Output prep
    Public m_edgeShp As IFeatureClass
    Public m_junctShp As IFeatureClass
    Public pWorkspaceI As IWorkspace
    'attribute dictionaries

    Public m_FinalJoin As IFeatureClass 'the last join feature class created during Output prep
    Public joinname As String 'table name of joining tblModelResults to TransRef edges


    Public m_prjAttr As PrjAttribute

    'global model
    Public advCheck As Boolean 'input preps signaling user can add additional projects in frmMapped
    Public pseudothin As Boolean 'input preps signaling whether to thin pseudonodes
    Public FCthin As Boolean 'input preps signaling whether to thin by functional class
    Public m_FC As String 'the functional class to thin by
    Public activethin As Boolean 'input preps signaling whether to thin by active link field
    Public m_Offset As Long 'the junction node offset to use
    Public FTthin As Boolean

    Public addweave As Boolean
    Public orgstr As String
    Public orgstr2 As String
    Public replstr As String
    Public replstr2 As String
    Public SplitM As Boolean
    Public spID As Long



    'selection-related global variables
    Public selectedprj As PrjSelected 'used by frmMapped to get all the projects to add to Model Run Scenario
    Public prjselectedCol As New Collection 'used by frmMapped to get all the projects to add to Model Run Scenario
    Public evtselectedCol As New Collection 'used by frmMapped to get all the projects with events to add to Model Run Scenario
    Public multiselected As New Collection 'used to save all the selectedprj objects from frmUncoded
    Public allselected() As PrjSelected 'used to store all projects listed in frmUncoded
    Public SelectAll As Boolean 'tell if user selected all projects in frmUncoded
    Public Selectcount As Long 'number of mulitselected projects from frmUncoded

    Public edgeshpCnt As Long
    Public g_ExportShapefiles As Boolean = False
    Public Class modelexport 'used in opening crystal reports in Output prep
        Public textstring As String
        Public position As Long
        Public rowvalue As String
        Public PosFrom As Integer
        Public PosTo As Integer
    End Class

    'edit-related global variables
    Public bigEid As Long 'stores the last edge id
    Public bigJid As Long 'stores the last junction id
    Public bigRid As Long 'stores the last projectroute id
    Public m_prjOID As Long 'used to store the project objectID currently performing edits to
    Public openCount As Long '1 if bigEid, bigJid, and bigRid have been initialized by scanning existing data
    Public edgeArray() As Long 'stores the edges to create project route
    Public junctionArray() As Long 'stores the junction to create project route
    Public edgesInroute As Integer 'number of edges needed to make route and then used to redim edgearray


    'Edit Case 4 globals
    Public case4 As Boolean 'used to flag that a junction will be splitting an edge
    Public deletedgeID As Long 'the psrcedgeid of the split edge into two new edges
    Public newedgeID1 As Long 'the first new edge psrcid
    Public newedgeID2 As Long 'the second new edge psrcid



    Public Function addLayerToTOC(ByVal strLayerName As String, ByVal fRefresh As Boolean) As Integer
        'adds layer name strLayer to TOC if is in geodb
        'refreshes the user view if fRefresh is TRUE
        'returns:
        '   1 if successful
        '   0 if layer does not exist
        '   -1 if error
        'derived from ESRI sample in E:\programs\ArcDocLib\ArcGIS_Desktop\DesktopDevGD_AppC.pdf

        'don't need the shapefile stuff
        '  Dim pWorkspaceFactory As IWorkspaceFactory
        '   pWorkspaceFactory = New ShapefileWorkspaceFactory
        '  Dim pWorkSpace As IFeatureWorkspace
        '   pWorkSpace = pWorkspaceFactory.OpenFromFile("C:\Source\", 0)

        ' we'll use the global WS variable
        On Error GoTo eh

        addLayerToTOC = 0

        Dim pClass As IFeatureClass
        pClass = g_FWS.OpenFeatureClass(strLayerName)
        Dim pLayer As IFeatureLayer
        pLayer = New FeatureLayer
        pLayer.FeatureClass = pClass
        pLayer.Name = pClass.AliasName



        pActiveView = m_Doc.ActiveView

        m_Map = pActiveView.FocusMap



        m_Map.AddLayer(pLayer)

        If fRefresh Then m_Doc.ActiveView.PartialRefresh(esriViewDrawPhase.esriViewGeography, pLayer, Nothing)

        addLayerToTOC = 1
        Exit Function

eh:
        'calling program must trap error!
        addLayerToTOC = -1
    End Function

    Public Sub init(ByVal pApp As IApplication, ByVal schemaName As String)
        'gives the GlobalMod handle to ArcMap
        'on error GoTo eh
        ' [072106] pan: temp hardcode to max transrefjunction + 1200 for debug
        '  LargestWJunct = 180623
        g_schema = schemaName
        SplitM = False

        'initialize global state variables [jaf 091206: editspecific]
        m_fEditMode = False
        m_fFatalError = False
        m_fNetChanged = False

        '[091806] jaf: moved task-specific items to their respective thisDocument.click handlers
        '              (verbose logging, transit built, edge/junct retain, open log file)

        Dim schema As String
        'schema = "sde.DBO."
        schema = schemaName

        If MsgBox("Are your layers in SDE?", vbYesNo, "Layer Source") = vbYes Then strLayerPrefix = "SDE" Else strLayerPrefix = ""

        If strLayerPrefix = "SDE" Then
            'MsgBox ("m_layers prefixed by sde.PSRC")
            m_layers(0) = schema & "TransRefEdges"
            m_layers(1) = schema & "TransRefJunctions"
            m_layers(2) = schema & "modeAttributes"
            m_layers(3) = schema & "tblLineProjects"
            m_layers(4) = schema & "evtLineProjectOutcomes"
            m_layers(5) = schema & "evtLineProjectOutcomes Events"
            m_layers(6) = schema & "evtPointProjectOutcomes"
            m_layers(7) = schema & "evtPointProjectOutcomes Events"
            m_layers(8) = schema & "ProjectRoutes"
            m_layers(9) = schema & "modeTolls"
            m_layers(10) = schema & "tblLineTolls"
            m_layers(11) = schema & "tblLineTolls Events"
            m_layers(12) = schema & "TurnMovements"
            m_layers(13) = schema & "TransitLines"
            m_layers(14) = schema & "tblTransitSegments"
            m_layers(15) = schema & "tblProjectsInScenarios"
            m_layers(16) = schema & "tblModelScenario"
            m_layers(17) = schema & "tblModelResults"
            m_layers(18) = schema & "tblModelResultstemp"
            m_layers(19) = schema & "tblModelResultsTR"
            m_layers(20) = schema & "Emme2edges"
            m_layers(21) = schema & "ScenarioJunct"
            m_layers(22) = schema & "ScenarioEdge"
            m_layers(23) = schema & "TransitPoints"
            m_layers(24) = schema & "tblScenarioProjects"
            m_layers(25) = schema & "vehiclecounts"
        Else
            'MsgBox "Not SDE"
            m_layers(0) = "TransRefEdges"
            m_layers(1) = "TransRefJunctions"
            m_layers(2) = "modeAttributes"
            m_layers(3) = "tblLineProjects"
            m_layers(4) = "evtLineProjectOutcomes"
            m_layers(5) = "evtLineProjectOutcomes Events"
            m_layers(6) = "evtPointProjectOutcomes"
            m_layers(7) = "evtPointProjectOutcomes Events"
            m_layers(8) = "ProjectRoutes"
            m_layers(9) = "modeTolls"
            m_layers(10) = "tblLineTolls"
            m_layers(11) = "tblLineTolls Events"
            m_layers(12) = "TurnMovements"
            m_layers(13) = "TransitLines"
            m_layers(14) = "tblTransitSegments"
            m_layers(15) = "tblProjectsInScenarios"
            m_layers(16) = "tblModelScenario"
            m_layers(17) = "tblModelResults"
            m_layers(18) = "tblModelResultstemp"
            m_layers(19) = "tblModelResultsTR"
            m_layers(20) = "Emme2edges"
            m_layers(21) = "ScenarioJunct"
            m_layers(22) = "ScenarioEdge"
            m_layers(23) = "TransitPoints"
            m_layers(24) = "tblScenarioProjects"
        End If

        Exit Sub

eh:

        CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.init")
        MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.init")

    End Sub
Public Sub addAttributes(tbl As Long, pOID As Long, rowcount As Long, fixedrow As Long)
 'used to assign attibute values from frmPrjattributes
 'may change so just popup the table to edit- not using as of 9/25/2003
  On Error GoTo eh
 
  Dim pWS As IWorkspace
        pWS = get_Workspace()
  Dim pWorkspaceEdit As IWorkspaceEdit
        pWorkspaceEdit = pWS
   
  pWorkspaceEdit.StartEditing (False)
  pWorkspaceEdit.StartEditOperation
  
  Dim pTable As ITable
  If (tbl = 1) Then
            pTable = GlobalMod.get_TableClass("sde.PSRC.tblLineProjects")
  Else
            pTable = GlobalMod.get_TableClass("sde.PSRC.tblPointProjects")
  End If
  'MsgBox "here"
  Dim i As Long, j As Long
  Dim sFld As String
  Dim vVal As String
  Dim pRow As IRow
  Dim idx As Long
  Dim pFlds As IFields
  Dim pFld As IField
  
        pRow = pTable.GetRow(pOID)
        pFlds = pRow.Fields
  Dim inservice As String
  Dim pFld2 As String
  
  Dim pDate As Date
        Dim CellValues() As String
  'CellValues = frmPrjAttributes.CellValues
  Dim temp As String
  'For i = flxGrid1.FixedRows To flxGrid1.Rows - 1
   ' For j = 0 To flxGrid1.Cols - 1 '+++ there are 2 cols
  For i = fixedrow To rowcount - 1
            For j = 0 To 2 - 1 '+++ there are 2 cols
                'SEC 052010- May need to fix below 
                'If j = 0 Then sFld = CellValues(i - 1, j)

                'If j = 1 Then vVal = CellValues(i - 1, j)

            Next j
     If (LCase(sFld) = "inservicedate") Then
        pFld2 = sFld
        inservice = vVal
     End If
     idx = pFlds.FindField(sFld)
     If Not Len(CStr(vVal)) = 0 Then
                pFld = pFlds.Field(idx)
                If pFld.Type = esriFieldType.esriFieldTypeInteger Or pFld.Type = esriFieldType.esriFieldTypeSmallInteger Then
                    If IsNumeric(vVal) Then pRow.Value(idx) = CInt(vVal)
                ElseIf pFld.Type = esriFieldType.esriFieldTypeDouble Or pFld.Type = esriFieldType.esriFieldTypeSingle Then
                    If IsNumeric(vVal) Then pRow.Value(idx) = CDbl(vVal)
                ElseIf pFld.Type = esriFieldType.esriFieldTypeString Then
                    pRow.Value(idx) = CStr(vVal)
                ElseIf pFld.Type = esriFieldType.esriFieldTypeDate Then

                End If '+++ have not accounted for date fields, etc
     End If
   Next i
   pRow.Store
   idx = pFlds.FindField(pFld2)
   pDate = CDate(vVal)
   temp = "#" + vVal + "#"
   
        If (pFlds.Field(idx).Type = esriFieldType.esriFieldTypeDate) Then
            pRow.Value(idx) = pDate
        Else
            pRow.Value(idx) = vVal
        End If
   pRow.Store
   pWorkspaceEdit.StopEditOperation
        pWorkspaceEdit.StopEditing(True)
   Exit Sub
eh:
   'pRow.Store  jaf: let's not do this here 'cause it may be the cause...
        MsgBox("Error #" & Err.Number & ", " & Err.Description, vbExclamation, "GlobalMod.addAttributes")
        GlobalMod.CloseLogFile("GlobalMod.addAttributes: closed log file due to " & Err.Description)
   Err.Clear
   
   If pWorkspaceEdit.IsBeingEdited Then
      pWorkspaceEdit.AbortEditOperation
            pWorkspaceEdit.StopEditing(False)
   End If
End Sub
Public Sub get_UpdatePrjRteID()
   'update the project route id in global type m_prjAttr...
   '...to match the project ID already in m_prjAttr.PrjID
   
   'jaf--not sure how we are certain that m_prjAttr.PrjID has the project we're looking for...
   
   'called by:
   '  writeProjOutcomesData
   On Error GoTo eh
   Dim pFeatLayer As IFeatureLayer
        pFeatLayer = get_FeatureLayer("sde.PSRC.ProjectRoutes")
   Dim pFeatSelect As IFeatureSelection
   Dim pSelectSet As ISelectionSet
   Dim pFC As IFeatureCursor
   Dim pFeature As IFeature
        pFeatSelect = pFeatLayer
    
        pFC = pFeatLayer.Search(Nothing, False)
        pFeature = pFC.NextFeature
    
   Dim index As Long
    
   Dim projid As String
   'MsgBox m_prjAttr.PrjId
   Do Until pFeature Is Nothing
      index = pFeatLayer.FeatureClass.FindField("projID")
      projid = pFeature.value(index)
      If (projid = m_prjAttr.PrjId) Then
         index = pFeatLayer.FeatureClass.FindField("projRteID")
         m_prjAttr.PrjRteID = pFeature.value(index)
         Exit Do
      End If
            pFeature = pFC.NextFeature
   Loop
   Exit Sub
eh:
   'pRow.Store  jaf: let's not do this here 'cause it may be the cause...
        MsgBox("Error #" & Err.Number & ", " & Err.Description, vbExclamation, "GlobalMod.get_UpdatePrjRteID")
        GlobalMod.CloseLogFile("GlobalMod.get_UpdatePrjRteID: closed log file due to " & Err.Description)
   Err.Clear
End Sub
Public Sub uninit()
        m_App = Nothing
        m_Map = Nothing
        m_TableTIP = Nothing
        m_TableMTP = Nothing
    SelectAll = False
        m_JunctSSet = Nothing
        m_EdgeSSet = Nothing
        m_FinalJoin = Nothing

    '[jaf 091206: editspecific]
        m_Editor = Nothing
    case4 = False

    'close the log file
        CloseLogFile("Terminated normally by call to uninit")
End Sub

Public Function get_App() As IApplication
  'on error GoTo eh
  'allows forms to get access to ArcMap if needed for quick internal routine
        get_App = m_App

  Exit Function

eh:

        CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.get_App")
        MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.get_App")

End Function


Public Function setSDEWorkspace() As IFeatureWorkspace
        Dim pPropset As ESRI.ArcGIS.esriSystem.IPropertySet
        pPropset = New ESRI.ArcGIS.esriSystem.PropertySet
    With pPropset
            .SetProperty("Server", "ELM")
            .SetProperty("Instance", "5151")
            .SetProperty("user", "PSRC")
            .SetProperty("password", "mpo")
            .SetProperty("Database", "sde")
            .SetProperty("version", "SDE.DEFAULT")

    '[111005] jaf: commented out above and  below instead
'    .SetProperty "Server", "sde"
'    .SetProperty "Instance", "5151"
'    .SetProperty "user", "CShong" '"PSRC"
'    .SetProperty "password", "devel902" '"mpo"
'    .SetProperty "Database", "sde"
'   .SetProperty "version", "SDE.DEFAULT"
    End With

  Dim pFactSDE As IWorkspaceFactory
        pFactSDE = New SdeWorkspaceFactory
        setSDEWorkspace = pFactSDE.Open(pPropset, 0)

End Function

Public Function getFeatureClass(strFeatureclass As String) As IFeatureClass

    Dim pFeatClass As IFeatureClass
    Dim i As Long

        'If g_FWS Is Nothing Then g_FWS = getPGDws
     pFeatClass = g_FWS.OpenFeatureClass(strFeatureclass)

    If pFeatClass Is Nothing Then
            MsgBox("GlobalMod.getFeatureClass: " & strFeatureclass & " not found")
        Exit Function
    End If

     getFeatureClass = pFeatClass
End Function

    Public Function get_FeatureLayer(ByVal featClsname As String) As IFeatureLayer
        'returns the feature class stated in featClsname
        'on error GoTo eh




        Dim pEnumLy As IEnumLayer

        If m_Map Is Nothing Then
            SetMainVariables()
        End If
        m_Map = pActiveView.FocusMap
        pEnumLy = m_Map.Layers

        Dim pLy As ILayer, pFLy As IFeatureLayer
        pLy = pEnumLy.Next
        Do Until pLy Is Nothing

            If pLy.Valid Then
                If TypeOf pLy Is IFeatureLayer Then
                    pFLy = pLy

                    'If UCase(pFLy.FeatureClass.AliasName) = UCase(featClsname) Then
                    If UCase(pFLy.Name) = UCase(featClsname) Then
                        get_FeatureLayer = pFLy
                        Exit Do

                    End If
                End If
            End If
            pLy = pEnumLy.Next
        Loop

        If Not get_FeatureLayer Is Nothing Then Exit Function

        Dim pFeatLayer As IFeatureLayer
        pFeatLayer = New FeatureLayer
        Dim pfeatcls As IFeatureClass
        pfeatcls = getFeatureClass(featClsname)
        pFeatLayer.FeatureClass = pfeatcls

        get_FeatureLayer = pFeatLayer
        If pFeatLayer Is Nothing Then
            MsgBox("Error: did not find the feature class " + featClsname)
        End If

        Exit Function

eh:
        MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.get_FeatureLayer")

    End Function
    

Public Sub SetMainVariables()

        m_Doc = m_App.Document
        m_Map = m_Doc.FocusMap
        pActiveView = m_Doc.FocusMap
        '  pSp = m_Map.SpatialReference
  loaded = True

    End Sub



'eventually openSDEWorksapce and openTIPorMTP will be combined
Public Sub openTIPorMTP(dbtype As Long, both As Boolean, filepath As String)
   'used currently to access external databases in MSAccess
  'on error GoTo eh
        WriteLogLine("called openTIPorMTP")
        WriteLogLine("-------------------")

   'testing without sde
    'MsgBox "Opening database: " + filepath
    Dim pFile As String, pFile2 As String
    pFile = filepath
    GlobalMod.m_tablefile = filepath

    Dim pWorkspaceFact As IWorkspaceFactory
        pWorkspaceFact = New AccessWorkspaceFactory 'workspace to get to access tables
    Dim pPropset As IPropertySet
        pPropset = New PropertySet
        pPropset.SetProperty("DATABASE", pFile)

    Dim pWorkspace As IWorkspace
        pWorkspace = pWorkspaceFact.Open(pPropset, 0)
        ' pWorkspace = openSDEWorkspace(server, instance, user, password)

    Dim pFeatWorkspace As IFeatureWorkspace
        pFeatWorkspace = pWorkspace

    Dim pTable As ITable
    Dim pTableSelect As ISelectionSet
    Dim pTableCursor As ICursor
    Dim pQueryFilter As IQueryFilter

    Dim pEnumDataset2 As IEnumDatasetName
        pEnumDataset2 = pWorkspace.DatasetNames(esriDatasetType.esriDTAny)

    Dim pDataset2 As IDatasetName
        pDataset2 = pEnumDataset2.Next
        'need to  up so user can navigate to this
    Do Until pDataset2 Is Nothing
        'If (dbtype = 0) Then
        'If (pDataset2.name = "Table1") Then
            ' pTable = pFeatWorkspace.OpenTable(pDataset2.name)
            ' m_TableTIP = pFeatWorkspace.OpenTable(pDataset2.name)
            'Exit Do
        'End If
        'Else
         'If (pDataset2.name = "tblProjects") Then
            ' pTable = pFeatWorkspace.OpenTable(pDataset2.name)
            ' m_TableMTP = pFeatWorkspace.OpenTable(pDataset2.name)
            'Exit Do
         'End If
        'End If

        'now give user control of selecting the table
        'pan I cant find listdataset
        'frmPath.Listdataset.AddItem pDataset2.name
         pDataset2 = pEnumDataset2.Next
        Loop
        'SEC 052010- may need to fix below
        'If frmPath.Listdataset.ListCount = 0 Then
        'ox("No Tables were found in this database")
        'End If
        pWorkspace = Nothing
        pEnumDataset2 = Nothing
        pDataset2 = Nothing
        Exit Sub
eh:

        CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.openTIPorMTP")
        MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.openTIPorMTP")

        'MsgBox Err.Description, vbInformation, "openTIPorMTP"
        pWorkspace = Nothing
        pEnumDataset2 = Nothing
        pDataset2 = Nothing

  End Sub

Public Sub openDBTable(tablename As String, dbtype As Long, both As Boolean)
  'opens the table selected by the user from frmPath
  'on error GoTo eh
        WriteLogLine("called openDBTable for tablename=" & tablename)
        WriteLogLine("----------------------------------------------")

'  Dim pathname As String
'  pathname = "c:\MODULEOUTPUT2.txt"
'  Open pathname For Output As #2
'  Print #2, "Maintenance openDBTable"

  Dim tempIndex As Long, i As Long

   Dim pStatusBar As IStatusBar
   Dim pProgbar As IStepProgressor

        pStatusBar = m_App.StatusBar
        pProgbar = pStatusBar.ProgressBar

        WriteLogLine(" up status bar")
    Dim pFile As String
    'MsgBox GlobalMod.m_tablefile
    pFile = GlobalMod.m_tablefile
        WriteLogLine(pFile)
    Dim pWorkspaceFact As IWorkspaceFactory
        pWorkspaceFact = New AccessWorkspaceFactory 'workspace to get to access tables
    Dim pPropset As IPropertySet
        pPropset = New PropertySet
        pPropset.SetProperty("DATABASE", pFile)

    Dim pWorkspace As IWorkspace
        pWorkspace = pWorkspaceFact.Open(pPropset, 0)
        ' pWorkspace = openSDEWorkspace(server, instance, user, password)
        WriteLogLine("Opened workspace ")
    Dim pFeatWorkspace As IFeatureWorkspace
        pFeatWorkspace = pWorkspace

    Dim pTable As ITable
    Dim pTableSelect As ISelectionSet
    Dim pTableCursor As ICursor
    Dim pQueryFilter As IQueryFilter

    Dim pEnumDataset2 As IEnumDatasetName
        pEnumDataset2 = pWorkspace.DatasetNames(esriDatasetType.esriDTAny)

    Dim pDataset As IDatasetName
        pDataset = pEnumDataset2.Next

        Dim MyfrmUnMap_model As New frmUnMap_model
        Dim MyfrmMapped As New frmMapped
    'need to  up so user can navigate to this
    Do Until pDataset Is Nothing
       If (dbtype = 0) Then
        If (pDataset.Name = tablename) Then

                    pTable = pFeatWorkspace.OpenTable(pDataset.Name)
                    m_TableTIP = pFeatWorkspace.OpenTable(pDataset.Name)
                    WriteLogLine("found MTP")
            Exit Do
        End If
       Else
         If (pDataset.Name = tablename) Then

                    pTable = pFeatWorkspace.OpenTable(pDataset.Name)
                    m_TableMTP = pFeatWorkspace.OpenTable(pDataset.Name)
                    WriteLogLine("found TIP")
            Exit Do
         End If
       End If
        'now give user control of selecting the table
            pDataset = pEnumDataset2.Next
    Loop

  Dim pFeatLayer As IFeatureLayer
        pFeatLayer = get_FeatureLayer(m_layers(8)) 'Project Routes

  Dim pTableCol As ITableCollection
        pTableCol = m_Map

  Dim fIndex As Long
  Dim idver As String, idverquery As String

  'If (dbtype = 0) Then
'    fIndex = pFeatLayer.FeatureClass.Fields.FindField("projID_Ver")
  'End If
  'If (dbtype = 1) Then
    '[100207] hyu: should be projRteID not projID_Ver
   fIndex = pFeatLayer.FeatureClass.Fields.FindField("projRteID") '"projID_Ver")
  'End If
  If (fIndex = -1) Then
            MsgBox("ProjectRoutes is missing the field projRteID") 'projID_Ver"
            'Unload(frmPath)
            g_frmChkProjects.ActiveForm.Show()
            'frmChkProjects.Show(vbModal)
   Exit Sub
  End If
        WriteLogLine("got project routes projRteID") 'projID_Ver"
   'check the projectroutes for the correct fields

    Dim length As Long, pos As Long
    Dim pField As IField
    Dim pFields As IFields
    Dim pFieldEdit As IFieldEdit
    Dim pFC As IFeatureCursor
    Dim pFeat As IFeature

    Dim pWSedit As IWorkspaceEdit

    Dim pWS As IWorkspace
        pWS = get_Workspace()

    Dim idvalue As String, temp As String, prstring As String
    temp = "-"
    tempIndex = pFeatLayer.FeatureClass.FindField("projID")
    If (tempIndex = -1) Then
            pWSedit.StartEditing(False)
    pWSedit.StartEditOperation
        MsgBox ("projectRoutes is missing the projID field!")
            pFields = pFeatLayer.FeatureClass.Fields
            pField = New Field
            pFieldEdit = pField
            pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString
            pFieldEdit.AliasName_2 = "projID"
            pFieldEdit.Name_2 = pField.AliasName
            pFeatLayer.FeatureClass.AddField(pField)

            pFC = pFeatLayer.Search(Nothing, False)
            pFeat = pFC.NextFeature
        tempIndex = pFeatLayer.FeatureClass.FindField("projID")
        Do Until pFeat Is Nothing
                If Not pFeat.Value(fIndex) Is Nothing And pFeat.Value(fIndex) = False Then
                    If fVerboseLog Then WriteLogLine(pFeat.Value(fIndex))
                    idvalue = pFeat.Value(fIndex)
                    length = Len(idvalue)
                    If fVerboseLog Then WriteLogLine(length)
                    If (length > 0) Then
                        pos = InStrRev(1, idvalue, temp, vbTextCompare)
                        If (pos <> 0) Then
                            pos = pos - 1
                            prstring = Left(idvalue, pos)
                            If fVerboseLog Then WriteLogLine(prstring)

                            pFeat.Value(tempIndex) = prstring
                        Else
                            MsgBox("This project route is blank!" + CStr(pFeat.OID))
                        End If
                        pFeat.Store()
                    Else
                        MsgBox("This project route is blank!" + CStr(pFeat.OID))
                    End If
                    pFeat = pFC.NextFeature

                Else
                    MsgBox("This project route has Null values!" + CStr(pFeat.OID))
                    pFeat = pFC.NextFeature
                End If

        Loop

        pWSedit.StopEditOperation
            pWSedit.StopEditing(True)
        '[090506] pan:  uncommented load frmchkprojects, etc
            'Unload frmPath
            'frmChkProjects.
            g_frmChkProjects.ActiveForm.Show()
      Exit Sub
    End If
        WriteLogLine("have all the projid")
    Dim pidIndex As Long, verIndex As Long, mapIndex As Long
    Dim pidValue As String, verValue As String, pidVerValue As String
    Dim mapValue As String

    'NOTE: need to change if naming convention changes in databases!
    pidIndex = pTable.Fields.FindField("ProjNo")
    verIndex = pTable.Fields.FindField("intVersionCount")
    mapIndex = pTable.Fields.FindField("Mappable")
    If (pidIndex = -1 Or verIndex = -1 Or mapIndex = -1) Then
            MsgBox("This table is missing either the ProjNo field, the intVersionCount field, or the Mappable field!")
            'Unload frmPath
            'Load frmChkProjects
            g_frmChkProjects.ActiveForm.Show()
        Exit Sub
    End If
    'check the projectroutes for the correct fields
        WriteLogLine("have all the table fields")
   Dim inserviceStr As String

  Dim pFields2 As IFields
        pFields2 = pTable.Fields
  For i = 0 To pFields2.FieldCount - 1
    If (LCase(pFields2.field(i).Name) = "inservicedate") Then
        inserviceStr = pFields2.field(i).Name
        Exit For
     End If
  Next i
  If (inserviceStr = "") Then
            MsgBox("Did not find the InServiceDate field in the table")
    Exit Sub
  End If
        WriteLogLine("got table inservice")
   'get all projects in Sercive by data specified in frmModelQuery
        pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "inServiceDate <= #" + CStr(inserviceDate) + "#"
        pTableCursor = pTable.Search(pQueryFilter, False)
    'MsgBox pTable.rowcount(pQueryFilter)
    pProgbar.position = 0
        pStatusBar.ShowProgressBar("Loading...", 0, pTable.RowCount(pQueryFilter), 1, True)

    Dim pRow As IRow
        pRow = pTableCursor.NextRow

  Dim pFeatSelection As IFeatureSelection
        pFeatSelection = pFeatLayer 'QI
  Dim pSelectSet As ISelectionSet
  Dim pFeatCursor As IFeatureCursor
  Dim pFeature As IFeature

    Do Until pRow Is Nothing
    'For i = 0 To 100
        pidValue = pRow.value(pidIndex)
        verValue = pRow.value(verIndex)
        pidVerValue = pidValue + "-" + verValue

        If (pRow.value(mapIndex) = 1) Then
          'this is mappable, now check if in projectRoutes feature class
          idver = pidVerValue
                pQueryFilter = New QueryFilter
          If (dbtype = 0) Then 'its tip
           pQueryFilter.SubFields = "projID_Ver"
           idverquery = "projID_Ver = '"
           idverquery = idverquery + idver + "'"
           pQueryFilter.WhereClause = idverquery
          End If
          If (dbtype = 1) Then 'its mtp
           pQueryFilter.SubFields = "projID_Ver"
           idverquery = "projID_Ver = '"
           idverquery = idverquery + idver + "'"
           pQueryFilter.WhereClause = idverquery
         End If
                pFeatSelection.SelectFeatures(pQueryFilter, esriSelectionResultEnum.esriSelectionResultNew, False)
                pSelectSet = pFeatSelection.SelectionSet
         If (pSelectSet.count = 0) Then
             'not coded yet or its a new version
                    pQueryFilter = New QueryFilter
             If (dbtype = 0) Then
              pQueryFilter.SubFields = "projID"
              idverquery = "projID = '"
              idverquery = idverquery + pidValue + "'"
              pQueryFilter.WhereClause = idverquery
             End If
             If (dbtype = 1) Then
              pQueryFilter.SubFields = "projID"
              idverquery = "projID = '"
              idverquery = idverquery + pidValue + "'"
              pQueryFilter.WhereClause = idverquery
            End If
                    pFeatSelection.SelectFeatures(pQueryFilter, esriSelectionResultEnum.esriSelectionResultNew, False)
                    pSelectSet = pFeatSelection.SelectionSet
                    'MsgBox pSelectSet.count
                  

            If (pSelectSet.count = 0) Then
                If (dbtype = 0) Then
                 If (both = True) Then
                                MyfrmUnMap_model.lstNewmtp.Items.Add(pidVerValue)

                  'End If
                 Else
                                MyfrmUnMap_model.ListBox1.Items.Add(pidVerValue)
                                'frmUnMap_model.ListBox1.AddItem(pidVerValue)
                  'End If
                 End If
                Else
                            If (both = True) Then
                                MyfrmUnMap_model.lstNewmtp.Items.Add(pidVerValue)
                                'frmUnMap_model.lstNewmtp.AddItem(pidVerValue)
                                'End If
                            Else
                                MyfrmUnMap_model.ListBox1.Items.Add(pidVerValue)
                                'frmUnMap_model.ListBox1.AddItem(pidVerValue)
                                'End If
                            End If
                End If
            Else
              'now check version number
              If (dbtype = 0) Then 'its tip
                            If (both = True) Then
                                MyfrmUnMap_model.lstNewtip.Items.Add(pidVerValue)
                                'frmUnMap_model.lstNewtip.AddItem(pidVerValue)
                                'End If
                            Else
                                MyfrmUnMap_model.ListBox1.Items.Add(pidVerValue)
                                'frmUnMap_model.ListBox1.AddItem(pidVerValue)
                                'End If
                            End If
              Else 'its mtp
                            If (both = True) Then
                                MyfrmUnMap_model.lstNewmtp.Items.Add(pidVerValue)
                                'frmUnMap_model.lstNewmtp.AddItem(pidVerValue)
                                'End If
                            Else
                                MyfrmUnMap_model.ListBox1.Items.Add(pidVerValue)
                                'frmUnMap_model.ListBox1.AddItem(pidVerValue)
                                'End If
                            End If
              End If
            End If
         Else 'selection count is greater than 0 so can write to frmMapped for user selection
            If (dbtype = 0) Then 'its tip
                        If (both = True) Then

                            MyfrmUnMap_model.lstNewmtp.Items.Add(pidVerValue)
                            'frmMapped.lstMtip.AddItem(pidVerValue)
                        Else
                            MyfrmUnMap_model.ListBox1.Items.Add(pidVerValue)
                            'frmMapped.ListBoxM.AddItem(pidVerValue)
                        End If
            Else 'it mtp
                        If (both = True) Then

                            MyfrmMapped.lstMmtp.Items.Add(pidVerValue)
                            'frmMapped.lstMmtp.AddItem(pidVerValue)
                        Else
                            MyfrmMapped.lstBoxM.Items.Add(pidVerValue)
                            'frmMapped.ListBoxM.AddItem(pidVerValue)
                        End If
            End If
         End If
       Else 'its unmappable
            If (dbtype = 0) Then 'its tip
                    If (both = True) Then
                        MyfrmUnMap_model.lstUntip.Items.Add(pidVerValue)
                        'frmUnMap_model.lstUntip.AddItem(pidVerValue)
                        ' End If
                    Else
                        MyfrmUnMap_model.ListBox2.Items.Add(pidVerValue)
                        'frmUnMap_model.ListBox2.AddItem(pidVerValue)
                        'End If
                    End If
            Else
                    If (both = True) Then
                        MyfrmUnMap_model.lstUnmtp.Items.Add(pidVerValue)
                        'frmUnMap_model.lstUnmtp.AddItem(pidVerValue)
                        'End If
                    Else
                        MyfrmUnMap_model.ListBox2.Items.Add(pidVerValue)
                        'frmUnMap_model.ListBox2.AddItem(pidVerValue)
                        'End If
                    End If
            End If
        End If

            pRow = pTableCursor.NextRow
        pStatusBar.StepProgressBar
    'Next i
   Loop


   Dim tblLine As ITable
        tblLine = get_TableClass(m_layers(3))
   'Dim tblPoint As ITable- no longer use
   ' tblPoint = get_TableClass("tblPointProjects", pWs)
   Dim evtLine As ITable
        evtLine = get_TableClass(m_layers(4))
   Dim evtPoint As ITable
        evtPoint = get_TableClass(m_layers(6))

   Dim pRowevt As IRow
        pTableCursor = Nothing
        pRow = Nothing
   Dim index As Long, dbindex As Long
   Dim tempDate As Date
   Dim pTCevt As ICursor

  'need to check projDBS also
   pidIndex = tblLine.Fields.FindField("projID_Ver")
        pQueryFilter = New QueryFilter
   pQueryFilter.WhereClause = "withEvents = 1"
   'pSelectSet.Search pQueryFilter, False, pFeatureCursor
        pTableCursor = tblLine.Search(pQueryFilter, False)
   If (tblLine.rowcount(pQueryFilter) > 0) Then

            pRow = pTableCursor.NextRow
    pidValue = pRow.value(pidIndex)

     Do Until pRow Is Nothing
                pQueryFilter = New QueryFilter

        If (evtLine.rowcount(pQueryFilter) > 0) Then
            pQueryFilter.WhereClause = "ProjID_Ver = '" + CStr(pidValue) + "'"
                    pTCevt = evtLine.Search(pQueryFilter, False)
            index = evtLine.FindField("inServiceDate")
            dbindex = evtLine.FindField("projDBS")
            'MsgBox "line"
        ElseIf (evtPoint.rowcount(pQueryFilter) > 0) Then
            pQueryFilter.WhereClause = "projID_Ver = '" + CStr(pidValue) + "'"
                    pTCevt = evtPoint.Search(pQueryFilter, False)
            index = evtPoint.FindField("inServiceDate")
            dbindex = evtPoint.FindField("projDBS")
            'MsgBox "point"
        End If
                pRowevt = pTCevt.NextRow
        tempDate = pRowevt.value(index)
        If (tempDate <= inserviceDate) Then
         If (pRowevt.value(dbindex) = "TIP" And dbtype = 0) Then 'its tip
                        If (both = True) Then

                            MyfrmMapped.ListBoxtipE.Items.Add(pidValue)
                            'frmMapped.ListBoxtipE.AddItem(pidValue)
                        Else
                            MyfrmMapped.ListBoxM.Items.Add(pidValue)
                            'frmMapped.ListBoxME.AddItem(pidValue)
                        End If
                    ElseIf (pRowevt.value(dbindex) = "MTP" And dbtype = 1) Then 'it mtp
                        If (both = True) Then
                            MyfrmMapped.ListBoxmtpE.Items.Add(pidValue)
                            'frmMapped.ListBoxmtpE.AddItem(pidValue)
                        Else
                            MyfrmMapped.ListBoxM.Items.Add(pidValue)
                            'frmMapped.ListBoxME.AddItem(pidValue)
                        End If
                    End If
                End If

                pRow = pTableCursor.NextRow
            Loop

        End If

        pQueryFilter.WhereClause = "withEvents = '1'"
        'pidIndex = tblPoint.Fields.FindField("projID_Ver")
        ' pTableCursor = tblPoint.Search(pQueryFilter, False)
        'If (tblPoint.rowcount(pQueryFilter) > 0) Then
        'deleted code 01-19-05
        'End If

        pStatusBar.HideProgressBar()
        pStatusBar.HideProgressBar()
        pWorkspace = Nothing
        pFeatWorkspace = Nothing
        pEnumDataset2 = Nothing
        pDataset = Nothing
        pTable = Nothing
        pFeature = Nothing
        pRow = Nothing
        WriteLogLine("FINISHED openDBTable")
        'Close #2

        Exit Sub

eh:

        CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.openDBTable")
        MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.openDBTable")

        'MsgBox Err.Description, vbInformation, "openDBTable"
        'Close #2
        pStatusBar.HideProgressBar()
        pWorkspace = Nothing
        pFeatWorkspace = Nothing
        pEnumDataset2 = Nothing
        pDataset = Nothing
        pTable = Nothing
        pFeature = Nothing
        pRow = Nothing
    End Sub

    Public Sub findProspectiveProjects(ByVal dbtype As Long, ByVal both As Boolean)
        'on error GoTo eh

        '[061907] jaf: created this sub as temp replacement for ADO_openTIPorMTP

        WriteLogLine("====================================================")
        WriteLogLine("findProspectiveProjects")
        WriteLogLine("====================================================")

        Dim pStatusBar As IStatusBar
        Dim pProgbar As IStepProgressor
        Dim pFeatLayer As IFeatureLayer
        Dim tempIndex As Long
        Dim length As Long
        Dim pos As Long
        Dim lngID As Long
        Dim MyfrmMapped As New frmMapped

        pStatusBar = m_App.StatusBar
        pProgbar = pStatusBar.ProgressBar
        '   pFeatLayer = get_FeatureLayer(m_layers(8)) 'ProjectRoutes
        '  tempIndex = pFeatLayer.FeatureClass.FindField("projID")
        '  If (tempIndex = -1) Then
        '        RemoveEvents
        '  End If

        Dim pTblPrj As ITable
        pTblPrj = get_TableClass(m_layers(3))

        Dim fIndex As Long
        Dim idver As String, idverquery As String
        Dim pWSedit As IWorkspaceEdit
        ' pws = get_Workspace
        pWSedit = g_FWS

        Dim pidIndex As Long, verIndex As Long, mapIndex As Long
        Dim pidValue As String, verValue As String, pidVerValue As String
        Dim mapValue As String
        Dim pField As IField
        Dim pFields As IFields
        Dim pFieldEdit As IFieldEdit
        Dim pFC As IFeatureCursor
        Dim pFeat As IFeature
        Dim pWS As IWorkspace
        Dim idvalue As String, temp As String, prstring As String

        '  temp = "-"
        '  tempIndex = pFeatLayer.FeatureClass.FindField("projID")

        Dim pQueryFilter As IQueryFilter
        Dim pFeatSelection As IFeatureSelection
        Dim pSelectSet As ISelectionSet
        Dim pRow As IRow
        Dim i As Long
        Dim value As Object
        Dim str As String
        Dim lIndex As Integer

        '[061907] jaf:  up query against ProjectRoutes (to which pFeatLayer points now)
        MyfrmMapped.lstMtip.Items.Clear()
        MyfrmMapped.lstMmtp.Items.Clear()
        MyfrmMapped.lstBoxM.Items.Clear()


        pQueryFilter = New QueryFilter

        Dim pCs As ICursor
        Dim pAtt As IRow
        Dim yr As Long
        Dim sProjID As String, lProjRteID As Long

        '[072307] hyu:  query clause accordingly
        If both = False Then
            If dbtype = 0 Then
                pQueryFilter.WhereClause = "projDBS='TIP'"
            Else
                pQueryFilter.WhereClause = "projDBS='MTP'"
            End If
        End If
        yr = Val(Trim(g_frmModelQuery.txtYear.ToString))
        If pQueryFilter.WhereClause <> "" Then pQueryFilter.WhereClause = pQueryFilter.WhereClause + " AND "
        pQueryFilter.WhereClause = pQueryFilter.WhereClause + " InServiceDate<=" & yr & " AND OutServiceDate>" & yr

        pCs = pTblPrj.Search(pQueryFilter, False)
        pAtt = pCs.NextRow

        If Not (pAtt Is Nothing) Then
            'we have Projects to include

            'populate the form listbox with the Projects
            Do Until pAtt Is Nothing

                lIndex = pAtt.Fields.FindField("projRteID")
                If Not pAtt.Value(lIndex) Is Nothing Then lProjRteID = "Null projRteID!" Else lProjRteID = pAtt.Value(lIndex)
                '[080808] SEC: "projID_Ver is no longer in tblLineProjects. Commented out the following lines related to this attribute field.
                'lIndex = pAtt.Fields.FindField("projID_Ver")
                'If IsDBNull(pAtt.value(lIndex)) Then sProjID = "Null projID_Ver!" Else sProjID = pAtt.value(lIndex)

                If (dbtype = 0) Then 'tip
                    If (both = True) Then
                        MyfrmMapped.lstMtip.Items.Add(lProjRteID) 'pidVerValue
                        'frmMapped.lstMtip.Column(1, frmMapped.lstMtip.ListCount - 1) = sProjID
                    Else
                        MyfrmMapped.ListBoxM.Items.Add(lProjRteID) 'pidVerValue
                        'frmMapped.ListBoxM.Column(1, frmMapped.ListBoxM.ListCount - 1) = sProjID
                    End If
                Else 'mtp
                    If (both = True) Then
                        MyfrmMapped.lstMmtp.Items.Add(lProjRteID) ' pidVerValue
                        'frmMapped.lstMmtp.Column(1, frmMapped.lstMmtp.ListCount - 1) = sProjID
                    Else
                        MyfrmMapped.ListBoxM.Items.Add(lProjRteID) ' pidVerValue
                        'frmMapped.ListBoxM.Column(1, frmMapped.ListBoxM.ListCount - 1) = sProjID
                    End If
                End If

                pAtt = pCs.NextRow

            Loop
        Else
            'no ProjectRoutes
            WriteLogLine("No project routes met Where clause='" & pQueryFilter.WhereClause & "'")
        End If 'do we have ProjectRoutes

        g_frmMapped.ListBoxtipE.Items.Clear()
        g_frmMapped.ListBoxM.Items.Clear()
        g_frmMapped.ListBoxmtpE.Items.Clear()


        WriteLogLine("done ProjectRoutes, SKIPPING events in this code version")

        '[061907] jaf: skip events for now
        Exit Sub

        Dim tblLine As ITable
        Dim evtLine As ITable
        Dim evtPoint As ITable
        Dim pRowevt As IRow
        Dim pTableCursor As ICursor
        Dim index As Long, dbindex As Long
        Dim tempDate As Long
        Dim pTCevt As ICursor
        Dim count As Long, ecnt As Long

        tblLine = get_TableClass(m_layers(3))
        'no longer using tblPoint table- deleting code from this version
        'Dim tblPoint As ITable
        ' tblPoint = get_TableClass("tblPointProjects", pWs)
        WriteLogLine(vbCrLf & "Starting to get events")

        evtLine = get_TableClass(m_layers(4))
        evtPoint = get_TableClass(m_layers(6))
        pRow = Nothing

        'need to check projDBS also
        pidIndex = pFeatLayer.FeatureClass.FindField("projID_Ver")
        pQueryFilter = New QueryFilter
        'need to add code here for choice of filters
        pQueryFilter.WhereClause = "withEvents > 0"
        WriteLogLine("going to select events")

        pFC = Nothing
        pFeat = Nothing
        pFC = pFeatLayer.Search(pQueryFilter, True)

        'pSelectSet.Search pQueryFilter, False, pfeaturecursor
        ' pTableCursor = tblLine.Search(pQueryFilter, False)
        'If (tblLine.rowcount(pQueryFilter) > 0) Then 'changed to looking in ProjectRoutes for withEvents field
        If (pFeatLayer.FeatureClass.featurecount(pQueryFilter) > 0) Then
            ' pRow = pTableCursor.NextRow
            pFeat = pFC.NextFeature

            WriteLogLine("proceed to loop through tblline with events")

            Do Until pFeat Is Nothing
                pidValue = pFeat.value(pidIndex)
                pQueryFilter = New QueryFilter
                For ecnt = 0 To 1 'can have both point and line events
                    pQueryFilter.WhereClause = "ProjID_Ver = '" + CStr(pidValue) + "' And InServiceDate <= " + CStr(inserviceyear) + " AND OutServiceDate >" + CStr(inserviceyear)
                    If fVerboseLog Then WriteLogLine(pidValue)

                    If (ecnt = 0) Then
                        pTCevt = evtLine.Search(pQueryFilter, False)
                        'index = evtLine.FindField("InServiceDate")
                        count = evtLine.rowcount(pQueryFilter)
                        dbindex = evtLine.FindField("projDBS")
                    End If

                    If (ecnt = 1) Then
                        pTCevt = evtPoint.Search(pQueryFilter, False)
                        count = evtPoint.rowcount(pQueryFilter)
                        dbindex = evtPoint.FindField("projDBS")
                    End If

                    If fVerboseLog Then WriteLogLine(count)
                    If (count > 0) Then
                        pRowevt = pTCevt.NextRow
                        Do Until pRowevt Is Nothing
                            If fVerboseLog Then WriteLogLine(pRowevt.OID)
                            If (pRowevt.value(dbindex) = "TIP" And dbtype = 0) Then 'its tip
                                If (both = True) Then
                                    g_frmMapped.ListBoxtipE.Items.Add(pidValue)
                                Else
                                    g_frmMapped.ListBoxM.Items.Add(pidValue)
                                End If
                            ElseIf (pRowevt.value(dbindex) = "MTP" And dbtype = 1) Then 'it mtp
                                If (both = True) Then
                                    g_frmMapped.ListBoxmtpE.Items.Add(pidValue)
                                Else
                                    g_frmMapped.ListBoxM.Items.Add(pidValue)
                                End If
                            End If

                            If fVerboseLog Then WriteLogLine("wrote")
                            pRowevt = pTCevt.NextRow
                        Loop
                    End If
                Next ecnt
                pFeat = pFC.NextFeature
            Loop
        End If

        'pQueryFilter.WhereClause = "withEvents = 1"
        'pidIndex = tblPoint.Fields.FindField("projID_Ver")
        ' pTableCursor = tblPoint.Search(pQueryFilter, False)
        'If (tblPoint.rowcount(pQueryFilter) > 0) Then
        'deleted code 01-19-05

        'End If
        'pStatusBar.HideProgressBar

        pFeat = Nothing
        pWS = Nothing
        ' pFeatWorkspace = Nothing
        ' pEnumDataset2 = Nothing
        ' pDataset = Nothing
        ' pTable = Nothing
        ' pfeature = Nothing

        ' clean up

        'Debug.Print
        WriteLogLine("FINISHED findProspectiveProjects at " & Now())
        'Close #2
        'Close #1
        Exit Sub
eh:
        CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.findProspectiveProjects")
        MsgBox(Err.Description, vbInformation, "GlobalMod.findProspectiveProjects")
        'Close #2
        'Close #2
        'Close #1
        If pWSedit.IsBeingEdited Then
            pWSedit.AbortEditOperation()
            pWSedit.StartEditing(False)
        End If
        'pStatusBar.HideProgressBar
        pWS = Nothing

        ' pfeature = Nothing
        uninit()
    End Sub

    Public Sub ADO_openTIPorMTP(ByVal dbtype As Long, ByVal both As Boolean)
       
    End Sub




    Public Function get_MasterNetwork(ByVal transrefType As Long, ByVal pMap As IMap) As IFeatureLayer
        'function that returns the TransRefEdges Layer if transrefType = 0
        'function that returns the TransRefJunction Layer if transrefType = 1
        'on error GoTo eh

        Dim pFeatLayer As IFeatureLayer

        Dim i As Long
        For i = 0 To pMap.LayerCount - 1
            If (transrefType = 0) Then
                If (pMap.Layer(i).Name = m_layers(0)) Then
                    pFeatLayer = pMap.Layer(i)

                End If
            Else
                If (pMap.Layer(i).Name = m_layers(1)) Then
                    pFeatLayer = pMap.Layer(i)

                End If
            End If
        Next i
        get_MasterNetwork = pFeatLayer
        If pFeatLayer Is Nothing Then
            MsgBox("Error: did not find the transRef feature class for transrefType=" & transrefType)
        End If

        Exit Function
eh:
        CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.get_MasterNetwork")
        MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.get_MasterNetwork")

    End Function

Public Function get_Workspace() As IWorkspace
  'returns this ArcMap's workspace- needed to open tables
  'on error GoTo eh

  'jaf 091206: original code replaced (by someone else)
        '            with use of global g_FWS ( in setSDEworkspace and ModelApp_Click)
  '            stripped out old code since it does nothing
  
        get_Workspace = g_FWS

  Exit Function

eh:
        CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.get_Workspace")
        MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.get_Workspace")
End Function

  Public Function get_TableClass(tablename As String) As ITable
    'returns the table stated in tablename
    'on error GoTo eh
        If fVerboseLog Then WriteLogLine("get_TableClass called with tablename=" & tablename)
        If fVerboseLog Then WriteLogLine("---------------------------------------------------")

        Try

        
            Dim pTable As ITable

            'pan get the right table SDE or not

            'If (strLayerPrefix = "SDE") Then
            pTable = g_FWS.OpenTable(tablename)
            'Else
            ' pTable = OpenTable(tablename)
            'End If

            get_TableClass = pTable
            If pTable Is Nothing Then
                MsgBox("Error: did not find the Table " + tablename)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        Exit Function

eh:
        CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.get_TableClass")
        MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.get_TableClass")
  End Function



   




    Public Function get_RouteEventSource(ByVal routename As String, ByVal app As IApplication) As ESRI.ArcGIS.Location.IRouteEventSource

        'on error GoTo eh
        WriteLogLine("get_RouteEventSource called with routename=" & routename)
        WriteLogLine("---------------------------------------------------")

        Dim pFeatLayer As IFeatureLayer
        Dim pFCls As IFeatureClass

        Dim i As Long
        m_Doc = app.Document
        m_Map = m_Doc.FocusMap

        For i = 0 To m_Map.LayerCount - 1
            If TypeOf m_Map.Layer(i) Is IFeatureLayer Then
                If (m_Map.Layer(i).Name = routename) Then
                    pFeatLayer = m_Map.Layer(i)
                    pFCls = pFeatLayer.FeatureClass
                    Exit For
                    'MsgBox "found " + CStr(pFeatLayer.name)
                End If
            End If
        Next i

        get_RouteEventSource = pFCls
        If pFCls Is Nothing Then
            'MsgBox "Adding Route event source feature class" + routename
            GlobalMod.AddEventLayers()
            For i = 0 To m_Map.LayerCount - 1
                If TypeOf m_Map.Layer(i) Is IFeatureLayer Then
                    If (m_Map.Layer(i).Name = routename) Then
                        pFeatLayer = m_Map.Layer(i)
                        pFCls = pFeatLayer.FeatureClass
                        Exit For
                        'MsgBox "found " + CStr(pFeatLayer.name)
                    End If
                End If
            Next i
            get_RouteEventSource = pFCls
        End If

        Exit Function

eh:
        CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.get_RouteEventSource")
        MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.get_RouteEventSource")

    End Function


Public Sub AddEventLayers()
  'way to add event source feature classes through code- not used currently but could be used later
  '+++ VBA code that shows how to add a line RouteEventSource as a layer in the Map
  'MsgBox "Adding missing event route source layeres"
  'on error GoTo eh

        WriteLogLine("called GlobalMod.AddEventLayers")
        WriteLogLine("-------------------------------")

  '+++ Get the event table. It is called 'pavement'.
  Dim pTblColl As ITableCollection
  Dim pEventLineTable As ITable
  Dim pEventPointTable As ITable
   Dim pEventTollTable As ITable
  Dim i As Long
  Dim pDS As IDataset


        pTblColl = m_Map
  For i = 0 To pTblColl.TableCount - 1
            pDS = pTblColl.Table(i)
    If pDS.BrowseName = m_layers(4) Then 'evtLineProjectOutcomes
                pEventLineTable = pDS
      'Exit For
    End If
    If pDS.BrowseName = m_layers(6) Then
                pEventPointTable = pDS
      'Exit For
    End If
    If pDS.BrowseName = m_layers(10) Then 'tblLineTolls
                pEventTollTable = pDS
      'Exit For
    End If
  Next i
  If pEventLineTable Is Nothing Then
            MsgBox("Event table evtLineProjectsOutcomes not in MXD", vbExclamation, "AddLineEventLayer")
    Exit Sub
  End If
  If pEventPointTable Is Nothing Then
            MsgBox("Event table evtPointProjectsOutcomes not in MXD", vbExclamation, "AddPointEventLayer")
    Exit Sub
  End If
  If pEventTollTable Is Nothing Then
            MsgBox("Event table tblLineTolls not in MXD", vbExclamation, "AddTollEventLayer")
    Exit Sub
  End If
  '+++ Get the route feature class.
  Dim pLayer As ILayer
  Dim pFLayer As IFeatureLayer
  Dim pRouteFc As IFeatureClass
  For i = 0 To m_Map.LayerCount - 1
            pLayer = m_Map.Layer(i)
    If pLayer.Name = m_layers(8) Then  'ProjectRoutes
        If TypeOf pLayer Is IFeatureLayer Then
                    pFLayer = pLayer
                    pRouteFc = pFLayer.FeatureClass
          Exit For
        End If
    End If
  Next i
  If pRouteFc Is Nothing Then
            MsgBox("Could not find the route feature class", vbExclamation, "AddEventLayers")
    Exit Sub
  End If


  '+++ Create the route event source for line...
        WriteLogLine("adding evtLineProjectsOutcomes")
  '+++ The route locator
  Dim pName As iName
  Dim pRMLName As IRouteLocatorName
        pDS = pRouteFc
        pName = pDS.FullName
        pRMLName = New ESRI.ArcGIS.Location.RouteMeasureLocatorName
  With pRMLName
            .RouteFeatureClassName = pName
    .RouteIDFieldName = "projRteID"
    .RouteIDIsUnique = True
            .RouteMeasureUnit = esriUnits.esriUnknownUnits
    '.RouteWhereClause = ""
  End With

  '+++ Create the route event properties
  '+++ We will be using IRouteEventProperties2 to take advantage of adding an error field
  Dim pRtProp As IRouteEventProperties2
        Dim pRMLnProp As ESRI.ArcGIS.Location.IRouteMeasureLineProperties
        pRtProp = New ESRI.ArcGIS.Location.RouteMeasureLineProperties

  With pRtProp
            .EventMeasureUnit = esriUnits.esriUnknownUnits
    .EventRouteIDFieldName = "projRteID"
    '.LateralOffsetFieldName = "offset"
    .AddErrorField = True             'add field for locating errors
    .ErrorFieldName = "LOC_ERRORS"    'specify name for the locating errors field
  End With
        pRMLnProp = pRtProp
  pRMLnProp.FromMeasureFieldName = "MStart"
  pRMLnProp.ToMeasureFieldName = "MStop"

        pDS = pEventLineTable
        pName = pDS.FullName
  Dim pRESN As IRouteEventSourceName
        pRESN = New ESRI.ArcGIS.Location.RouteEventSourceName
  With pRESN
            .EventTableName = pName
            .EventProperties = pRMLnProp
            .RouteLocatorName = pRMLName
  End With

  '+++ By opening a route event source name object, you have a 'dynamic'
  '+++ feature class ...
  Dim pEventFC As IFeatureClass
        pName = pRESN
        pEventFC = pName.Open

  '+++ Create the layer and add it to the current map
  Dim pActive As IActiveView
        pFLayer = New FeatureLayer
        pFLayer.FeatureClass = pEventFC
  pFLayer.Name = m_layers(5)
        m_Map.AddLayer(pFLayer)

        pActive = m_Doc.ActiveView
  pActive.Refresh


  '+++ Create the route event source for point...
        WriteLogLine("adding evtPointProjectsOutcomes")
  '+++ The route locator
  Dim pNamep As iName
  Dim pRMPName As IRouteLocatorName
        pDS = pRouteFc
        pNamep = pDS.FullName
        pRMPName = New ESRI.ArcGIS.Location.RouteMeasureLocatorName
  With pRMPName
            .RouteFeatureClassName = pName
    .RouteIDFieldName = "projRteID"
    .RouteIDIsUnique = True
            .RouteMeasureUnit = esriUnits.esriUnknownUnits
    '.RouteWhereClause = ""
  End With

  '+++ Create the route event properties
  '+++ We will be using IRouteEventProperties2 to take advantage of adding an error field
  Dim pRtPropP As IRouteEventProperties2
        Dim pRMPtProp As ESRI.ArcGIS.Location.IRouteMeasurePointProperties
        pRtPropP = New ESRI.ArcGIS.Location.RouteMeasurePointProperties
  With pRtPropP
            .EventMeasureUnit = esriUnits.esriUnknownUnits
    .EventRouteIDFieldName = "projRteID"
    '.LateralOffsetFieldName = "offset"
    .AddErrorField = True             'add field for locating errors
    .ErrorFieldName = "LOC_ERRORS"    'specify name for the locating errors field
  End With
        pRMPtProp = pRtPropP

  pRMPtProp.MeasureFieldName = "M"

        pDS = pEventPointTable
        pNamep = pDS.FullName
  Dim pRESNp As IRouteEventSourceName
        pRESNp = New ESRI.ArcGIS.Location.RouteEventSourceName
  With pRESNp
            .EventTableName = pNamep
            .EventProperties = pRMPtProp
            .RouteLocatorName = pRMPName
  End With

  '+++ By opening a route event source name object, you have a 'dynamic'
  '+++ feature class ...
  Dim pEventFCp As IFeatureClass
        pNamep = pRESNp
        pEventFCp = pNamep.Open

  '+++ Create the layer and add it to the current map

        pFLayer = New FeatureLayer
        pFLayer.FeatureClass = pEventFCp
  pFLayer.Name = m_layers(7)  '+++ "Pavement_Events"
        m_Map.AddLayer(pFLayer)

        pActive = m_Doc.ActiveView
  pActive.Refresh

  '+++ Create the route mode event properties
  '+++ We will be using IRouteEventProperties2 to take advantage of adding an error field
        WriteLogLine("adding tblLineTolls")
        pRtProp = New ESRI.ArcGIS.Location.RouteMeasureLineProperties
  With pRtProp
            .EventMeasureUnit = esriUnits.esriUnknownUnits
    .EventRouteIDFieldName = "projRteID"
    '.LateralOffsetFieldName = "offset"
    .AddErrorField = True             'add field for locating errors
    .ErrorFieldName = "LOC_ERRORS"    'specify name for the locating errors field
  End With
        pRMLnProp = pRtProp
  pRMLnProp.FromMeasureFieldName = "Mstart" 'this measure fields in tblLineTolls
  pRMLnProp.ToMeasureFieldName = "Mstop"

        pDS = pEventTollTable
        pName = pDS.FullName
        pRESN = New ESRI.ArcGIS.Location.RouteEventSourceName
  With pRESN
            .EventTableName = pName
            .EventProperties = pRMLnProp
            .RouteLocatorName = pRMLName
  End With

  '+++ By opening a route event source name object, you have a 'dynamic'
  '+++ feature class ...
        pName = pRESN
        pEventFC = pName.Open

  '+++ Create the layer and add it to the current map
        pFLayer = New FeatureLayer
        pFLayer.FeatureClass = pEventFC
  pFLayer.Name = m_layers(11) 'tblLineTolls Events
        m_Map.AddLayer(pFLayer)

        pActive = m_Doc.ActiveView
  pActive.Refresh

  Exit Sub

eh:

        CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.AddEventLayers")
        MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.AddEventLayers")

'  Dim lNum As Long, sDesc As String, sSrc As String
'  lNum = Err.Number
'  sDesc = Err.Description
'  sSrc = Err.Source
'  Err.Raise lNum, sSrc, sDesc
End Sub

Public Function ExportFeatureLayer(pFLayer As IFeatureLayer, sLocation As String, pSelSet As ISelectionSet, o_ShpFilename As String) As IFeatureClass
  'on error GoTo eh

  Dim sProcName As String
  sProcName = "ExportFeatureLayer"

  ' init
        ExportFeatureLayer = Nothing

    Dim pInFc As IFeatureClass
        pInFc = pFLayer.FeatureClass

    Dim pDispTable As IDisplayTable
        pDispTable = pFLayer

    Dim pTable As ITable
        pTable = pDispTable.DisplayTable

  ' Get the FcName from the featureclass
    Dim pInDataset As IDataset
        pInDataset = pTable 'pInFc

    Dim pInDsName As IDatasetName
        pInDsName = pInDataset.FullName

  ' Create a new feature class name
  ' Define the output feature class name
    Dim pOutFeatureClassName As IFeatureClassName
        pOutFeatureClassName = New FeatureClassName

    Dim pOutDatasetName As IDatasetName
        pOutDatasetName = pOutFeatureClassName

    pOutDatasetName.Name = o_ShpFilename   ' here you  the output shpfile name

    Dim pOutWorkspaceName As IWorkspaceName
     pOutWorkspaceName = New WorkspaceName

    pOutWorkspaceName.pathName = sLocation ' here is your location
    pOutWorkspaceName.WorkspaceFactoryProgID = "esriDataSourcesFile.ShapefileWorkspaceFactory.1"
     pOutDatasetName.WorkspaceName = pOutWorkspaceName

  ' output shapefile characteristics
        pOutFeatureClassName.FeatureType = esriFeatureType.esriFTSimple
    pOutFeatureClassName.ShapeType = pInFc.ShapeType
    pOutFeatureClassName.ShapeFieldName = "Shape"

  ' Export
        Dim pExportOp As IExportOperation
    'MsgBox "before pExportOp.ExportFeatureClass"
     pExportOp = New ExportOperation
        pExportOp.ExportFeatureClass(pInDsName, Nothing, pSelSet, Nothing, pOutDatasetName, 0)
    'pExportOp.ExportFeatureClass pInDsName, Nothing, pSelSet, Nothing, pOutDatasetName, 0
    Dim pOutFeatureClass As IFeatureClass
  ' get the feature classes from the name
  ' see pg 72 Vol2, Exploring ArcObjects Book
    Dim pOutName As iName
    'MsgBox "Before  pOutFeatureClass = pOutName.Open"
     pOutName = pOutFeatureClassName    'QI
     pOutFeatureClass = pOutName.Open
   '  the function return
    'MsgBox "Before  ExportFeatureLayer"

     ExportFeatureLayer = pOutFeatureClass

    Exit Function

eh:

        CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.ExportFeatureLayer")
        MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.ExportFeatureLayer")

'    MsgBox "Error number: " & CStr(Err.Number) & _
'           vbCrLf & "Error description: " & Err.Description & _
'           vbCrLf & "Location:" & sProcName, _
'           vbExclamation + vbOKOnly, "Your Title"

End Function

Public Sub exportGfeatureclass(pFLayer As IFeatureLayer, ptempws1 As IWorkspace, pSelSet As ISelectionSet, jname As String)
    'on error GoTo eh
        Debug.Print("exportGfeatureClass: " & jname & " start - " & Now())
    Dim pInFc As IFeatureClass
     pInFc = pFLayer.FeatureClass

    Dim pDispTable As IDisplayTable
     pDispTable = pFLayer

    Dim pTable As ITable
     pTable = pDispTable.DisplayTable
     'Get the FcName from the featureclass
    Dim pInDataset As IDataset
     pInDataset = pTable 'pInFc

     Dim pInDsName As IDatasetName
     pInDsName = pInDataset.FullName

 Dim pOutWSN As IWorkspaceName
  pOutWSN = New WorkspaceName
Dim pOutDSN As IDatasetName
        Dim pOutFeatDSN As IFeatureDatasetName
        Dim pOutFCN As IFeatureClassName
        Dim pOutDSN2 As IDatasetName

 pOutWSN.ConnectionProperties = ptempws1.ConnectionProperties
 'MsgBox "ptempws1.Type=" + ptempws1.Type
        If ptempws1.Type = esriWorkspaceType.esriRemoteDatabaseWorkspace Then
            pOutWSN.WorkspaceFactoryProgID = "esriCore.SdeWorkspaceFactory.1"
        ElseIf ptempws1.Type = esriWorkspaceType.esriLocalDatabaseWorkspace Then
            pOutWSN.WorkspaceFactoryProgID = "esriCore.AccessWorkspaceFactory.1"
        Else
            pOutWSN.WorkspaceFactoryProgID = "esriCore.ShapefileWorkspaceFactory.1"
        End If


        If Not pOutWSN.Type = esriWorkspaceType.esriFileSystemWorkspace Then   '+++ makes no sense for shapefile
           
            '+++ Create the new output FeatureClassName object
            
            pOutFCN = New FeatureClassName
            pOutDSN2 = pOutFCN
            pOutDSN2.WorkspaceName = pOutWSN 'esp. necessary when pOutDSN is Nothing
            ' pDataset = pTable
            pOutDSN2.Name = jname

        End If
'MsgBox pInDsName.name
'MsgBox pSelSet.count
        Dim pExportOp As IExportOperation
     pExportOp = New ExportOperation
    'MsgBox "Before export operation"
        pExportOp.ExportFeatureClass(pInDsName, Nothing, pSelSet, Nothing, pOutDSN2, 0)
  ' MsgBox "exported"

        WriteLogLine("Completed exporting feature class" & pOutDSN2.Name)
        Debug.Print("exportGfeatureClass: " & jname & " end - " & Now())
        Exit Sub

eh:

        CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.exportGfeatureclass")
        MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.exportGfeatureclass")


End Sub


Public Sub getFuture_prjSelectionSet()
'trial run approach of getting the future projects to be in service
'do not use in either Input or Ouput 9/11/2003
 'start w/ ProjID_ver-> get prjrteid to link w/ tblpoint/lineprojects-> get edge/junciton id
  Dim pStatusBar As IStatusBar
  Dim pProgbar As IStepProgressor
   pStatusBar = m_App.StatusBar
   pProgbar = pStatusBar.ProgressBar

  Dim pFeatLayerE As IFeatureLayer, pFeatLayerJ As IFeatureLayer
        pFeatLayerE = get_MasterNetwork(0, m_Map)
        pFeatLayerJ = get_MasterNetwork(1, m_Map)

  Dim pTblLine As ITable
  'Dim ptblPoint As ITable no loger using tblPoint table
  Dim pTblMode As ITable
  Dim pWorkspace As IWorkspace
   pWorkspace = get_Workspace
   pTblLine = get_TableClass(m_layers(3))
  ' ptblPoint = get_TableClass("tblPointProjects", pWorkspace)
   pTblMode = get_TableClass(m_layers(2))

  Dim pFeatSelect As IFeatureSelection
  Dim pSSet As ISelectionSet

  Dim pQF As IQueryFilter, pQFt As IQueryFilter
  Dim pFeaturePR As IFeature
  Dim pRow As IRow, pdRow As IRow
  'Dim tempfeatID As String
  Dim tempID As Long
  Dim rIndex As Long, idIndex As Long, index As Long
  Dim tDate As Date, t2Date As Date

  Dim pFeatSelection As IFeatureSelection, pFeatSelectPR As IFeatureSelection
  Dim pESSet As ISelectionSet, pJSSet As ISelectionSet
  Dim pFC As IFeatureCursor
  Dim pTC As ICursor, pDuplicateTC As ICursor

  Dim i As Long
  For i = 0 To prjselectedCol.count - 1
     pQFt = New QueryFilter
    pQFt.WhereClause = "projID_Ver = '" + prjselectedCol.Item(i).PrjId + "'"

    If (pTblLine.rowcount(pQFt) > 0) Then
        'MsgBox "found it in line"
         pFeatSelect = pFeatLayerE
                pFeatSelect.SelectFeatures(Nothing, esriSelectionResultEnum.esriSelectionResultNew, False)
         pSSet = pFeatSelect.SelectionSet
         pTC = pTblLine.Search(pQFt, False)
         pRow = pTC.NextRow
        idIndex = pRow.Fields.FindField("PSRCEdgeID")
        tempID = pRow.value(idIndex)
         pQF = New QueryFilter
        pQF.WhereClause = "PSRCEdgeID = " + CStr(tempID)
                pSSet.Search(pQF, False, pFC)

    'ElseIf (ptblPoint.rowcount(pQFt) > 0) Then
         'MsgBox "found it in point"
         ' pTC = ptblPoint.Search(pQFt, False)
         ' pRow = pTC.NextRow
         'deleted code 01-19-05
    Else
        'MsgBox "didn't find it"
    End If

 Next i

End Sub





Public Sub addNewJunction(projRteID As String, evtType As Integer, pPoint As IPoint)
  'on error GoTo eh
        WriteLogLine("called GlobalMod.addNewJunction with projRteID=" & projRteID)
        WriteLogLine("------------------------------------------------------------")

  'jaf--this creates a new feature in the junction shapefile
  Dim pJunctFeat As IFeature 'to use to add new junctions created from events
   pJunctFeat = m_junctShp.CreateFeature
  Dim pGeom As IGeometry
   pGeom = pPoint
   pJunctFeat.Shape = pGeom
  pJunctFeat.Store
  Dim index As Integer
  index = pJunctFeat.Fields.FindField("Scen_Node")
  LargestJunct = LargestJunct + 1

  pJunctFeat.value(index) = LargestJunct
  index = pJunctFeat.Fields.FindField("PSRCJunctID")
  pJunctFeat.value(index) = LargestJunct
  pJunctFeat.Store

  Exit Sub

eh:

        CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.addNewJunction")
        MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.addNewJunction")

End Sub

    Public Sub SplitFeature(ByVal pFeature As IFeature, ByVal pSplitPoint As IPoint, ByVal projRte As String)
        'MsgBox "spliting"

        'on error GoTo eh
        If fVerboseLog Then WriteLogLine("called splitfeature with projRte=" & projRte)

        Dim pEnumVertex As IEnumVertex
        Dim pGeoColl As IGeometryCollection
        Dim pPolyCurve As IPolycurve2
        'Dim pEnumSplitPoint As IEnumSplitPoint
        Dim pNewFeature As IFeature
        Dim PartCount As Integer
        Dim index As Long, indexJ As Long

        Dim pFeatLayer As IFeatureLayer
        pFeatLayer = New FeatureLayer
        pFeatLayer.FeatureClass = m_junctShp

        Dim pFilter As ISpatialFilter
        Dim pFC As IFeatureCursor
        Dim pJFeat As IFeature

        Dim SplitID As Long
        SplitID = LargestJunct

        pPolyCurve = pFeature.Shape
        Dim bSplit As Boolean, lPart As Long, lSeg As Long
        pPolyCurve.SplitAtPoint(pSplitPoint, False, True, bSplit, lPart, lSeg)
        ' MsgBox bSplit
        pGeoColl = pPolyCurve

        Dim originalI As Long
        Dim originalJ As Long
        index = pFeature.Fields.FindField("INode")
        originalI = pFeature.value(index)
        index = pFeature.Fields.FindField("JNode")
        originalJ = pFeature.value(index)

        Dim nodecnt As Long
        Dim i As Integer
        For i = 0 To 1
            If i = 0 Then
                pFeature.Shape = BUILDPOLYLINE(pGeoColl.Geometry(i)) 'orig feature becomes first part
                pFeature.Store()
                pNewFeature = pFeature
            Else
                pNewFeature = m_edgeShp.CreateFeature
                pNewFeature.Shape = BUILDPOLYLINE(pGeoColl.Geometry(i))
                CopyAttributes(pFeature, pNewFeature, projRte)
                pNewFeature.Store()
            End If
            pFilter = New SpatialFilter
            With pFilter
                .Geometry = pNewFeature.ShapeCopy
                .GeometryField = m_junctShp.ShapeFieldName
                .SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects
            End With
            If m_junctShp.featurecount(pFilter) > 0 Then

                pFC = pFeatLayer.Search(pFilter, False)
                pJFeat = pFC.NextFeature

                Do Until pJFeat Is Nothing
                    nodecnt = 0
                    'keep original direction
                    index = pFeature.Fields.FindField("UseEmmeN")
                    If pNewFeature.value(index) = 0 Then
                        indexJ = pJFeat.Fields.FindField("PSRCJunctID")
                    Else
                        indexJ = pJFeat.Fields.FindField("Emme2nodeID")
                    End If
                    If (IsDBNull(pJFeat.Value(indexJ)) = False) Then
                        If (originalI = pJFeat.Value(indexJ)) Then

                            index = pNewFeature.Fields.FindField("INode")
                            pNewFeature.Value(index) = pJFeat.Value(indexJ)
                            index = pNewFeature.Fields.FindField("JNode")
                            pNewFeature.Value(index) = SplitID

                            index = pNewFeature.Fields.FindField("UseEmmeN")

                            If pNewFeature.Value(index) > 0 Then pNewFeature.Value(index) = 3
                            If pNewFeature.Value(index) = 0 Then pNewFeature.Value(index) = 2
                        ElseIf (originalJ = pJFeat.Value(indexJ)) Then

                            index = pNewFeature.Fields.FindField("JNode")
                            pNewFeature.Value(index) = pJFeat.Value(indexJ)
                            index = pNewFeature.Fields.FindField("INode")
                            pNewFeature.Value(index) = SplitID
                            index = pNewFeature.Fields.FindField("UseEmmeN")

                            If pNewFeature.Value(index) > 0 Then pNewFeature.Value(index) = 3
                            If pNewFeature.Value(index) = 0 Then pNewFeature.Value(index) = 1
                        End If
                    Else 'its Null
                        If fVerboseLog Then WriteLogLine("PSRCJunctionID is null")
                    End If

                    pJFeat = pFC.NextFeature
                Loop
                pNewFeature.Store()
            End If
            pJFeat = Nothing
        Next i

        pFeatLayer = Nothing
        Exit Sub

eh:
        CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.SplitFeature")
        MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.SplitFeature")

        'Print #1, Err.Description, vbInformation, "SplitFeatures"
        pFeatLayer = Nothing
    End Sub

    Public Sub SplitFeatureByM0(ByVal pFeature As IFeature, ByVal pSplitM As Double, ByVal onode As Long, ByVal NewPoint As IPoint, ByVal lType As String, ByVal dctEdges As Dictionary(Of Object, Object), ByVal dctJct As Dictionary(Of Object, Object))
        'pSplit needs to be shadow length from FromPoint
        'on error GoTo ErrChk
        If fVerboseLog Then WriteLogLine("called SplitFeatureByM")
        Dim tempSplitM As Double
        tempSplitM = pSplitM

        Dim pEnumVertex As IEnumVertex
        Dim pGeoColl As IGeometryCollection
        Dim pPolyCurve As IPolycurve2
        'Dim pEnumSplitPoint As IEnumSplitPoint
        Dim pNewFeature As IFeature
        Dim PartCount As Integer
        Dim index As Long, indexJ As Long

        Dim pFeatLayer As IFeatureLayer
        pFeatLayer = New FeatureLayer
        pFeatLayer.FeatureClass = m_junctShp

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

        pPline = pFeature.ShapeCopy
        pPolyCurve = pFeature.Shape
        FromPt = pPline.FromPoint
        ToPt = pPline.ToPoint
        pFilter = New SpatialFilter
        With pFilter
            .Geometry = ToPt
            .GeometryField = m_junctShp.ShapeFieldName
            .SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects
        End With

        If pFeatLayer.FeatureClass.FeatureCount(pFilter) > 0 Then
            pFC = pFeatLayer.Search(pFilter, False)
            pJFeat = pFC.NextFeature
            index = pJFeat.Fields.FindField("Scen_Node")
            If pJFeat.Value(index) = onode Then
                tempSplitM = pPline.Length - tempSplitM
                found = True
                If fVerboseLog Then WriteLogLine("its the from node" & pPline.Length & "splitm" & tempSplitM)

                pPolyCurve.GetSubcurve(0, tempSplitM, bratio, pSplit1)
                pGeoColl = New GeometryBag
                pGeoColl.AddGeometry(pSplit1)    'pSplit1 is the long part
                pPolyCurve.GetSubcurve(tempSplitM, pPolyCurve.Length, False, pSplit2)
                pGeoColl.AddGeometry(pSplit2)    'pSplit2 is the short part
            End If
        End If

        If found = False Then 'check from node
            pFilter = New SpatialFilter
            With pFilter
                .Geometry = FromPt
                .GeometryField = m_junctShp.ShapeFieldName
                .SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects
            End With
            If pFeatLayer.FeatureClass.FeatureCount(pFilter) > 0 Then
                pFC = pFeatLayer.Search(pFilter, False)
                pJFeat = pFC.NextFeature
                index = pJFeat.Fields.FindField("Scen_Node")
                If pJFeat.Value(index) = onode Then
                    tempSplitM = tempSplitM
                    found = True
                    If fVerboseLog Then WriteLogLine("its the to node" & pPline.Length & "splitm" & tempSplitM)

                    pPolyCurve.GetSubcurve(0, tempSplitM, bratio, pSplit2)
                    pGeoColl = New GeometryBag
                    pGeoColl.AddGeometry(pSplit2)    'pSplit1 is the long part
                    pPolyCurve.GetSubcurve(tempSplitM, pPolyCurve.Length, False, pSplit1)
                    pGeoColl.AddGeometry(pSplit1)    'pSplit2 is the long part
                End If
            End If
        End If
        'else it must be the from so already  tempsplitm ok
        SplitM = True



        '[033006] hyu: declear two variable to hold the new geometries from split


        bpart = True 'so create new part
        '    '[033006] hyu: SplitAtDistance doesn't work. change to GetSubCurve
        ''  pPolyCurve.SplitAtDistance tempSplitM, bratio, bpart, bSplit, lpart, lSeg
        ''  pGeoColl = pPolyCurve
        '    pPolyCurve.GetSubcurve 0, tempSplitM, bratio, pSplit1
        '     pGeoColl = New GeometryBag
        '    pGeoColl.AddGeometry pSplit1
        '    pPolyCurve.GetSubcurve tempSplitM, pPolyCurve.length, False, pSplit2
        '    pGeoColl.AddGeometry pSplit2

        Dim originalI As Long
        Dim originalJ As Long
        index = pFeature.Fields.FindField("INode")
        originalI = pFeature.Value(index)
        index = pFeature.Fields.FindField("JNode")
        originalJ = pFeature.Value(index)
        orgstr = CStr(originalI + m_Offset) + " " + CStr(originalJ + m_Offset)
        orgstr2 = CStr(originalJ + m_Offset) + " " + CStr(originalI + m_Offset)

        Dim i As Integer
        For i = 0 To 1
            If i = 0 Then
                '[033106] hyu: no need to build polyline.
                ' pFeature.Shape = BUILDPOLYLINE(pGeoColl.Geometry(i)) 'orig feature becomes first part
                pFeature.Shape = pSplit1
                pFeature.Store()
                pNewFeature = pFeature

            Else
                pNewFeature = m_edgeShp.CreateFeature
                '[033106] hyu: no need to build polyline.
                '        pNewFeature.Shape = BUILDPOLYLINE(pGeoColl.Geometry(i))
                pNewFeature.Shape = pSplit2
                CopyAttributes(pFeature, pNewFeature, "")
                pNewFeature.Store()

            End If
            pFilter = New SpatialFilter
            Dim pGeom As IGeometry
            pGeom = pGeoColl.Geometry(i)
            With pFilter
                .Geometry = pNewFeature.Shape
                .GeometryField = m_junctShp.ShapeFieldName
                .SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects
            End With

            If m_junctShp.FeatureCount(pFilter) > 0 Then 'should get one node- the new one is not in the intermediate junct shp

                pFC = m_junctShp.Search(pFilter, False)
                pJFeat = pFC.NextFeature
                Do Until pJFeat Is Nothing
                    Dim nodecnt As Long
                    nodecnt = 0
                    Dim temp As Long
                    'keep original direction
                    indexJ = pJFeat.Fields.FindField("Scen_Node")
                    If (IsDBNull(pJFeat.Value(indexJ)) = False) Then

                        If ((originalI + m_Offset) = pJFeat.Value(indexJ)) Then
                            index = pNewFeature.Fields.FindField("INode")
                            temp = pJFeat.Value(indexJ) - m_Offset
                            pNewFeature.Value(index) = temp
                            index = pNewFeature.Fields.FindField("JNode")


                            pNewFeature.Value(index) = SplitID - m_Offset

                            temp = pJFeat.Value(indexJ)
                            If i = 0 Then
                                replstr = CStr(temp) + " " + CStr(SplitID)
                                replstr2 = CStr(SplitID) + " " + CStr(temp)
                            Else

                                index = pNewFeature.Fields.FindField("Split" + lType)
                                pNewFeature.Value(index) = lType
                            End If
                            pNewFeature.Store()

                        ElseIf ((originalJ + m_Offset) = pJFeat.Value(indexJ)) Then
                            index = pNewFeature.Fields.FindField("JNode")
                            temp = pJFeat.Value(indexJ) - m_Offset
                            pNewFeature.Value(index) = temp
                            index = pNewFeature.Fields.FindField("INode")
                            pNewFeature.Value(index) = SplitID - m_Offset

                            temp = pJFeat.Value(indexJ)
                            If i = 0 Then
                                replstr = CStr(SplitID) + " " + CStr(temp)
                                replstr2 = CStr(temp) + " " + CStr(SplitID)
                            Else
                                index = pNewFeature.Fields.FindField("Split" + lType)
                                pNewFeature.Value(index) = lType
                            End If
                            pNewFeature.Store()
                        End If
                    Else 'its Null
                        WriteLogLine("DATA ERROR: PSRCJunctionID is null in SplitFeatureByM")
                    End If
                    pPline = pNewFeature.ShapeCopy

                    If pPline.FromPoint.X = FromPt.X And pPline.FromPoint.Y = FromPt.Y Then 'need to get point so can get geomtry
                        NewPoint = pPline.ToPoint
                    Else
                        NewPoint = pPline.FromPoint
                    End If
                    pJFeat = pFC.NextFeature
                Loop
                pNewFeature.Store()

                index = pNewFeature.Fields.FindField(pNewFeature.Class.OIDFieldName)
                spID = pNewFeature.Value(index)

                If dctEdges.ContainsKey(CStr(spID)) Then
                    dctEdges.Item(CStr(spID)) = pNewFeature
                Else
                    dctEdges.Add(CStr(spID), pNewFeature)
                End If



            End If
            pJFeat = Nothing
        Next i

        Dim pNewJct As IFeature
        pNewJct = m_junctShp.CreateFeature
        pNewJct.Shape = NewPoint
        pNewJct.Value(pNewJct.Fields.FindField(g_PSRCJctID)) = LargestJunct
        pNewJct.Value(pNewJct.Fields.FindField("Scen_Node")) = LargestJunct + m_Offset
        pNewJct.Store()
        dctJct.Add(LargestJunct + m_Offset, pNewJct)

        '[033006] hyu: "FID" cannot be found, so change to OIDFieldName
        'index = pNewFeature.Fields.FindField("FID")
        'Index = pNewFeature.Fields.FindField(pNewFeature.Class.OIDFieldName)
        'spID = pNewFeature.value(Index)

        pFeatLayer = Nothing
        Exit Sub

ErrChk:
        CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.SplitFeatureByM")
        MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.SplitFeatureByM")
        'Print #9, Err.Description, vbInformation, "SplitFeatures"
        pFeatLayer = Nothing
    End Sub

Public Sub CopyAttributes(pSourceFeature As IRow, pDestinationFeature As IRow, projRte As String)
  'on error GoTo eh
  'MsgBox "copying attrib"
  Dim pField As IField
  Dim pFields As IFields
  Dim pRow As IRow
  Dim FieldCount As Integer
  Dim index As Long, Tindex As Long

   pFields = pSourceFeature.Fields
  For FieldCount = 0 To pFields.FieldCount - 1
     pField = pFields.field(FieldCount)
    'If pField.Editable Then
            If Not pField.Type = esriFieldType.esriFieldTypeOID And Not pField.Type = esriFieldType.esriFieldTypeGeometry Then
                '[033006] hyu: check the editible property before assigning value
                index = pDestinationFeature.Fields.FindField(pField.Name)
                If index > -1 Then
                    If pDestinationFeature.Fields.Field(index).Editable Then pDestinationFeature.Value(index) = pSourceFeature.Value(FieldCount)
                End If
            End If
    'End If
  Next FieldCount
 pDestinationFeature.Store
 Exit Sub

eh:

        CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.CopyAttributes")
        MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.CopyAttributes")

End Sub

    Public Function BUILDPOLYLINE(ByVal pSegColl As ISegmentCollection) As IPolyline
        'on error GoTo eh

        If fVerboseLog Then WriteLogLine("called GlobalMod.buildpolyline")
        If fVerboseLog Then WriteLogLine("------------------------------")


        Dim pPolyline As IGeometryCollection
        Dim pGeometry As IGeometry
        pPolyline = New Polyline
        pPolyline.AddGeometries(1, pSegColl)
        BUILDPOLYLINE = pPolyline

        Exit Function

eh:
        CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.BUILDPOLYLINE")
        MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.BUILDPOLYLINE")
        'MsgBox BUILDPOLYLINE.length
    End Function


Public Sub dissolveTransEdgesNoX_0()
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
     pWorkspace = get_Workspace
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
    Dim lIndex As Long, upIndex As Long
    Dim psrcId As String
    Dim i As Long
    Dim fcountI, fcountIJ, mergeCountIJ, mergeCountJI As Long
    Dim idxJunctID As Integer
    Dim match As Boolean
    match = True
    Dim edgesdelete As Boolean
    edgesdelete = False

    Dim newI As Long, newJ As Long

    Dim pWorkspaceEdit As IWorkspaceEdit
     pWorkspaceEdit = pWorkspace

        pWorkspaceEdit.StartEditing(False)
    pWorkspaceEdit.StartEditOperation
    Dim pFSelect As IFeatureSelection
    Dim pSSet As ISelectionSet, pSSet2 As ISelectionSet

     pFSelect = pFLayerE
    'pFSelect.SelectFeatures Nothing, esriSelectionResultNew, False 'get all intermediate edges

        pSSet = m_edgeShp.Select(Nothing, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pWorkspace)  'pFSelect.SelectionSet
        pSSet2 = m_edgeShp.Select(Nothing, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pWorkspace) 'pFSelect.SelectionSet

        pSSet.Search(Nothing, False, pFeatCursor)
     pFeat = pFeatCursor.NextFeature
    '[011806] pan currently no PSRC_E2ID in edges or modeAttributes.
    'idIndex = pFeat.Fields.FindField("PSRC_E2ID")
    idIndex = pFeat.Fields.FindField("PSRCEDGEID")
    upIndex = pFeat.Fields.FindField("Updated1")

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

    Do Until pFeat Is Nothing
        For m = 0 To 1 'first look at this edges INode and then its JNode to see if should dissolve
            'if node dissolved should exit for loop
            match = True
            edgesdelete = False
            If m = 0 Then
                same = "INode"
                opposite = "JNode"
            Else
                same = "JNode"
                opposite = "INode"
            End If
            lSFld = pFeat.Fields.FindField(same)
            dnode = pFeat.value(lSFld)
            If (dnode > 0) Then
                lSFld = pFeat.Fields.FindField(opposite)
                If m = 0 Then
                    newJ = pFeat.value(lSFld)
                Else
                    newI = pFeat.value(lSFld)
                End If
                'should be psrcjunctionid
                 pQF = New QueryFilter
                pQF.WhereClause = same + " = " + CStr(dnode)
                fcountI = 0
                fcountIJ = 0
                fcountI = pFLayerE.FeatureClass.featurecount(pQF)
                If (fcountI <= 2) Then
                     pQF2 = New QueryFilter
                    pQF2.WhereClause = opposite + " = " + CStr(dnode)
                    fcountIJ = fcountI + pFLayerE.FeatureClass.featurecount(pQF2)
                     pQF = New QueryFilter
                    pQF.WhereClause = "PSRCJunctID = " + CStr(dnode) 'this at node in turnmovemnts
                    If (fcountIJ = 2) Then 'then only at most two edges coming in and this node {[061505] turn check removed here}
                        'now check if junction really real
                         pQFj = New QueryFilter
                        pQFj.WhereClause = "PSRCJunctID = " + CStr(dnode)
                         pFCj = pFLayerJ.FeatureClass.Search(pQFj, False)
                        'MsgBox "potential dnode" + CStr(dnode)
                        If (fcountI = 2) Then
                             pQF = New QueryFilter
                            pQF.WhereClause = same + " = " + CStr(dnode)

                                pSSet2.Search(pQF, False, pFC)

                             pFeat2 = pFC.NextFeature
                             pNextFeat = pFC.NextFeature
                            lSFld = pNextFeat.Fields.FindField(opposite)
                            dir = False
                            If m = 0 Then
                                newJ = pFeat2.value(lSFld)
                                newI = pNextFeat.value(lSFld)
                            Else
                              newI = pFeat2.value(lSFld)
                              newJ = pNextFeat.value(lSFld)
                            End If
                        Else
                             pFeat2 = pFeat
                            pSSet2.Search (pQF2, False, pFC2)

                             pNextFeat = pFC2.NextFeature
                            lSFld = pNextFeat.Fields.FindField(same)
                            dir = True
                            If m = 0 Then
                                newI = pNextFeat.value(lSFld)
                            Else
                                newJ = pNextFeat.value(lSFld)
                            End If
                        End If

                         pQFt = New QueryFilter
                        pQFt.WhereClause = "PSRCEdgeID = " + CStr(pFeat2.value(idIndex))
                         pTC = pTblMode.Search(pQFt, False)
                        'count the presence of records using this edge in modeAttributes, modeTolls, and tblTransitSegments
                        If (pTblMode.rowcount(pQFt) > 0) Then 'found first edge in modeAttributes {[061505] toll and transit check removed here}
            '[060705] pre-Transit Xchk code: If (pTblMode.rowcount(pQFt) > 0 And pTolls.rowcount(pQFt) = 0) Then 'found first edge in modeAttributes and NOT in modeTolls

                             pRow = pTC.NextRow
                             pQFt = New QueryFilter
                            pQFt.WhereClause = "PSRCEdgeID = " + CStr(pNextFeat.value(idIndex))
                             pTC = pTblMode.Search(pQFt, False)
                            If (pTblMode.rowcount(pQFt) > 0) Then 'found second edge in modeAttributes {[061505] toll and transit check removed here}
    '[060705] pre-transit Xchk code:  If (pTblMode.rowcount(pQFt) > 0 And pTolls.rowcount(pQFt) = 0) Then 'found second edge in modeAttributes and NOT in modeTolls
                                 pNextRow = pTC.NextRow

                                'compare attributes to detect if we have "real" pseudonode
                                If dir = True Then
                                    For i = 1 To pRow.Fields.FieldCount - 1
                                        If (pRow.Fields.field(i).Name <> "PSRCEdgeID") Then
                                            If (pRow.value(i) <> pNextRow.value(i)) Then
                                              match = False
                                              'MsgBox pRow.Fields.field(i).name + "diff"
                                            End If
                                        End If
                                    Next i
                                Else 'need to compare IJ to JI to match direction
                                    Dim j As Integer, l As Integer, t As Integer
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
                                        If (pRow.value(index2) <> pNextRow.value(indexNext)) Then
                                            match = False
                                              'MsgBox "speed diff"
                                        End If
                                        '[011806] pan FFS floating point comparison removed
                                        'index2 = pRow.Fields.FindField(dir2 + "FFS")
                                        'indexnext = pNextRow.Fields.FindField(dirnext + "FFS")
                                        'If (pRow.value(index2) <> pNextRow.value(indexnext)) Then
                                        '    match = False
                                              'MsgBox "ffs diff"
                                        'End If
                                        index2 = pRow.Fields.FindField(dir2 + "VDFunc")
                                        indexNext = pNextRow.Fields.FindField(dirNext + "VDFunc")
                                        If (pRow.value(index2) <> pNextRow.value(indexNext)) Then
                                            match = False
                                              'MsgBox "vdf diff"
                                        End If
                                        index2 = pRow.Fields.FindField(dir2 + "SideWalks")
                                        indexNext = pNextRow.Fields.FindField(dirNext + "SideWalks")
                                        If (pRow.value(index2) <> pNextRow.value(indexNext)) Then
                                            match = False
                                              'MsgBox "side diff"
                                        End If
                                        index2 = pRow.Fields.FindField(dir2 + "BikeLanes")
                                        indexNext = pNextRow.Fields.FindField(dirNext + "BikeLanes")
                                        If (pRow.value(index2) <> pNextRow.value(indexNext)) Then
                                            match = False
                                             'MsgBox "bike diff"
                                        End If
                                        For l = 0 To 3
                                            If l > 1 Then
                                                index2 = pRow.Fields.FindField(dir2 + "Lanes" + lType(l))
                                                indexNext = pNextRow.Fields.FindField(dirNext + "Lanes" + lType(l))
                                                If (pRow.value(index2) <> pNextRow.value(indexNext)) Then
                                                  match = False
                                                    'MsgBox "lanes diff"
                                                End If
                                            Else
                                                index2 = pRow.Fields.FindField(dir2 + "LaneCap" + lType(l))
                                                indexNext = pNextRow.Fields.FindField(dirNext + "LaneCap" + lType(l))
                                                If (pRow.value(index2) <> pNextRow.value(indexNext)) Then
                                                    match = False
                                                    'MsgBox "lanecap diff"
                                                End If
                                                For t = 0 To 4
                                                    index2 = pRow.Fields.FindField(dir2 + "Lanes" + lType(l) + timePd(t))
                                                    indexNext = pNextRow.Fields.FindField(dirNext + "Lanes" + lType(l) + timePd(t))

                                                    If (pRow.value(index2) <> pNextRow.value(indexNext)) Then
                                                        match = False
                                                        'MsgBox "lanes diff"
                                                    End If
                                                Next t
                                            End If
                                        Next l
                                    Next j
                                End If

                                If pFeat.value(upIndex) <> pNextFeat.value(upIndex) Then
                                    match = False
                                    'MsgBox "update diff"
                                End If
                                i = pFeat.Fields.FindField("Modes")
                                If pFeat.value(i) <> pNextFeat.value(i) Then
                                    match = False
                                     'MsgBox "modes diff"
                                End If
                                i = pFeat.Fields.FindField("LinkType")
                                If pFeat.value(i) <> pNextFeat.value(i) Then
                                    match = False
                                     'MsgBox "link diff"
                                End If

                                i = pFeat.Fields.FindField("FacilityType")
                                If pFeat.value(i) <> pNextFeat.value(i) Then
                                    match = False
                                     'MsgBox "fac diff"
                                End If

                                If (match = True) Then
                                    'MsgBox "found match"
                                     pGeometryBag = New GeometryBag
                                     pGeometry = pFeat2.ShapeCopy
                                        pGeometryBag.AddGeometry(pGeometry)
                                     pGeometry = pNextFeat.ShapeCopy
                                        pGeometryBag.AddGeometry(pGeometry)

                                    'update the merge count for the log
                                    If dir Then
                                      mergeCountIJ = mergeCountIJ + 1
                                    Else
                                      mergeCountJI = mergeCountJI + 1
                                    End If

                                    'merge the edge features
                                     pPolyline = New Polyline
                                     pTopo = pPolyline
                                        pTopo.ConstructUnion(pGeometryBag)
                                     pNewFeat = pFLayerE.FeatureClass.CreateFeature
                                     pNewFeat.Shape = pPolyline
                                    pNewFeat.Store

                                     pFlds = pNewFeat.Fields
                                    For i = 1 To pFlds.FieldCount - 1
                                        lIndex = pFeat2.Fields.FindField(pFlds.field(i).Name)
                                        If (lIndex <> -1) Then
                                                If (pFlds.Field(i).Type = esriFieldType.esriFieldTypeOID And _
                                            Not pFlds.Field(i).Type = esriFieldType.esriFieldTypeGeometry Or _
                                            pFlds.Field(i).Name <> "Shape") Then
                                                    'inode and jnode need to be  differently
                                                    If (pFlds.Field(i).Name = same) Then
                                                        pNewFeat.Value(i) = newI
                                                    ElseIf (pFlds.Field(i).Name = opposite) Then
                                                        pNewFeat.Value(i) = newJ
                                                    ElseIf (pFlds.Field(i).Name = "length") Then
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
                                    pNewFeat.Store
                                    edgesdelete = True
'                                    pFeat2.Delete
'                                    pNextFeat.Delete
                                    'find junction node and delete it

                                    If (pFLayerJ.FeatureClass.featurecount(pQFj) > 0) Then
                                        'MsgBox "found junction to delete"
                                         pJunctFeat = pFCj.NextFeature
                                        If fVerboseLog Then
                                            idxJunctID = pJunctFeat.Fields.FindField("PSRCJunctID")
                                            If idxJunctID < 0 Then
                                                    WriteLogLine("Deleting intrmed. shp junct OID=" & pJunctFeat.OID)
                                            Else
                                                    WriteLogLine("Deleting intrmed. shp junctID=" & pJunctFeat.Value(idxJunctID))
                                            End If
                                        End If
                                        pJunctFeat.Delete
                                    End If
                                        pSSet.Add(pNewFeat.OID)
                                    'MsgBox pSSet.count
                                End If
                            End If  'match = true
                        End If
                    End If
                End If
            End If

            If (edgesdelete = True) Then

                pFeat2.Delete
                pNextFeat.Delete
                edgesdelete = False
                'MsgBox "edges deleted"
                Exit For
            End If
        Next m
         pFeat = pFeatCursor.NextFeature
    Loop

    pWorkspaceEdit.StopEditOperation
        pWorkspaceEdit.StopEditing(True)

    m_Map.ClearSelection
    Dim PMx As IMxDocument
     PMx = m_App.Document
    PMx.ActiveView.Refresh

     pFLayerE = Nothing
     pFeat = Nothing
     pFeat2 = Nothing
     pNextFeat = Nothing
     pNewFeat = Nothing

        WriteLogLine("FINISHED dissolveTransEdgesNoX at " & Now())
        WriteLogLine("Same dir merges=" & mergeCountIJ & "; Opp dir edges merges=" & mergeCountJI)

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





Public Sub dissolveTransEdges2_0()
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
   pTblMode = get_TableClass(m_layers(2)) 'modeAttributes
   pTolls = get_TableClass(m_layers(9)) 'modeTolls
  Dim pTurn As IFeatureLayer
   pTurn = get_FeatureLayer(m_layers(12)) 'TurnMovements
  Dim tblTSeg As ITable
   tblTSeg = get_TableClass(m_layers(14)) 'tbltransitsegments

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
  Dim fcountI, fcountIJ, mergeCountIJ, mergeCountJI As Long

  Dim match As Boolean
  match = True
  Dim edgesdelete As Boolean
  edgesdelete = False

  Dim newI As Long, newJ As Long

  Dim pWorkspaceEdit As IWorkspaceEdit, pWorkspaceEdit2 As IWorkspaceEdit
   pWorkspaceEdit = g_FWS
        pWorkspaceEdit.StartEditing(False)
  pWorkspaceEdit.StartEditOperation

  'this one just loops thru junct layer- should be faster
   pFCj = m_junctShp.Search(Nothing, False)
   pJunctFeat = pFCj.NextFeature

  Dim lSFld As Integer, dnode As Long, idIndex As Integer
  Dim lIndex As Integer, upIndex As Integer
  Dim idxN As Integer, idxJ As Integer
  Dim idxMode As Integer, idxLT As Integer, idxFT As Integer
  Dim idxJunctID As Integer, idxEdgeID As Integer
  Dim ifld As Integer
  idxJunctID = pJunctFeat.Fields.FindField("PSRCJunctID")
  '
  '[122305]pan--revised field names for current SDE layer match
  '
  For ifld = 0 To m_edgeShp.Fields.FieldCount - 1
    If m_edgeShp.Fields.field(ifld).Name = "PSRCEDGEID" Then
      idxEdgeID = ifld
    ElseIf m_edgeShp.Fields.field(ifld).Name = "PSRC_E2ID" Then
      idIndex = ifld
    ElseIf m_edgeShp.Fields.field(ifld).Name = "Updated1" Then
      upIndex = ifld
    ElseIf m_edgeShp.Fields.field(ifld).Name = "INODE" Then
      idxN = ifld
     'MsgBox idxN
    ElseIf m_edgeShp.Fields.field(ifld).Name = "JNODE" Then
      idxJ = ifld
      'MsgBox idxJ
    ElseIf m_edgeShp.Fields.field(ifld).Name = "MODES" Then
      idxMode = ifld
    ElseIf m_edgeShp.Fields.field(ifld).Name = "LINKTYPE" Then
      idxLT = ifld
    ElseIf m_edgeShp.Fields.field(ifld).Name = "FACILITYTYPE" Then
      idxFT = ifld
   End If
  Next ifld

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
  pCounter = 1
  Dim pFeat As IFeature

  Do Until pJunctFeat Is Nothing
    If pCounter = 500 Then
                WriteLogLine("pCounter reaches 500")
    '[011806] pan pCounter appears to serve no purpose
       pWorkspaceEdit.StopEditOperation
                pWorkspaceEdit.StopEditing(True)
       pCounter = 1
                WriteLogLine("saved edits" & "  " & Now)
                pWorkspaceEdit.StartEditing(False)
       pWorkspaceEdit.StartEditOperation
    End If
    pCounter = pCounter + 1
    match = True
    edgesdelete = False
    dnode = pJunctFeat.value(idxJunctID)
            WriteLogLine("potential dnode" & CStr(dnode) & "  " & Now)
         pQF2 = New QueryFilter
        pQF2.WhereClause = "PSRCJunctID = " + CStr(dnode) 'this at node in turnmovemnts
          pQF = New QueryFilter
          pFC = Nothing
         pQF.WhereClause = "INode = " + CStr(dnode) + " Or " + "JNode = " + CStr(dnode)
         fcountIJ = m_edgeShp.featurecount(pQF)
         'WriteLogLine "number of edges" + CStr(fcountIJ)
        If (fcountIJ = 2 And pTurn.FeatureClass.featurecount(pQF2) = 0) Then 'then only at most two edges coming in and this node not part of a turn movement
          'now check if junction really real
           pFC = m_edgeShp.Search(pQF, False)
             pFeat = pFC.NextFeature
            For i = 0 To 1
              If i = 0 Then
                 pFeat2 = pFeat
             End If
             If i = 1 Then
                 pNextFeat = pFeat
             End If
              pFeat = pFC.NextFeature
           Next i

             dir = False
            If pFeat2.value(idxJ) = dnode Then
              If pNextFeat.value(idxJ) = dnode Then 'then deleting both jnodes
                newI = pFeat2.value(idxN)
                newJ = pNextFeat.value(idxN)
              Else 'then deleting jnode of first edge and inode of second
                newJ = pNextFeat.value(idxJ)

                newI = pFeat2.value(idxN)
                dir = True
              End If

            Else
              newJ = pFeat2.value(idxJ)

             If pNextFeat.value(idxJ) = dnode Then 'then deleting inode of first edge and jnode of second

                newI = pNextFeat.value(idxN)
                dir = True
              Else 'then deleting both inodes
                 newI = pNextFeat.value(idxJ)
              End If
           End If
        '[122305]pan References to PSRC_E2ID assume no nulls but currently neither
        'sde.PSRC.TransRefEdges nor sde.PSRC.modeAttributes have this field.

           pQFt = New QueryFilter
          '[122305] pan--idIndex is for PSRC_E2ID
          pQFt.WhereClause = "PSRCEDGEID = " + CStr(pFeat2.value(idxEdgeID))
          'pQFt.WhereClause = "PSRC_E2ID = " + CStr(pFeat2.value(idIndex))
           pTC = pTblMode.Search(pQFt, False)
          'count the presence of records using this edge in modeAttributes, modeTolls, and tblTransitSegments
          If (pTblMode.rowcount(pQFt) > 0 And pTolls.rowcount(pQFt) = 0) Then
            pQFt = New QueryFilter
           '[122305] pan--idIndex is for PSRC_E2ID
          pQFt.WhereClause = "PSRCEDGEID = " + CStr(pFeat2.value(idxEdgeID)) '+ " And change > 0"
           'pQFt.WhereClause = "PSRC_E2ID = " + CStr(pFeat2.value(idIndex)) '+ " And change > 0"
           If (tblTSeg.rowcount(pQFt) = 0) Then 'found first edge in modeAttributes, NOT in modeTolls, NOT in transitSegs
'[060705] pre-Transit Xchk code: If (pTblMode.rowcount(pQFt) > 0 And pTolls.rowcount(pQFt) = 0) Then 'found first edge in modeAttributes and NOT in modeTolls
                        WriteLogLine("no mode, tools, transit  " & Now)
             pRow = pTC.NextRow
             pQFt = New QueryFilter
            '[122305] pan--idIndex is for PSRC_E2ID
            pQFt.WhereClause = "PSRCEDGEID = " + CStr(pNextFeat.value(idxEdgeID))
            'pQFt.WhereClause = "PSRC_E2ID = " + CStr(pNextFeat.value(idIndex))
             pTC = pTblMode.Search(pQFt, False)
            If (pTblMode.rowcount(pQFt) > 0 And pTolls.rowcount(pQFt) = 0) Then
               pQFt = New QueryFilter
              '[122305] pan--idIndex is for PSRC_E2ID
              pQFt.WhereClause = "PSRCEDGEID = " + CStr(pNextFeat.value(idxEdgeID)) '+ " And change > 0"
               'pQFt.WhereClause = "PSRC_E2ID = " + CStr(pNextFeat.value(idIndex)) '+ " And change > 0"
            If (tblTSeg.rowcount(pQFt) = 0) Then 'found second edge in modeAttributes, NOT in modeTolls, NOT in transitSegs
'[060705] pre-transit Xchk code:  If (pTblMode.rowcount(pQFt) > 0 And pTolls.rowcount(pQFt) = 0) Then 'found second edge in modeAttributes and NOT in modeTolls
               pNextRow = pTC.NextRow
                                WriteLogLine("2 no mode, tools, transit" & Now)
              'compare attributes to detect if we have "real" pseudonode
              If dir = True Then
                For i = 1 To pRow.Fields.FieldCount - 1
                  If (pRow.Fields.field(i).Name <> "PSRCEDGEID" And pRow.Fields.field(i).Name <> "IJFFS" And pRow.Fields.field(i).Name <> "JIFFS") Then
                    If (pRow.value(i) <> pNextRow.value(i)) Then
                      match = False
                                                WriteLogLine(pRow.Fields.Field(i).Name + "diff")
                    End If
                  End If
                Next i
              Else 'need to compare IJ to JI to match direction
                Dim j As Integer, l As Integer, t As Integer
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
                  If (pRow.value(index2) <> pNextRow.value(indexNext)) Then
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
                  If (pRow.value(index2) <> pNextRow.value(indexNext)) Then
                      match = False
                                            WriteLogLine("dir vdf diff")
                   End If
                  index2 = pRow.Fields.FindField(dir2 + "SideWalks")
                  indexNext = pNextRow.Fields.FindField(dirNext + "SideWalks")
                  If (pRow.value(index2) <> pNextRow.value(indexNext)) Then
                      match = False
                                            WriteLogLine("dir side diff")
                  End If
                  index2 = pRow.Fields.FindField(dir2 + "BikeLanes")
                  indexNext = pNextRow.Fields.FindField(dirNext + "BikeLanes")
                  If (pRow.value(index2) <> pNextRow.value(indexNext)) Then
                     match = False
                                            WriteLogLine("dir bike diff")
                  End If
                  For l = 0 To 3
                    If l > 1 Then
                      index2 = pRow.Fields.FindField(dir2 + "Lanes" + lType(l))
                      indexNext = pNextRow.Fields.FindField(dirNext + "Lanes" + lType(l))
                      If (pRow.value(index2) <> pNextRow.value(indexNext)) Then
                        match = False
                                                    WriteLogLine("dir lanes diff")
                      End If
                    Else
                       index2 = pRow.Fields.FindField(dir2 + "LaneCap" + lType(l))
                       indexNext = pNextRow.Fields.FindField(dirNext + "LaneCap" + lType(l))
                      If (pRow.value(index2) <> pNextRow.value(indexNext)) Then
                        match = False
                                                    WriteLogLine("dir lanecap diff")
                      End If
                      For t = 0 To 4
                        index2 = pRow.Fields.FindField(dir2 + "Lanes" + lType(l) + timePd(t))
                        indexNext = pNextRow.Fields.FindField(dirNext + "Lanes" + lType(l) + timePd(t))

                        If (pRow.value(index2) <> pNextRow.value(indexNext)) Then
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

              If pFeat2.value(upIndex) <> pNextFeat.value(upIndex) Then
                  match = False
                                    WriteLogLine("update diff")
              End If

              If pFeat2.value(idxMode) <> pNextFeat.value(idxMode) Then
                  match = False
                                    WriteLogLine("modes diff")
              End If

              If pFeat2.value(idxLT) <> pNextFeat.value(idxLT) Then
                  match = False
                                    WriteLogLine("linktype diff")
              End If
              i = pFeat2.Fields.FindField("FacilityTy")
              If pFeat2.value(idxFT) <> pNextFeat.value(idxFT) Then
                  match = False
                                    WriteLogLine("fac diff")
              End If

              If (match = True) Then
                                    WriteLogLine("found match " & Now)
                 pGeometryBag = New GeometryBag
                 pGeometry = pFeat2.ShapeCopy
                                    pGeometryBag.AddGeometry(pGeometry)
                 pGeometry = pNextFeat.ShapeCopy
                                    pGeometryBag.AddGeometry(pGeometry)

                'update the merge count for the log
                If dir Then
                  mergeCountIJ = mergeCountIJ + 1
                Else
                  mergeCountJI = mergeCountJI + 1
                End If

                'merge the edge features
                 pPolyline = New Polyline
                 pTopo = pPolyline
                                    pTopo.ConstructUnion(pGeometryBag)
                 pNewFeat = m_edgeShp.CreateFeature
                 pNewFeat.Shape = pPolyline
                pNewFeat.Store
                 pFlds = pNewFeat.Fields
                For i = 1 To pFlds.FieldCount - 1
                  lIndex = pFeat2.Fields.FindField(pFlds.field(i).Name)
                  If (lIndex <> -1) Then
                                            If (pFlds.Field(i).Type = esriFieldType.esriFieldTypeOID And Not pFlds.Field(i).Type = esriFieldType.esriFieldTypeGeometry Or pFlds.Field(i).Name <> "Shape") Then
                                                'inode and jnode need to be  differently
                                                If (pFlds.Field(i).Name = "INODE") Then
                                                    pNewFeat.Value(i) = newI
                                                ElseIf (pFlds.Field(i).Name = "JNODE") Then
                                                    pNewFeat.Value(i) = newJ
                                                    '[122905] pan--SHAPE.len is internal field not editable
                                                    'ElseIf (pFlds.field(i).name = "SHAPE.len") Then
                                                    '  pnewfeat.value(i) = pPolyline.length
                                                    'Else
                                                    '  pnewfeat.value(i) = pFeat2.value(lindex)
                                                End If
                                            End If
                  End If
                Next i
                'pan Edit session incompatible with Store
                'pan 12-16-05 re-inserted store and edge delete
                pNewFeat.Store
                edgesdelete = True
                pFeat2.Delete
                pNextFeat.Delete
                'find junction node and delete it


                  If fVerboseLog Then
                    '[122905] pan--I changed check to gt because want to delete this junction
                    'If idxJunctID < 0 Then
                    If idxJunctID > 0 Then
                                            WriteLogLine("Deleting intrmed. feature junct OID=" & pJunctFeat.OID)
                    Else
                                            WriteLogLine("Deleting intrmed. feature junctID=" & pJunctFeat.Value(idxJunctID))
                    End If
                  End If
                  pJunctFeat.Delete

                'MsgBox pSSet.count
                End If
              End If  'match = true

             End If 'if has tblseg next
            End If 'if has tbltoll next
            End If 'if has tblseg
          End If 'if has tbltoll
     '[122905] pan--this already deleted above, this is duplicate
     '[011906] pan --this may be needed after all
      'If (edgesdelete = True) Then
        'pFeat2.Delete
        'pNextFeat.Delete
        'edgesdelete = False
            'WriteLogLine "edges deleted " & now
      'End If
     pJunctFeat = pFCj.NextFeature
     pQF = Nothing
     pQF2 = Nothing
     pQFt = Nothing
     pTC = Nothing
     pFC = Nothing
     pFC2 = Nothing

  Loop

  pWorkspaceEdit.StopEditOperation
        pWorkspaceEdit.StopEditing(True)

  m_Map.ClearSelection
  Dim PMx As IMxDocument
   PMx = m_App.Document
  PMx.ActiveView.Refresh

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
    tblTSeg = Nothing
    pWorkspaceEdit = Nothing


        WriteLogLine("FINISHED dissolveTransEdges2 at " & Now())
        WriteLogLine("Same dir merges= " & mergeCountIJ & "; Opp dir edges merges= " & mergeCountJI)
        WriteLogLine(m_edgeShp.AliasName + CStr(m_edgeShp.FeatureCount(Nothing)))
  'Note: needs to be a round two version of this based on atrributes that don't match but are not significant
  'need to get these attributes!
Exit Sub

ErrChk:
        CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.dissolveTransEdges2")
        MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.dissolveTransEdges2")
  'Print #1, Err.Description, vbInformation, "dissolveTransEdges2"

  If pWorkspaceEdit.IsBeingEdited Then
    pWorkspaceEdit.StopEditOperation
            pWorkspaceEdit.StopEditing(True)
  End If


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
    tblTSeg = Nothing
    pWorkspaceEdit = Nothing

End Sub


Public Function fAttribMatch(rowOne As IRow, rowTwo As IRow, fSameDir As Boolean)
  'tests the modeAttributes values of the two separate rows passed
  'returns TRUE if the rows have identical attributes, otherwise FALSE
  'fSameDir is TRUE if the IJ direction for rowOne matches the IJ dir for rowTwo

  'on error GoTo eh

  Dim index As Integer
  Dim intHold1, intHold2 As Integer
  Dim dblHold1, dblHold2 As Double

  fAttribMatch = True

  If fSameDir Then
    'check IJ vs IJ

    intHold1 = rowOne.Fields.FindField("IJLanesGPAM")
    If rowOne.value(intHold1) <> rowTwo.value(intHold1) Then
      'asymmetric!
      fAttribMatch = False
      Exit Function
    End If
    intHold1 = rowOne.Fields.FindField("IJLanesGPMD")
    If rowOne.value(intHold1) <> rowTwo.value(intHold1) Then
      'asymmetric!
      fAttribMatch = False
      Exit Function
    End If
    intHold1 = rowOne.Fields.FindField("IJLanesGPPM")
    If rowOne.value(intHold1) <> rowTwo.value(intHold1) Then
      'asymmetric!
      fAttribMatch = False
      Exit Function
    End If
    intHold1 = rowOne.Fields.FindField("IJLanesGPEV")
    If rowOne.value(intHold1) <> rowTwo.value(intHold1) Then
      'asymmetric!
      fAttribMatch = False
      Exit Function
    End If
    intHold1 = rowOne.Fields.FindField("IJLanesGPNI")
    If rowOne.value(intHold1) <> rowTwo.value(intHold1) Then
      'asymmetric!
      fAttribMatch = False
      Exit Function
    End If

    intHold1 = rowOne.Fields.FindField("IJLanesHOVAM")
    If rowOne.value(intHold1) <> rowTwo.value(intHold1) Then
      'asymmetric!
      fAttribMatch = False
      Exit Function
    End If
    intHold1 = rowOne.Fields.FindField("IJLanesHOVMD")
    If rowOne.value(intHold1) <> rowTwo.value(intHold1) Then
      'asymmetric!
      fAttribMatch = False
      Exit Function
    End If
    intHold1 = rowOne.Fields.FindField("IJLanesHOVPM")
    If rowOne.value(intHold1) <> rowTwo.value(intHold1) Then
      'asymmetric!
      fAttribMatch = False
      Exit Function
    End If
    intHold1 = rowOne.Fields.FindField("IJLanesHOVEV")
    If rowOne.value(intHold1) <> rowTwo.value(intHold1) Then
      'asymmetric!
      fAttribMatch = False
      Exit Function
    End If
    intHold1 = rowOne.Fields.FindField("IJLanesHOVNI")
    If rowOne.value(intHold1) <> rowTwo.value(intHold1) Then
      'asymmetric!
      fAttribMatch = False
      Exit Function
    End If

    intHold1 = rowOne.Fields.FindField("IJSpeedLimit")
    If rowOne.value(intHold1) <> rowTwo.value(intHold1) Then
      'asymmetric!
      fAttribMatch = False
      Exit Function
    End If

    intHold1 = rowOne.Fields.FindField("IJFFS")
    If rowOne.value(intHold1) <> rowTwo.value(intHold1) Then
      'asymmetric!
      fAttribMatch = False
      Exit Function
    End If

    intHold1 = rowOne.Fields.FindField("IJVDFunc")
    If rowOne.value(intHold1) <> rowTwo.value(intHold1) Then
      'asymmetric!
      fAttribMatch = False
      Exit Function
    End If

    intHold1 = rowOne.Fields.FindField("IJLaneCapGP")
    If rowOne.value(intHold1) <> rowTwo.value(intHold1) Then
      'asymmetric!
      fAttribMatch = False
      Exit Function
    End If

    intHold1 = rowOne.Fields.FindField("IJLaneCapHOV")
    If rowOne.value(intHold1) <> rowTwo.value(intHold1) Then
      'asymmetric!
      fAttribMatch = False
      Exit Function
    End If

    intHold1 = rowOne.Fields.FindField("IJSidewalks")
    If rowOne.value(intHold1) <> rowTwo.value(intHold1) Then
      'asymmetric!
      fAttribMatch = False
      Exit Function
    End If

    intHold1 = rowOne.Fields.FindField("IJBikelanes")
    If rowOne.value(intHold1) <> rowTwo.value(intHold1) Then
      'asymmetric!
      fAttribMatch = False
      Exit Function
    End If

    intHold1 = rowOne.Fields.FindField("IJLanesTR")
    If rowOne.value(intHold1) <> rowTwo.value(intHold1) Then
      'asymmetric!
      fAttribMatch = False
      Exit Function
    End If
    intHold1 = rowOne.Fields.FindField("IJLanesTK")
    If rowOne.value(intHold1) <> rowTwo.value(intHold1) Then
      'asymmetric!
      fAttribMatch = False
      Exit Function
    End If
  Else
    'check JI vs IJ
    intHold1 = rowOne.Fields.FindField("IJLanesGPAM")
    intHold2 = rowTwo.Fields.FindField("JILanesGPAM")
    If rowOne.value(intHold1) <> rowTwo.value(intHold2) Then
      'asymmetric!
      fAttribMatch = False
      Exit Function
    End If
    intHold1 = rowOne.Fields.FindField("IJLanesGPMD")
    intHold2 = rowTwo.Fields.FindField("JILanesGPMD")
    If rowOne.value(intHold1) <> rowTwo.value(intHold2) Then
      'asymmetric!
      fAttribMatch = False
      Exit Function
    End If
    intHold1 = rowOne.Fields.FindField("IJLanesGPPM")
    intHold2 = rowTwo.Fields.FindField("JILanesGPPM")
    If rowOne.value(intHold1) <> rowTwo.value(intHold2) Then
      'asymmetric!
      fAttribMatch = False
      Exit Function
    End If
    intHold1 = rowOne.Fields.FindField("IJLanesGPEV")
    intHold2 = rowTwo.Fields.FindField("JILanesGPEV")
    If rowOne.value(intHold1) <> rowTwo.value(intHold2) Then
      'asymmetric!
      fAttribMatch = False
      Exit Function
    End If
    intHold1 = rowOne.Fields.FindField("IJLanesGPNI")
    intHold2 = rowTwo.Fields.FindField("JILanesGPNI")
    If rowOne.value(intHold1) <> rowTwo.value(intHold2) Then
      'asymmetric!
      fAttribMatch = False
      Exit Function
    End If

    intHold1 = rowOne.Fields.FindField("IJLanesHOVAM")
    intHold2 = rowTwo.Fields.FindField("JILanesHOVAM")
    If rowOne.value(intHold1) <> rowTwo.value(intHold2) Then
      'asymmetric!
      fAttribMatch = False
      Exit Function
    End If
    intHold1 = rowOne.Fields.FindField("IJLanesHOVMD")
    intHold2 = rowTwo.Fields.FindField("JILanesHOVMD")
    If rowOne.value(intHold1) <> rowTwo.value(intHold2) Then
      'asymmetric!
      fAttribMatch = False
      Exit Function
    End If
    intHold1 = rowOne.Fields.FindField("IJLanesHOVPM")
    intHold2 = rowTwo.Fields.FindField("JILanesHOVPM")
    If rowOne.value(intHold1) <> rowTwo.value(intHold2) Then
      'asymmetric!
      fAttribMatch = False
      Exit Function
    End If
    intHold1 = rowOne.Fields.FindField("IJLanesHOVEV")
    intHold2 = rowTwo.Fields.FindField("JILanesHOVEV")
    If rowOne.value(intHold1) <> rowTwo.value(intHold2) Then
      'asymmetric!
      fAttribMatch = False
      Exit Function
    End If
    intHold1 = rowOne.Fields.FindField("IJLanesHOVNI")
    intHold2 = rowTwo.Fields.FindField("JILanesHOVNI")
    If rowOne.value(intHold1) <> rowTwo.value(intHold2) Then
      'asymmetric!
      fAttribMatch = False
      Exit Function
    End If

    intHold1 = rowOne.Fields.FindField("IJSpeedLimit")
    intHold2 = rowTwo.Fields.FindField("JISpeedLimit")
    If rowOne.value(intHold1) <> rowTwo.value(intHold2) Then
      'asymmetric!
      fAttribMatch = False
      Exit Function
    End If

    intHold1 = rowOne.Fields.FindField("IJFFS")
    intHold2 = rowTwo.Fields.FindField("JIFFS")
    If rowOne.value(intHold1) <> rowTwo.value(intHold2) Then
      'asymmetric!
      fAttribMatch = False
      Exit Function
    End If

    intHold1 = rowOne.Fields.FindField("IJVDFunc")
    intHold2 = rowTwo.Fields.FindField("JIVDFunc")
    If rowOne.value(intHold1) <> rowTwo.value(intHold2) Then
      'asymmetric!
      fAttribMatch = False
      Exit Function
    End If

    intHold1 = rowOne.Fields.FindField("IJLaneCapGP")
    intHold2 = rowTwo.Fields.FindField("JILaneCapGP")
    If rowOne.value(intHold1) <> rowTwo.value(intHold2) Then
      'asymmetric!
      fAttribMatch = False
      Exit Function
    End If

    intHold1 = rowOne.Fields.FindField("IJLaneCapHOV")
    intHold2 = rowTwo.Fields.FindField("JILaneCapHOV")
    If rowOne.value(intHold1) <> rowTwo.value(intHold2) Then
      'asymmetric!
      fAttribMatch = False
      Exit Function
    End If

    intHold1 = rowOne.Fields.FindField("IJSidewalks")
    intHold2 = rowTwo.Fields.FindField("JISidewalks")
    If rowOne.value(intHold1) <> rowTwo.value(intHold2) Then
      'asymmetric!
      fAttribMatch = False
      Exit Function
    End If

    intHold1 = rowOne.Fields.FindField("IJBikelanes")
    intHold2 = rowTwo.Fields.FindField("JIBikelanes")
    If rowOne.value(intHold1) <> rowTwo.value(intHold2) Then
      'asymmetric!
      fAttribMatch = False
      Exit Function
    End If

    intHold1 = rowOne.Fields.FindField("IJLanesTR")
    intHold2 = rowTwo.Fields.FindField("JILanesTR")
    If rowOne.value(intHold1) <> rowTwo.value(intHold2) Then
      'asymmetric!
      fAttribMatch = False
      Exit Function
    End If
    intHold1 = rowOne.Fields.FindField("IJLanesTK")
    intHold2 = rowTwo.Fields.FindField("JILanesTK")
    If rowOne.value(intHold1) <> rowTwo.value(intHold2) Then
      'asymmetric!
      fAttribMatch = False
      Exit Function
    End If
  End If
  Exit Function

eh:

        CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.fAttribMatch")
        MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.fAttribMatch")

'  MsgBox "Error " & Err.Number & ", " & Err.Description, , "GlobalMod.fAttribMatch"

End Function

    Public Sub createWeaveLink0(ByVal lType As String, ByVal pFeat As IFeature, ByVal pRow1 As IRow, ByVal ptstring As String, _
    ByVal wString As String, ByVal wNodes As String, ByVal astring As String, ByVal tpd As String, ByVal lanes As Long, ByVal dir As String, _
    ByVal dctEmme2Nodes As Dictionary(Of Object, Object), ByVal dctWeaveNodes As Dictionary(Of Long, Long), ByVal dctWeaveNodes2 As Dictionary(Of Object, Object), _
    ByVal dctJcts As Dictionary(Of Long, Feature), ByVal dctEdges As Dictionary(Of Long, Feature))
        'dctWeaveNodes is a dictionary of weave node IDs. key=cstr(weavenode id), item=weavenode id
        'dctWeaveNodes2 is a dictionary of Scen_Node and the corresponding weave node. key=Scen_Node, item=weave node

        'ptstring: string for weave nodes
        'wstring: string for weave links
        'astring: string for attributes

        'on error GoTo eh
        Const wID_Type = " 0 0"
        If fVerboseLog Then WriteLogLine("createWeaveLink called with ptstring=" & ptstring)
        If fVerboseLog Then WriteLogLine("feat id" & pFeat.OID)
        If fVerboseLog Then WriteLogLine("-------------------------------------------------")

        Dim pQF As IQueryFilter, pQFt As IQueryFilter
        Dim dctWNs As Dictionary(Of Long, Long) = dctWeaveNodes
        Dim dctWNs2 As Dictionary(Of Object, Object) = dctWeaveNodes2

        Dim pFC As IFeatureCursor, pFCj As IFeatureCursor
        Dim pNewFeat As IFeature
        Dim pefeat As IFeature, pJFeat As IFeature
        Dim pGeom As IGeometry
        Dim pPt As IPoint, pOpt As IPoint
        Dim pFilter As ISpatialFilter
        Dim pFeatSelect As IFeatureSelection
        Dim pPoint As IPoint
        Dim pOpoint As IPoint
        Dim pXpoint As IPoint

        Dim iNode As Long, iwNode As Long
        Dim JNode As Long, jwNode As Long
        Dim iIndex As Integer, iwIndex As Integer
        Dim jIndex As Integer, jwIndex As Integer
        Dim index As Integer
        Dim wDefault As String
        Dim strFFT As String    'holds formatted FFT number

        Dim pTable As ITable
        Dim pTC As ICursor
        Dim pRow As IRow
        Dim pWS As IWorkspace
        'Dim pGeom As IGeometry
        pWS = get_Workspace()

        Dim modeWeave As String
        modeWeave = "ahijstuvbw"
        Dim timePd As String
        timePd = tpd
        Dim laneFld As String
        Dim pWSedit As IWorkspaceEdit
        Dim pNString As String
        'Dim dctWeaveNodes? As Dictionary
        '[071206] - pan shadow link node ids  to unique values with offset
        'of largestwjunct. Public largestjunct has been incremented in createShapefiles and
        'during splits, but GetLargestID used anyway to find top of stack.

        'GetLargestID m_junctShp, "PSRCJunctID", largestjunct
        'largestwjunct = largestjunct + 1

        '[050407] hyu
        Dim nodeID As Long
        If pFeat.OID = 1006 Then MsgBox(1)

        iIndex = pFeat.Fields.FindField("INode")
        jIndex = pFeat.Fields.FindField("JNode")
        iNode = pFeat.Value(iIndex)
        JNode = pFeat.Value(jIndex)
        'now make these into their scen_node ids
        index = pFeat.Fields.FindField("UseEmmeN")

        '[021706] hyu: shouldn't UseemmeN=3 means both ij nodes are emme2id?
        '    If pFeat.value(index) = 3 Then
        If pFeat.Value(index) = 0 Or IsDBNull(pFeat.Value(index)) Then
            iNode = iNode + m_Offset
            JNode = JNode + m_Offset
        ElseIf pFeat.Value(index) = 1 Then 'inode is emme2id
            JNode = JNode + m_Offset
            '[042506] hyu
            iNode = dctEmme2Nodes.Item(CStr(iNode))
        ElseIf pFeat.Value(index) = 2 Then 'jnode is emme2id
            iNode = iNode + m_Offset
            JNode = dctEmme2Nodes.Item(CStr(JNode))
        End If
        If fVerboseLog Then WriteLogLine("linkij" + CStr(iNode) + " " + CStr(JNode))
        'if equal 3 then both ia nd j are emmeid's and are ok for scen_node values
        Dim tempnode As Long
        Dim i As Integer, j As Integer
        Dim Leng As Double
        Dim facType As Integer
        index = pFeat.Fields.FindField("FacilityType")
        facType = pFeat.Value(index)
        If fVerboseLog Then WriteLogLine("factype" + CStr(facType))
        index = pFeat.Fields.FindField("LinkType")
        'now assign defaults to weave: length, modes, link type, lanes
        Leng = getWeaveLen(lType)
        If fVerboseLog Then WriteLogLine("leng of weave" + CStr(Leng))
        If Not IsDBNull(pFeat.Value(index)) And pFeat.Value(index) > 0 Then
            wDefault = wDefault + " " + CStr(Leng) + " " + modeWeave + " " + CStr(pFeat.Value(index)) + " " + CStr(lanes)
        Else
            wDefault = wDefault + " " + CStr(Leng) + " " + modeWeave + " 90 " + CStr(lanes)
        End If

        'vdf, capacity, FFT, facility type
        Dim ijtemp As String, jitemp As String
        Dim direction As String
        ijtemp = "1 1800 60 0"
        jitemp = "1 1800 60 0"

        iwIndex = pFeat.Fields.FindField(lType + "_I")
        jwIndex = pFeat.Fields.FindField(lType + "_J")

        Debug.Print(pFeat.OID)
        If (Not IsDBNull(pFeat.Value(iwIndex)) And pFeat.Value(iwIndex) > 0) Then 'jnode should also have a value
            iwNode = pFeat.Value(iwIndex)
            jwNode = pFeat.Value(jwIndex)

            '[051006] hyu: get weavenode, weave links.
            '        If dir = "IJ" Then wnodes = CStr(iwNode) + " " + CStr(jwNode)
            '        If dir = "JI" Then wnodes = CStr(jwNode) + " " + CStr(iwNode)
            If fVerboseLog Then WriteLogLine("have weave' " + wNodes)

            If dir = "IJ" Then
                wNodes = CStr(iwNode) + " " + CStr(jwNode)
            Else
                wNodes = CStr(jwNode) + " " + CStr(iwNode)
            End If

            If Not dctWeaveNodes.ContainsKey(CStr(iwNode)) Then
                '[050407] hyu:
                '             pQF = New QueryFilter
                '            pQF.WhereClause = "Scen_Node = " + CStr(iNode)
                '             pFCj = m_junctShp.Search(pQF, False)
                '             pJFeat = pFCj.NextFeature

                '[051507]hyu: if the weavenode is a junction, it's the split point for case 2 during weave link creation
                If dctJcts.ContainsKey(CStr(iwNode)) Then
                    pJFeat = dctJcts.Item(iwNode)
                    pPoint = pJFeat.ShapeCopy
                Else
                    pJFeat = dctJcts.Item(iNode)
                    pPoint = getWeavenode(pJFeat, lType)
                    ptstring = "a " + CStr(iwNode) + " " + CStr(pPoint.X) + " " + CStr(pPoint.Y) + " 0 0 0 0"
                End If


                wString = "a " + CStr(iwNode) + " " + CStr(iNode) + " " + wDefault + " " + ijtemp
                astring = CStr(iwNode) + " " + CStr(iNode) + wID_Type

                wString = wString + vbCrLf + "a " + CStr(iNode) + " " + CStr(iwNode) + " " + wDefault + " " + ijtemp
                astring = astring + vbCrLf + CStr(iNode) + " " + CStr(iwNode) + wID_Type

                dctWeaveNodes.Add(CStr(iwNode), iwNode)
                If Not dctWeaveNodes2.ContainsKey(iNode) Then dctWeaveNodes2.Add(iNode, iwNode)
            End If

            If Not dctWeaveNodes.ContainsKey(CStr(jwNode)) Then
                '[050407]hyu:
                '             pQF = New QueryFilter
                '            pQF.WhereClause = "Scen_Node = " + CStr(jNode)
                '             pFCj = m_junctShp.Search(pQF, False)
                '             pJFeat = pFCj.NextFeature

                '[051507]hyu: if the weavenode is a junction, it's the split point for case 2 during weave link creation
                If dctJcts.ContainsKey(CStr(jwNode)) Then
                    pJFeat = dctJcts.Item(jwNode)
                    pPoint = pJFeat.ShapeCopy
                Else
                    pJFeat = dctJcts.Item(JNode)
                    pPoint = getWeavenode(pJFeat, lType)
                    ptstring = ptstring + IIf(ptstring = "", "", vbCrLf) + "a " + CStr(jwNode) + " " + CStr(pPoint.X) + " " + CStr(pPoint.Y) + " 0 0 0 0"

                End If

                ptstring = ptstring + IIf(ptstring = "", "", vbCrLf) + "a " + CStr(jwNode) + " " + CStr(pPoint.X) + " " + CStr(pPoint.Y) + " 0 0 0 0"
                wString = wString + IIf(wString = "", "", vbCrLf) + "a " + CStr(jwNode) + " " + CStr(JNode) + " " + wDefault + " " + jitemp
                astring = astring + IIf(astring = "", "", vbCrLf) + CStr(JNode) + " " + CStr(jwNode) + wID_Type

                wString = wString + vbCrLf + "a " + CStr(JNode) + " " + CStr(jwNode) + " " + wDefault + " " + jitemp
                astring = astring + vbCrLf + CStr(jwNode) + CStr(JNode) + wID_Type

                dctWeaveNodes.Add(CStr(jwNode), jwNode)
                dctWeaveNodes2.Add(JNode, jwNode)
            End If
            'have merge nodes and therefore create weave link
            'so can return now with proper inode and jnode for wltype(i)
            Exit Sub
        Else    'no existing weave node recorded in the attributes
            Dim bWeaveNodeExists As Boolean
            For j = 0 To 1
                bWeaveNodeExists = False
                If j = 0 Then
                    If dctWeaveNodes2.ContainsKey(iNode) Then
                        pFeat.Value(pFeat.Fields.FindField(lType & "_I")) = dctWeaveNodes2.Item(iNode)

                        bWeaveNodeExists = True
                        iwNode = dctWeaveNodes2.Item(iNode)
                    End If
                Else
                    If dctWeaveNodes2.ContainsKey(JNode) Then
                        pFeat.Value(pFeat.Fields.FindField(lType & "_J")) = dctWeaveNodes2.Item(JNode)
                        bWeaveNodeExists = True
                        jwNode = dctWeaveNodes.Item(JNode)
                    End If
                End If
                pFeat.Store()
                '[052007] hyu: per discuss w/ Jeff & Andy, temperarily we assume one base node can only correspond to one weave node.
                'so here we will use the existing weave node.
                If bWeaveNodeExists = True Then

                    If j = 0 Then
                        If fVerboseLog Then WriteLogLine("weave node existing, inode")
                        If dir = "IJ" Then
                            wNodes = CStr(iwNode)
                        Else
                            wNodes = CStr(jwNode)
                        End If
                    Else
                        If fVerboseLog Then WriteLogLine("weave node existing, jnode")
                        If Not ptstring = "" Then ptstring = ptstring + vbCrLf
                        If Not wString = "" Then wString = wString + vbCrLf
                        If Not astring = "" Then astring = astring + vbCrLf
                        If dir = "IJ" Then
                            wNodes = wNodes + " " + CStr(jwNode)
                        Else
                            wNodes = wNodes + " " + CStr(iwNode)
                        End If
                    End If
                    If fVerboseLog Then WriteLogLine(wNodes)

                Else    'bWeaveNodeExists = False
                    If fVerboseLog Then WriteLogLine("for looping " + CStr(j))
                    'first need to check if connected edges have weave i,j's
                    pQF = New QueryFilter
                    If j = 0 Then
                        If dir = "IJ" Then

                            pQF.WhereClause = "Scen_Node = " + CStr(iNode)
                            tempnode = iNode

                            '[050407] hyu
                            pJFeat = dctJcts.Item(iNode)
                        Else
                            pQF.WhereClause = "Scen_Node = " + CStr(JNode)
                            tempnode = JNode

                            '[050407] hyu
                            pJFeat = dctJcts.Item(JNode)
                        End If
                    ElseIf j = 1 Then
                        If dir = "IJ" Then
                            pQF.WhereClause = "Scen_Node = " + CStr(JNode)
                            tempnode = JNode
                            '[050407]hyu
                            pJFeat = dctJcts.Item(JNode)
                        Else
                            pQF.WhereClause = "Scen_Node = " + CStr(iNode)
                            tempnode = iNode
                            '[050407] hyu
                            pJFeat = dctJcts.Item(iNode)
                        End If
                    End If
                    '             pFCj = m_junctShp.Search(pQF, False)
                    '             pJFeat = pFCj.NextFeature 'should be just one since its a unique id


                    pGeom = pJFeat.ShapeCopy
                    pOpt = pGeom
                    If fVerboseLog Then WriteLogLine("createWeaveLink new point")
                    pFilter = New SpatialFilter
                    With pFilter
                        .Geometry = pOpt 'all the edges in intermediate shp
                        .GeometryField = m_edgeShp.ShapeFieldName
                        .SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects
                    End With
                    'now get the opposite point- used in calculating angles for intersections
                    pQF = New QueryFilter
                    If j = 0 Then
                        If dir = "IJ" Then
                            pQF.WhereClause = "Scen_Node = " + CStr(JNode)
                            '[050407] hyu
                            pJFeat = dctJcts.Item(JNode)
                        Else
                            pQF.WhereClause = "Scen_Node = " + CStr(iNode)
                            '[050407] hyu
                            pJFeat = dctJcts.Item(iNode)
                        End If
                    ElseIf j = 1 Then
                        If dir = "IJ" Then
                            pQF.WhereClause = "Scen_Node = " + CStr(iNode)
                            '[050407] hyu
                            pJFeat = dctJcts.Item(iNode)
                        Else
                            pQF.WhereClause = "Scen_Node = " + CStr(JNode)
                            '[050407] hyu
                            pJFeat = dctJcts.Item(JNode)
                        End If
                    End If
                    '             pFCj = m_junctShp.Search(pQF, False)
                    '             pJFeat = pFCj.NextFeature 'should be just one since its a unique id
                    pGeom = pJFeat.ShapeCopy
                    pPt = pGeom
                    pJFeat = Nothing
                    If (m_edgeShp.FeatureCount(pFilter) > 2) Then   '[022006] hyu: why >2?
                        pFC = m_edgeShp.Search(pFilter, False)
                        pefeat = pFC.NextFeature
                        Dim pPoint2 As IPoint
                        Dim found As Boolean, fce As Boolean, fie As Boolean
                        Dim bigangle As Double, smangle As Double, tempangle As Double
                        Dim Tindex As Integer
                        fce = False
                        fie = False

                        bigangle = 0
                        smangle = 9999.9

                        found = False
                        If fVerboseLog Then WriteLogLine("createWeaveLink angle false ")
                        Dim pIedge As IFeature, pCedge As IFeature
                        'need to get Iedge and Cedge
                        Dim iLanes As Long, otheri As Long, otherj As Long, clanes As Long
                        iLanes = 0
                        Do Until pefeat Is Nothing

                            If pefeat.OID = pFeat.OID Then
                                '   pefeat = pFC.NextFeature
                                If fVerboseLog Then WriteLogLine("skipping")
                            Else
                                If fVerboseLog Then WriteLogLine("looking at " + CStr(pefeat.OID))
                                index = pefeat.Fields.FindField("UseEmmeN")
                                '[022006] hyu: UseEmmeN=3 means both i j are reserved
                                If IsDBNull(pefeat.Value(index)) Or pefeat.Value(index) = 0 Or pefeat.Value(index) = 1 Then
                                    '                        If pefeat.value(index) = 3 Or pefeat.value(index) = 1 Then
                                    otherj = pefeat.Value(jIndex) + m_Offset
                                Else
                                    '[042506] hyu
                                    'otherj = pefeat.value(jIndex)
                                    otherj = dctEmme2Nodes.Item(CStr(pefeat.Value(jIndex)))
                                End If
                                '[022006] hyu: shouldn't UseEmmeN=3 means both i j has offset
                                If IsDBNull(pefeat.Value(index)) Or pefeat.Value(index) = 0 Or pefeat.Value(index) = 2 Then
                                    '                        If pefeat.value(index) = 3 Or pefeat.value(index) = 2 Then
                                    otheri = pefeat.Value(iIndex) + m_Offset
                                Else
                                    '[042506] hyu
                                    'otheri = pefeat.value(iIndex)
                                    otheri = dctEmme2Nodes.Item(CStr(pefeat.Value(iIndex)))
                                End If
                                If fVerboseLog Then WriteLogLine("this edge ij " + CStr(otheri) + " " + CStr(otherj))

                                If otheri = tempnode Then
                                    '                            pQF.WhereClause = "Scen_Node = " + CStr(otherj)
                                    nodeID = otherj
                                End If
                                If otherj = tempnode Then
                                    '                            pQF.WhereClause = "Scen_Node = " + CStr(otheri)
                                    nodeID = otheri
                                End If

                                'need to skip ones that share the same i and j nodes as pfeat
                                Dim same As Integer
                                same = 0
                                If otheri = iNode Or otheri = JNode Then
                                    same = 1
                                End If
                                If otherj = iNode Or otherj = JNode Then
                                    If same = 0 Then
                                        same = 2
                                    Else
                                        same = 3
                                    End If
                                End If

                                If same < 3 Then    'not the base link
                                    '[050407] hyu
                                    If dctJcts.ContainsKey(nodeID) Then
                                        pJFeat = dctJcts.Item(nodeID)
                                        '                             pFCj = m_junctShp.Search(pQF, False)
                                        '                            If (m_junctShp.featurecount(pQF) > 0) Then
                                        '                                 pJFeat = pFCj.NextFeature 'should be just one since its a unique id
                                        pPoint2 = pJFeat.ShapeCopy 'this x1
                                        tempangle = calAngle(pOpt, pPoint2, pPt)

                                        If fVerboseLog Then WriteLogLine(tempangle)
                                        '[052107] hyu: condition of clanes should happen later. It isn't assigned value yet. _
                                        'isn't  tempangel <45
                                        'If smangle > tempangle And tempangle < 15 And clanes = 0 Then
                                        If smangle > tempangle And tempangle < 45 Then

                                            smangle = tempangle
                                            pCedge = pefeat
                                            '[052107] hyu: fce=true should be moved to aftering clanes getting value
                                            'fce = True
                                            pTable = get_TableClass(m_layers(2)) 'need to put in check to see if need to pull out of tblLine or evt tables
                                            index = pCedge.Fields.FindField("PSRCEdgeID")
                                            pQFt = New QueryFilter
                                            pQFt.WhereClause = "PSRCEdgeID = " + CStr(pCedge.Value(index))
                                            pTC = pTable.Search(pQFt, False)
                                            pRow = pTC.NextRow
                                            If lType = "HOV" Or lType = "GP" Then
                                                laneFld = CStr(lType) + CStr(timePd)
                                            Else
                                                laneFld = CStr(lType)
                                            End If
                                            'want both direction test- weave is by default bi-directional
                                            index = pRow.Fields.FindField("IJLanes" + laneFld)
                                            clanes = pRow.Value(index)
                                            index = pRow.Fields.FindField("JILanes" + laneFld)
                                            clanes = clanes + pRow.Value(index)

                                            If clanes = 0 Then fce = True

                                            If fVerboseLog Then WriteLogLine("clanes" + CStr(clanes))
                                            '[052107] hyu: isn't tempangle>135. iLanes should be concerned later. It hasn't have value yet.
                                            'ElseIf bigangle < tempangle And tempangle < 135 And ilanes = 0 Then 'check if at T junction
                                        ElseIf bigangle < tempangle And tempangle > 135 Then
                                            bigangle = tempangle
                                            pIedge = pefeat
                                            '[052107] hyu: fie=true comes later after it getting value
                                            'fie = True
                                            pTable = get_TableClass(m_layers(2)) 'need to put in check to see if need to pull out of tblLine or evt tables
                                            index = pIedge.Fields.FindField("PSRCEdgeID")
                                            pQFt = New QueryFilter
                                            pQFt.WhereClause = "PSRCEdgeID = " + CStr(pIedge.Value(index))
                                            pTC = pTable.Search(pQFt, False)
                                            pRow = pTC.NextRow

                                            If lType = "HOV" Or lType = "GP" Then
                                                laneFld = CStr(lType) + CStr(timePd)
                                            Else
                                                laneFld = CStr(lType)
                                            End If
                                            index = pRow.Fields.FindField("IJLanes" + laneFld)
                                            iLanes = pRow.Value(index)
                                            index = pRow.Fields.FindField("JILanes" + laneFld)
                                            iLanes = iLanes + pRow.Value(index)
                                            If iLanes = 0 Then fie = True
                                            If fVerboseLog Then WriteLogLine("ilanes" + CStr(iLanes))
                                        End If  'angle
                                    End If  'dctJcts.ContainsKey(nodeID)
                                End If  'same<3
                            End If  'pefeat.OID = pFeat.OID
                            pefeat = pFC.NextFeature
                        Loop    '

                        found = False
                        If fce = True Then
                            If fVerboseLog Then WriteLogLine("in here with fce")

                            If clanes > 0 Then
                                index = pCedge.Fields.FindField("UseEmmeN")
                                If pCedge.Value(index) = 0 Or pCedge.Value(index) = 1 Then
                                    otherj = pCedge.Value(jIndex) + m_Offset
                                Else
                                    '[042506] hyu
                                    'otherj = pCedge.value(jIndex)
                                    otherj = dctEmme2Nodes.Item(CStr(pCedge.Value(jIndex)))
                                End If
                                If pCedge.Value(index) = 0 Or pCedge.Value(index) = 2 Then
                                    otheri = pCedge.Value(iIndex) + m_Offset
                                Else
                                    '[042506] hyu
                                    'otheri = pCedge.value(iIndex)
                                    otheri = dctEmme2Nodes.Item(CStr(pCedge.Value(iIndex)))
                                End If
                                If fVerboseLog Then WriteLogLine("other ij " + CStr(otheri) + " or " + CStr(otherj))
                                If otheri = tempnode Then
                                    If fVerboseLog Then WriteLogLine("fce other i = tempnode")
                                    If (pCedge.Value(iwIndex) > 0) Then 'jnode should also have a value
                                        If j = 0 And dir = "IJ" Then
                                            pFeat.Value(iwIndex) = pCedge.Value(iwIndex)
                                            iwNode = pCedge.Value(iwIndex)
                                        ElseIf j = 0 And dir = "JI" Then
                                            pFeat.Value(jwIndex) = pCedge.Value(iwIndex)
                                            jwNode = pCedge.Value(iwIndex)
                                        ElseIf j = 1 And dir = "IJ" Then
                                            pFeat.Value(jwIndex) = pCedge.Value(iwIndex)
                                            jwNode = pCedge.Value(iwIndex)
                                        ElseIf j = 1 And dir = "JI" Then
                                            pFeat.Value(iwIndex) = pCedge.Value(iwIndex)
                                            iwNode = pCedge.Value(iwIndex)
                                        End If
                                        pFeat.Store()
                                        found = True
                                    End If
                                ElseIf otherj = tempnode Then
                                    If fVerboseLog Then WriteLogLine("fce other j = tempnode")
                                    If (pCedge.Value(jwIndex) > 0) Then
                                        'have a weave link to reuse
                                        If j = 0 And dir = "IJ" Then
                                            pFeat.Value(iwIndex) = pCedge.Value(jwIndex)
                                            iwNode = pCedge.Value(jwIndex)
                                        ElseIf j = 0 And dir = "JI" Then
                                            pFeat.Value(jwIndex) = pCedge.Value(jwIndex)
                                            jwNode = pCedge.Value(jwIndex)
                                        ElseIf j = 1 And dir = "IJ" Then
                                            pFeat.Value(jwIndex) = pCedge.Value(jwIndex)
                                            jwNode = pCedge.Value(jwIndex)
                                        ElseIf j = 1 And dir = "JI" Then
                                            pFeat.Value(iwIndex) = pCedge.Value(jwIndex)
                                            iwNode = pCedge.Value(jwIndex)
                                        End If
                                        pFeat.Store()
                                        found = True
                                    End If
                                End If
                            End If  'fce=true

                            'now write this part out
                            If found = True Then
                                If fVerboseLog Then WriteLogLine("print found = true other")
                                If j = 0 Then
                                    If dir = "IJ" Then
                                        wNodes = CStr(iwNode)
                                    Else
                                        wNodes = CStr(jwNode)
                                    End If
                                Else
                                    If Not ptstring = "" Then ptstring = ptstring + vbCrLf
                                    If Not wString = "" Then wString = wString + vbCrLf
                                    If Not astring = "" Then astring = astring + vbCrLf
                                    If dir = "IJ" Then
                                        wNodes = wNodes + " " + CStr(jwNode)
                                    Else
                                        wNodes = wNodes + " " + CStr(iwNode)
                                    End If
                                End If
                                If fVerboseLog Then WriteLogLine(wNodes)
                            End If

                        End If

                        Dim wcase As Integer
                        If found = False Then
                            If fVerboseLog Then WriteLogLine("cedge null or found is false")
                            'then don't have a weave linke to re-use
                            'determine if case 2 or 3
                            If fie = True Then
                                If facType = 1 Then 'freeway
                                    If fVerboseLog Then WriteLogLine("case3")
                                    wcase = 3
                                ElseIf iLanes = 0 And pIedge.Value(pIedge.Fields.FindField("FacilityType")) <> 1 Then 'non-freeway
                                    If fVerboseLog Then WriteLogLine("case 2 split")
                                    wcase = 2 'split
                                Else
                                    If fVerboseLog Then WriteLogLine("fie and ilanes not 0")
                                    wcase = 1
                                End If
                            Else
                                If fVerboseLog Then WriteLogLine("not fie")
                                wcase = 1
                            End If

                            'case 3 or case 1-code is the same
                            If fVerboseLog Then WriteLogLine("wcase " + CStr(wcase))
                            If wcase = 1 Or wcase = 3 Then

                                pQF = New QueryFilter
                                If j = 0 Then
                                    If dir = "IJ" Then
                                        iwNode = LargestJunct 'LargestWJunct
                                        pFeat.Value(iwIndex) = iwNode
                                        '[050407] hyu
                                        'pQF.WhereClause = "Scen_Node = " + CStr(iNode)
                                        nodeID = iNode
                                    Else 'ji
                                        jwNode = LargestJunct 'LargestWJunct
                                        pFeat.Value(jwIndex) = jwNode
                                        '[050407] hyu
                                        'pQF.WhereClause = "Scen_Node = " + CStr(jNode)
                                        nodeID = JNode
                                    End If

                                Else ' j = 1
                                    If dir = "IJ" Then
                                        jwNode = LargestJunct 'LargestWJunct
                                        pFeat.Value(jwIndex) = jwNode
                                        '[050407] hyu
                                        'pQF.WhereClause = "Scen_Node = " + CStr(jNode)
                                        nodeID = JNode
                                    Else
                                        iwNode = LargestJunct 'LargestWJunct
                                        pFeat.Value(iwIndex) = iwNode
                                        '[050407] hyu
                                        'pQF.WhereClause = "Scen_Node = " + CStr(iNode)
                                        nodeID = iNode
                                    End If
                                End If
                                pFeat.Store()
                                'LargestWJunct = LargestWJunct + 1
                                LargestJunct = LargestJunct + 1

                                '[050407] hyu
                                ' pFC = m_junctShp.Search(pQF, False)
                                'If (m_junctShp.featurecount(pQF)) Then
                                '     pJFeat = pFC.NextFeature
                                If dctJcts.ContainsKey(nodeID) Then
                                    pJFeat = dctJcts.Item(nodeID)

                                    If j = 0 Then
                                        pPoint = getWeavenode(pJFeat, lType)
                                        If dir = "IJ" Then
                                            wNodes = CStr(iwNode)
                                            ptstring = "a " + CStr(iwNode) + " " + CStr(pPoint.X) + " " + CStr(pPoint.Y) + " 0 0 0 0"
                                            wString = wString + "a " + CStr(iwNode) + " " + CStr(iNode) + " " + wDefault + " " + ijtemp

                                            '[041806] hyu: remove the space at the beginning of astring
                                            'astring = astring + " " + CStr(iwNode) + " " + CStr(iNode) + " 0 99"
                                            astring = astring + CStr(iwNode) + " " + CStr(iNode) + wID_Type
                                            'now write opposite
                                            wString = wString + vbCrLf + "a " + CStr(iNode) + " " + CStr(iwNode) + " " + wDefault + " " + ijtemp
                                            '[041806] hyu: remove the space at the beginning of astring
                                            'astring = astring + vbCrLf + " " + CStr(iNode) + " " + CStr(iwNode) + " 0 99"
                                            astring = astring + vbCrLf + CStr(iNode) + " " + CStr(iwNode) + wID_Type
                                        Else    'dir<>"IJ"
                                            wNodes = CStr(jwNode)
                                            ptstring = "a " + CStr(jwNode) + " " + CStr(pPoint.X) + " " + CStr(pPoint.Y) + " 0 0 0 0"
                                            wString = wString + "a " + CStr(jwNode) + " " + CStr(JNode) + " " + wDefault + " " + ijtemp
                                            '[041806] hyu: remove the space at the beginning of astring
                                            'astring = astring + " " + CStr(jwNode) + " " + CStr(jNode) + " 0 99"
                                            astring = astring + CStr(jwNode) + " " + CStr(JNode) + wID_Type
                                            'now write opposite
                                            wString = wString + vbCrLf + "a " + CStr(JNode) + " " + CStr(jwNode) + " " + wDefault + " " + jitemp
                                            '[041806] hyu: remove the space at the beginning of astring
                                            'astring = astring + vbCrLf + " " + CStr(jNode) + " " + CStr(jwNode) + " 0 99"
                                            astring = astring + vbCrLf + CStr(JNode) + " " + CStr(jwNode) + wID_Type
                                        End If  'dir
                                        If fVerboseLog Then WriteLogLine(wNodes)
                                        If fVerboseLog Then WriteLogLine(ptstring)
                                        If fVerboseLog Then WriteLogLine(wString)
                                    Else    'j<>0
                                        pPoint = getWeavenode(pJFeat, lType)
                                        If Not ptstring = "" Then ptstring = ptstring + vbCrLf
                                        If Not wString = "" Then wString = wString + vbCrLf
                                        If Not astring = "" Then astring = astring + vbCrLf
                                        If dir = "IJ" Then
                                            wNodes = wNodes + " " + CStr(jwNode)
                                            ptstring = ptstring + "a " + CStr(jwNode) + " " + CStr(pPoint.X) + " " + CStr(pPoint.Y) + " 0 0 0 0"
                                            wString = wString + "a " + CStr(jwNode) + " " + CStr(JNode) + " " + wDefault + " " + jitemp
                                            '[041806] hyu: remove the space at the beginning of astring
                                            'astring = astring + " " + CStr(jwNode) + " " + CStr(jNode) + " 0 99"
                                            astring = astring + CStr(jwNode) + " " + CStr(JNode) + wID_Type
                                            'now write opposite
                                            wString = wString + vbCrLf + "a " + CStr(JNode) + " " + CStr(jwNode) + " " + wDefault + " " + jitemp
                                            '[041806] hyu: remove the space at the beginning of astring
                                            'astring = astring + vbCrLf + " " + CStr(jNode) + " " + CStr(jwNode) + " 0 99"
                                            astring = astring + vbCrLf + CStr(JNode) + " " + CStr(jwNode) + wID_Type
                                        Else    'dir<>"IJ"
                                            wNodes = wNodes + " " + CStr(iwNode)
                                            ptstring = ptstring + "a " + CStr(iwNode) + " " + CStr(pPoint.X) + " " + CStr(pPoint.Y) + " 0 0 0 0"
                                            wString = wString + "a " + CStr(iwNode) + " " + CStr(iNode) + " " + wDefault + " " + jitemp
                                            '[041806] hyu: remove the space at the beginning of astring
                                            'astring = astring + " " + CStr(iwNode) + " " + CStr(iNode) + " 0 99"
                                            astring = astring + CStr(iwNode) + " " + CStr(iNode) + wID_Type
                                            'now write opposite
                                            wString = wString + vbCrLf + "a " + CStr(iNode) + " " + CStr(iwNode) + " " + wDefault + " " + ijtemp
                                            '[041806] hyu: remove the space at the beginning of astring
                                            'astring = astring + vbCrLf + " " + CStr(iNode) + " " + CStr(iwNode) + " 0 99"
                                            astring = astring + vbCrLf + CStr(iNode) + " " + CStr(iwNode) + wID_Type
                                        End If  'dir
                                        If fVerboseLog Then WriteLogLine(wNodes)
                                        If fVerboseLog Then WriteLogLine(ptstring)
                                        If fVerboseLog Then WriteLogLine(wString)
                                    End If  'j
                                End If  'dctjcts.ContainsKey

                            Else    'wcase=2

                                pWSedit = pWorkspaceI

                                If fVerboseLog Then WriteLogLine("case 2")
                                'case 2 -require a split on pIedge and new node for built file
                                'new weave node needs to leng up the pIedge
                                'need to see if fromnode same as tempnode else need to re-calculate distace
                                'need to put all weave and shadow in different text file
                                'then if splits- need to redo the edge built files
                                Dim newPt As IPoint
                                Dim splitFt As IFeature
                                newPt = New ESRI.ArcGIS.Geometry.Point

                                '[033006] hyu: Editing already started Create_NetFile2
                                '                        pWSEdit.StartEditing False
                                '                        pWSEdit.StartEditOperation

                                '[051507]hyu: check the Iedge to make sure it isn't split repeatedly.
                                If pIedge.Value(pIedge.Fields.FindField("Split" & lType)) = lType Then
                                    'if the Iedge has already been split, then no more split
                                    If fVerboseLog Then WriteLogLine("use existing splitM")
                                    If j = 0 Then
                                        If dir = "IJ" Then
                                            If pIedge.Value(pIedge.Fields.FindField("INODE")) = tempnode Then
                                                iwNode = pIedge.Value(pIedge.Fields.FindField("JNODE")) + m_Offset
                                            Else
                                                iwNode = pIedge.Value(pIedge.Fields.FindField("INODE")) + m_Offset
                                            End If

                                            pFeat.Value(iwIndex) = iwNode

                                            splitFt = dctJcts.Item(iwNode)
                                            newPt = splitFt.Shape
                                        Else 'ji
                                            If pIedge.Value(pIedge.Fields.FindField("INODE")) = tempnode Then
                                                jwNode = pIedge.Value(pIedge.Fields.FindField("JNODE")) + m_Offset
                                            Else
                                                jwNode = pIedge.Value(pIedge.Fields.FindField("INODE")) + m_Offset
                                            End If
                                            pFeat.Value(jwIndex) = jwNode
                                            pFeat.Store()
                                            splitFt = dctJcts.Item(jwNode)
                                            newPt = splitFt.Shape
                                        End If

                                    Else ' j = 1
                                        If dir = "IJ" Then
                                            If pIedge.Value(pIedge.Fields.FindField("INODE")) = tempnode Then
                                                jwNode = pIedge.Value(pIedge.Fields.FindField("JNODE")) + m_Offset
                                            Else
                                                jwNode = pIedge.Value(pIedge.Fields.FindField("INODE")) + m_Offset
                                            End If
                                            pFeat.Value(jwIndex) = jwNode

                                            splitFt = dctJcts.Item(jwNode)
                                            newPt = splitFt.Shape
                                        Else
                                            If pIedge.Value(pIedge.Fields.FindField("INODE")) = tempnode Then
                                                iwNode = pIedge.Value(pIedge.Fields.FindField("JNODE")) + m_Offset
                                            Else
                                                iwNode = pIedge.Value(pIedge.Fields.FindField("INODE")) + m_Offset
                                            End If
                                            pFeat.Value(iwIndex) = iwNode
                                            pFeat.Store()
                                            splitFt = dctJcts.Item(iwNode)
                                            newPt = splitFt.Shape
                                        End If
                                    End If

                                Else
                                    pNString = " "
                                    'SplitFeatureByM(pIedge, Leng, tempnode, newPt, lType, dctEdges, dctJcts, dctWNs, dctWNs, pNString)
                                    updateTransitLines(newPt)
                                    Debug.Print("split " & pIedge.OID)
                                    'Debug.Print("create " & ctype((dctEdges.Item(dctEdges.Count - 1), string)
                                End If


                                '[051407] hyu: no need to incrememt the largest junct and largestwjunct
                                '' [071206] pan:   ID base up to include shadow link nodes created
                                '                        largestjunct = largestjunct + 1
                                '' [071406] pan:  this didn't work, re-added largestjunct increment above
                                '                        largestwjunct = largestwjunct + 1

                                '                        pWSEdit.StopEditOperation
                                '                        pWSEdit.StopEditing True

                                If fVerboseLog Then WriteLogLine("after splitM")
                                If j = 0 Then
                                    If dir = "IJ" Then
                                        '[051407] hyu: iwnode should be the junction id of the new junction at the splitting point
                                        'iwNode = largestwjunct
                                        iwNode = LargestJunct + m_Offset
                                        pFeat.Value(iwIndex) = iwNode

                                    Else 'ji
                                        '[051407] hyu: jwnode should be the junction id of the new junction at the splitting point
                                        jwNode = LargestJunct + m_Offset
                                        'jwNode = largestwjunct
                                        pFeat.Value(jwIndex) = jwNode
                                    End If
                                Else 'j = 1
                                    If dir = "IJ" Then
                                        '[051407] hyu: iwnode should be the junction id of the new junction at the splitting point
                                        'jwNode = largestwjunct
                                        jwNode = LargestJunct + m_Offset
                                        pFeat.Value(jwIndex) = jwNode
                                    Else
                                        '[051407] hyu: iwnode should be the junction id of the new junction at the splitting point
                                        'iwNode = largestwjunct
                                        iwNode = LargestJunct + m_Offset
                                        pFeat.Value(iwIndex) = iwNode
                                    End If
                                End If
                                pFeat.Store()
                                'largestwjunct = largestwjunct + 1

                                If j = 0 Then
                                    If fVerboseLog Then WriteLogLine("using new point")
                                    pPoint = newPt
                                    If dir = "IJ" Then
                                        wNodes = CStr(iwNode)
                                        ptstring = "a " + CStr(iwNode) + " " + CStr(pPoint.X) + " " + CStr(pPoint.Y) + " 0 0 0 0"
                                        wString = wString + "a " + CStr(iwNode) + " " + CStr(iNode) + " " + wDefault + " " + ijtemp
                                        '[041806] hyu: remove the space at the beginning of astring
                                        'astring = astring + " " + CStr(iwNode) + " " + CStr(iNode) + " 0 99"
                                        astring = astring + CStr(iwNode) + " " + CStr(iNode) + wID_Type
                                        'now write opposite
                                        wString = wString + vbCrLf + "a " + CStr(iNode) + " " + CStr(iwNode) + " " + wDefault + " " + ijtemp
                                        '[041806] hyu: remove the space at the beginning of astring
                                        'astring = astring + vbCrLf + " " + CStr(iNode) + " " + CStr(iwNode) + " 0 99"
                                        astring = astring + vbCrLf + CStr(iNode) + " " + CStr(iwNode) + wID_Type
                                    Else
                                        wNodes = CStr(jwNode)
                                        ptstring = "a " + CStr(jwNode) + " " + CStr(pPoint.X) + " " + CStr(pPoint.Y) + " 0 0 0 0"
                                        wString = wString + "a " + CStr(jwNode) + " " + CStr(JNode) + " " + wDefault + " " + ijtemp
                                        '[041806] hyu: remove the space at the beginning of astring
                                        'astring = astring + " " + CStr(jwNode) + " " + CStr(jNode) + " 0 99"
                                        astring = astring + CStr(jwNode) + " " + CStr(JNode) + wID_Type
                                        'now write opposite
                                        wString = wString + vbCrLf + "a " + CStr(JNode) + " " + CStr(jwNode) + " " + wDefault + " " + jitemp
                                        '[041806] hyu: remove the space at the beginning of astring
                                        'astring = astring + vbCrLf + " " + CStr(jNode) + " " + CStr(jwNode) + " 0 99"
                                        astring = astring + vbCrLf + CStr(JNode) + " " + CStr(jwNode) + wID_Type
                                    End If
                                    If fVerboseLog Then WriteLogLine(wNodes)
                                    If fVerboseLog Then WriteLogLine(ptstring)
                                    If fVerboseLog Then WriteLogLine(wString)
                                Else
                                    ' pPoint = getWeavenode(pJFeat, ltype)
                                    pPoint = newPt
                                    If Not ptstring = "" Then ptstring = ptstring + vbCrLf
                                    If Not wString = "" Then wString = wString + vbCrLf
                                    If Not astring = "" Then astring = astring + vbCrLf
                                    If dir = "IJ" Then
                                        wNodes = wNodes + " " + CStr(jwNode)
                                        ptstring = ptstring + "a " + CStr(jwNode) + " " + CStr(pPoint.X) + " " + CStr(pPoint.Y) + " 0 0 0 0"
                                        wString = wString + "a " + CStr(jwNode) + " " + CStr(JNode) + " " + wDefault + " " + jitemp
                                        astring = astring + CStr(jwNode) + " " + CStr(JNode) + wID_Type
                                        'now write opposite
                                        wString = wString + vbCrLf + "a " + CStr(JNode) + " " + CStr(jwNode) + " " + wDefault + " " + jitemp
                                        astring = astring + vbCrLf + CStr(JNode) + " " + CStr(jwNode) + wID_Type
                                    Else
                                        wNodes = wNodes + " " + CStr(iwNode)
                                        ptstring = ptstring + "a " + CStr(iwNode) + " " + CStr(pPoint.X) + " " + CStr(pPoint.Y) + " 0 0 0 0"
                                        wString = wString + "a " + CStr(iwNode) + " " + CStr(iNode) + " " + wDefault + " " + jitemp
                                        astring = astring + CStr(iwNode) + " " + CStr(iNode) + wID_Type
                                        'now write opposite
                                        wString = wString + vbCrLf + "a " + CStr(iNode) + " " + CStr(iwNode) + " " + wDefault + " " + ijtemp
                                        astring = astring + vbCrLf + CStr(iNode) + " " + CStr(iwNode) + wID_Type
                                    End If
                                    If fVerboseLog Then WriteLogLine(wNodes)
                                    If fVerboseLog Then WriteLogLine(ptstring)
                                    If fVerboseLog Then WriteLogLine(wString)
                                End If  'j
                            End If  'wcase
                        End If  'found

                    Else    '(m_edgeShp.featurecount(pFilter) <= 2
                        'need to create the weave links, and emme2 nodes

                        If fVerboseLog Then WriteLogLine("here1")
                        pQF = New QueryFilter
                        If j = 0 Then
                            If dir = "IJ" Then
                                iwNode = LargestJunct 'LargestWJunct
                                pFeat.Value(iwIndex) = iwNode
                                '[050407] hyu
                                'pQF.WhereClause = "Scen_Node = " + CStr(iNode)
                                nodeID = iNode
                            Else 'ji
                                jwNode = LargestJunct 'LargestWJunct
                                pFeat.Value(jwIndex) = jwNode
                                '[050407] hyu
                                'pQF.WhereClause = "Scen_Node = " + CStr(jNode)
                                nodeID = JNode
                            End If
                            pFeat.Store()
                        Else ' j = 1
                            If dir = "IJ" Then
                                jwNode = LargestJunct 'LargestWJunct
                                pFeat.Value(jwIndex) = jwNode
                                '[050407] hyu
                                'pQF.WhereClause = "Scen_Node = " + CStr(jNode)
                                nodeID = JNode
                            Else
                                iwNode = LargestJunct 'LargestWJunct
                                pFeat.Value(iwIndex) = iwNode
                                '[050407] hyu
                                'pQF.WhereClause = "Scen_Node = " + CStr(iNode)
                                nodeID = iNode
                            End If
                            pFeat.Store()
                        End If
                        'LargestWJunct = LargestWJunct + 1
                        LargestJunct = LargestJunct + 1

                        '[050407] hyu
                        ' pFC = m_junctShp.Search(pQF, False)
                        'If (m_junctShp.featurecount(pQF)) Then
                        '     pJFeat = pFC.NextFeature
                        If dctJcts.ContainsKey(nodeID) Then
                            pJFeat = dctJcts.Item(nodeID)
                            If j = 0 Then
                                pPoint = getWeavenode(pJFeat, lType)
                                If dir = "IJ" Then
                                    wNodes = CStr(iwNode)
                                    ptstring = "a " + CStr(iwNode) + " " + CStr(pPoint.X) + " " + CStr(pPoint.Y) + " 0 0 0 0"
                                    wString = wString + "a " + CStr(iwNode) + " " + CStr(iNode) + " " + wDefault + " " + ijtemp
                                    '[041806] hyu: remove the space at the beginning of astring
                                    'astring = astring + " " + CStr(iwNode) + " " + CStr(iNode) + " 0 99"
                                    astring = astring + CStr(iwNode) + " " + CStr(iNode) + wID_Type
                                    'now write opposite
                                    wString = wString + vbCrLf + "a " + CStr(iNode) + " " + CStr(iwNode) + " " + wDefault + " " + ijtemp
                                    '[041806] hyu: remove the space at the beginning of astring
                                    'astring = astring + vbCrLf + " " + CStr(iNode) + " " + CStr(iwNode) + " 0 99"
                                    astring = astring + vbCrLf + CStr(iNode) + " " + CStr(iwNode) + wID_Type
                                Else
                                    wNodes = CStr(jwNode)
                                    ptstring = "a " + CStr(jwNode) + " " + CStr(pPoint.X) + " " + CStr(pPoint.Y) + " 0 0 0 0"
                                    wString = wString + "a " + CStr(jwNode) + " " + CStr(JNode) + " " + wDefault + " " + ijtemp
                                    '[041806] hyu: remove the space at the beginning of astring
                                    'astring = astring + " " + CStr(jwNode) + " " + CStr(jNode) + " 0 99"
                                    astring = astring + CStr(jwNode) + " " + CStr(JNode) + wID_Type
                                    'now write opposite
                                    wString = wString + vbCrLf + "a " + CStr(JNode) + " " + CStr(jwNode) + " " + wDefault + " " + jitemp
                                    '[041806] hyu: remove the space at the beginning of astring
                                    'astring = astring + vbCrLf + " " + CStr(jNode) + " " + CStr(jwNode) + " 0 99"
                                    astring = astring + vbCrLf + CStr(JNode) + " " + CStr(jwNode) + wID_Type
                                End If
                                If fVerboseLog Then WriteLogLine(wNodes)
                                If fVerboseLog Then WriteLogLine(ptstring)
                                If fVerboseLog Then WriteLogLine(wString)
                            Else    'j=1
                                pPoint = getWeavenode(pJFeat, lType)
                                If Not ptstring = "" Then ptstring = ptstring + vbCrLf
                                If Not wString = "" Then wString = wString + vbCrLf
                                If Not astring = "" Then astring = astring + vbCrLf
                                If dir = "IJ" Then
                                    wNodes = wNodes + " " + CStr(jwNode)
                                    ptstring = ptstring + "a " + CStr(jwNode) + " " + CStr(pPoint.X) + " " + CStr(pPoint.Y) + " 0 0 0 0"
                                    wString = wString + "a " + CStr(jwNode) + " " + CStr(JNode) + " " + wDefault + " " + jitemp
                                    astring = astring + CStr(jwNode) + " " + CStr(JNode) + wID_Type
                                    'now write opposite
                                    wString = wString + vbCrLf + "a " + CStr(JNode) + " " + CStr(jwNode) + " " + wDefault + " " + jitemp
                                    astring = astring + vbCrLf + CStr(JNode) + " " + CStr(jwNode) + wID_Type
                                Else
                                    wNodes = wNodes + " " + CStr(iwNode)
                                    ptstring = ptstring + "a " + CStr(iwNode) + " " + CStr(pPoint.X) + " " + CStr(pPoint.Y) + " 0 0 0 0"
                                    wString = wString + "a " + CStr(iwNode) + " " + CStr(iNode) + " " + wDefault + " " + jitemp
                                    astring = astring + CStr(iwNode) + " " + CStr(iNode) + wID_Type
                                    'now write opposite
                                    wString = wString + vbCrLf + "a " + CStr(iNode) + " " + CStr(iwNode) + " " + wDefault + " " + ijtemp
                                    astring = astring + vbCrLf + CStr(iNode) + " " + CStr(iwNode) + wID_Type
                                End If
                                If fVerboseLog Then WriteLogLine(wNodes)
                                If fVerboseLog Then WriteLogLine(ptstring)
                                If fVerboseLog Then WriteLogLine(wString)
                            End If  'j
                        End If  'dctJcts.ContainsKey(nodeID)
                    End If  'm_edgeShp.featurecount(pFilter) <= 2
                    If j = 0 Then
                        If Not dctWeaveNodes.ContainsKey(CStr(iwNode)) Then
                            dctWeaveNodes.Add(CStr(iwNode), iwNode)
                            dctWeaveNodes2.Add(iNode, iwNode)
                        End If
                    Else
                        If Not dctWeaveNodes.ContainsKey(CStr(jwNode)) Then
                            dctWeaveNodes.Add(CStr(jwNode), jwNode)
                            dctWeaveNodes2.Add(JNode, jwNode)
                        End If
                    End If
                End If  'bWeavenodeExists=true
            Next j
        End If  '(Not IsDBNull(pFeat.value(iwIndex)) And pFeat.value(iwIndex) > 0)
        If fVerboseLog Then WriteLogLine("finished createWeaveLink call with ptstring=" & ptstring)
        Exit Sub

eh:
        CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.createWeaveLink")
        MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.createWeaveLink")
        'Print #9, Err.Description, vbInformation, "createWeaveLink"
        If pWSedit.IsBeingEdited Then
            pWSedit.AbortEditOperation()
            pWSedit.StopEditing(False)
        End If

    End Sub


Public Function getWeaveLen(lType As String) As Double
  'on error GoTo eh

 Dim pPt As IPoint
 Dim lenHOV As Double
 Dim lenTR As Double
 Dim lenTK As Double
 '[072006] pan:  these are changes to match specs
 '[051507] weave link length is in unit of mile, should be converted to feet
 'lenHOV = 0.015
 lenHOV = 0.01 * 5280
 'lenTR = 0.01
 lenTR = 0.015 * 5280
 'lenTK = 0.005
        lenTK = 0.005 * 5280
        If fVerboseLog Then WriteLogLine("weave link type  " + CStr(lType))
 Select Case lType
    Case "HOV"
      getWeaveLen = lenHOV
    Case "TR"
      getWeaveLen = lenTR
    Case "TK"
      getWeaveLen = lenTK
 End Select

 Exit Function

eh:
        CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.getWeaveLen")
        MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.getWeaveLen")

End Function

    Public Function getWeavenode(ByVal pJFeat As IFeature, ByVal lType As String) As IPoint
        'the amount to +/- to point coordinates is calculated through a2 + b2 = len**2
        'on error GoTo eh

        Dim moveX As Double
        Dim moveY As Double
        Dim Leng As Double
        Leng = getWeaveLen(lType)
        moveX = (Math.Sqrt(Leng)) / 2
        moveX = Math.Sqrt(moveX)
        moveY = moveX 'move the same amount as x

        g_xMove = moveX
        g_yMove = moveY

        Dim pGeom As IGeometry
        Dim pPt As IPoint
        pGeom = pJFeat.ShapeCopy
        pPt = pGeom
        pPt.X = pPt.X + moveX
        pPt.Y = pPt.Y + moveY
        getWeavenode = pPt

        Exit Function

eh:

        CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.getWeavenode")
        MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.getWeavenode")

    End Function



Public Function calAngle(pC As IPoint, pB As IPoint, pA As IPoint) As Double

'on error GoTo eh

'have to use the law of cosines for blique triangles
'a^2 = c^2 + b^2 - 2cb cosA
'derives to cosA = -a^2 + c^2 + b^2 / 2cb
'finally to A = arccos (-a^2 + c^2 + b^2) / 2cb)
   Dim alen As Double, blen As Double, clen As Double
   Dim angleA As Double
   'MsgBox " c " + CStr(pC.x) + " ," + CStr(pC.Y) + " b " + CStr(pB.x) + " ," + CStr(pB.Y) + " a " + CStr(pA.x) + " ," + CStr(pA.Y)
   alen = ((pB.x - pC.x) ^ 2) + ((pB.y - pC.y) ^ 2) 'the result of is alen squared
   blen = ((pA.x - pC.x) ^ 2) + ((pA.y - pC.y) ^ 2) 'the result of is blen squared
   clen = ((pA.x - pB.x) ^ 2) + ((pA.y - pB.y) ^ 2) 'the result of is clen squared
   ' MsgBox "a " + CStr(alen) + " b " + CStr(blen) + " c " + CStr(clen)
   'now calculate the angle using the tangent rule tan Y = opposite side/ adjacent side
 
        angleA = (blen + clen - alen) / (2 * Math.Sqrt(blen) * Math.Sqrt(clen))
   'vb has no command to calculate arccos- so has to be derived
   'arccos(X) = Atn(-X/Sqr(-X*X+1)) +2 * Atn(1)
        angleA = Math.Atan(-angleA / Math.Sqrt(-angleA * angleA + 1)) + 2 * Math.Atan(1) 'result is in radians

   Const pi As Double = 3.14159265358979
      calAngle = angleA * (180 / pi) 'now in degrees

    Exit Function

eh:
        CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.calAngle")
        MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.calAngle")

End Function

Public Sub create_NetFile0(pathnameN As String, filenameN As String, pWS As IWorkspace)
    'on error GoTo eh
    'called by GlobalMod.create_ScenarioShapefiles
    'pathnameN and filenameN input by user in frmNetLayer
    'm_edgeShp and m_junctShp are the intermediate shapefiles used to create the Netfile
    Dim strFFT As String    'holds formatted FFT number
    strFFT = ""

'    Dim pathname As String
'    pathname = "c:\createNet_ERROR.txt"
'    Open pathname For Output As #9

        WriteLogLine("")
        WriteLogLine("========================================")
        WriteLogLine("create_NetFile2 started " & Now())
        WriteLogLine("Model year " & GlobalMod.inserviceyear)
        WriteLogLine("========================================")
        WriteLogLine("")

    Dim pStatusBar As IStatusBar
    Dim pProgbar As IStepProgressor
     pStatusBar = m_App.StatusBar
     pProgbar = pStatusBar.ProgressBar

    Dim pathnameAm As String, pathnameM As String, pathnamePm As String, pathnameE As String, pathnameNi As String
    Dim attribname As String
    pathnameAm = pathnameN + "\" + filenameN + "AM.txt"
    pathnameM = pathnameN + "\" + filenameN + "M.txt"
    pathnamePm = pathnameN + "\" + filenameN + "PM.txt"
    pathnameE = pathnameN + "\" + filenameN + "E.txt"
    pathnameNi = pathnameN + "\" + filenameN + "N.txt"

    attribname = pathnameN + "\" + filenameN + "Attr.txt"
        'FileOpen(1, "c:\SumStats.txt", OpenMode.Append, OpenAccess.ReadWrite, OpenShare.Default)
        FileOpen(1, pathnameAm, OpenMode.Output, OpenAccess.ReadWrite, OpenShare.Default)
        FileOpen(7, pathnameM, OpenMode.Output, OpenAccess.ReadWrite, OpenShare.Default)
        FileOpen(4, pathnamePm, OpenMode.Output, OpenAccess.ReadWrite, OpenShare.Default)
        FileOpen(5, pathnameE, OpenMode.Output, OpenAccess.ReadWrite, OpenShare.Default)
        FileOpen(6, pathnameNi, OpenMode.Output, OpenAccess.ReadWrite, OpenShare.Default)
        FileOpen(3, attribname, OpenMode.Output, OpenAccess.ReadWrite, OpenShare.Default)

    

        'open temp files
        pathnameAm = pathnameN + "\" + filenameN + "AMtemp.txt"

        pathnameM = pathnameN + "\" + filenameN + "Mtemp.txt"
        pathnamePm = pathnameN + "\" + filenameN + "PMtemp.txt"
        pathnameE = pathnameN + "\" + filenameN + "Etemp.txt"
        pathnameNi = pathnameN + "\" + filenameN + "Ntemp.txt"

        FileOpen(11, pathnameAm, OpenMode.Output, OpenAccess.ReadWrite, OpenShare.Default)
        FileOpen(17, pathnameM, OpenMode.Output, OpenAccess.ReadWrite, OpenShare.Default)
        FileOpen(14, pathnamePm, OpenMode.Output, OpenAccess.ReadWrite, OpenShare.Default)
        FileOpen(15, pathnameE, OpenMode.Output, OpenAccess.ReadWrite, OpenShare.Default)
        FileOpen(16, pathnameNi, OpenMode.Output, OpenAccess.ReadWrite, OpenShare.Default)

        Dim mydate As Date
        mydate = Date.Now
        PrintLine(1, "c Exported AM Network File from ArcMap /Emme/2 Interface: " + CStr(mydate))
        PrintLine(7, "c Exported Midday Network File from ArcMap /Emme/2 Interface: " + CStr(mydate))
        PrintLine(4, "c Exported PM Network File from ArcMap /Emme/2 Interface: " + CStr(mydate))
        PrintLine(5, "c Exported Evening Network File from ArcMap /Emme/2 Interface: " + CStr(mydate))
        PrintLine(6, "c Exported Night Network File from ArcMap /Emme/2 Interface: " + CStr(mydate))
    Printline (3, "c Exported Link Attribute File from ArcMap /Emme/2 Interface: " + CStr(mydate))

        Dim pWorkspace As IWorkspace
        pWorkspace = g_FWS
        Dim pTblLine As ITable
        'Dim ptblPoint As ITable no longer using this table
        Dim pTblMode As ITable
        Dim pEvtLine As ITable
        Dim pevtPoint As ITable

        Dim wString As String
        Dim pstring As String
        wString = ""
        pstring = ""
        Dim nodes As String
        nodes = ""

        'get tables needed for the attribute data need to send to emme/2
        pTblLine = get_TableClass(m_layers(3)) 'use if future project
        ' ptblPoint = get_TableClass("tblPointProjects", pWorkspace) 'use if future project
        pTblMode = get_TableClass(m_layers(2))
        pEvtLine = get_TableClass(m_layers(4)) 'use if future event project
        pevtPoint = get_TableClass(m_layers(6)) 'use if future event project
        Dim pTC As ICursor, pTCt As ICursor
        Dim pRow As IRow
        Dim pQFt As IQueryFilter, pQF As IQueryFilter
        Dim tempAttrib As String
        Dim pFeatCursor As IFeatureCursor

        Dim lType(4) As String
        lType(0) = "GP"
        '    ltype(1) = "HOV"
        '    ltype(2) = "TR"
        lType(2) = "HOV"
        lType(1) = "TR"
        lType(3) = "TK"
        Dim timePd(5) As String
        timePd(0) = "AM"
        timePd(1) = "MD"
        timePd(2) = "PM"
        timePd(3) = "EV"
        timePd(4) = "NI"

        Dim pFeat As IFeature
        Dim pFlds As IFields
        Dim lSFld As Long, upIndex As Long, tempIndex As Long
        Dim Tindex As Long
        Dim i As Long
        Dim pWSedit As IWorkspaceEdit
        pWSedit = pWS

        Dim count As Long
        count = m_EdgeSSet.Count '+ m_JunctSSet.count
        pProgbar.Position = 0
        pStatusBar.ShowProgressBar("Creating NetFile...", 0, count, 1, True)
        'jaf--let's keep the user updated
        g_frmNetLayer.lblStatus.Text = ("GlobalMod.create_Netfile: Starting Netfile Build...")
        'pan frmNetLayer.Refresh

        GetLargestID(m_junctShp, "PSRCJunctID", LargestJunct)

        pFeatCursor = m_junctShp.Search(Nothing, False)
        PrintLine(1, "t nodes init")
        PrintLine(7, "t nodes init")
        PrintLine(4, "t nodes init")
        PrintLine(5, "t nodes init")
        PrintLine(6, "t nodes init")

        Dim test As Boolean
        Dim tempString As String, tstring As String
        Dim pPoint As IPoint
        Dim pMRow As IRow

        test = False
        Dim j As Long, shpindex As Long
        Dim tempprj As ClassPrjSelect
        pFeat = pFeatCursor.NextFeature
        Dim astring As String
        Dim HaveSplits As Boolean
        HaveSplits = False
        Dim splitquery As String

        lSFld = pFeat.Fields.FindField("Scen_Node") 'this id used by emme2

        upIndex = pFeat.Fields.FindField("Updated1")
        shpindex = pFeat.Fields.FindField("shptype")
        Dim Subtype As Long
        Subtype = pFeat.Fields.FindField("JunctionType")

        'jaf--let's keep the user updated
        g_frmNetLayer.lblStatus.Text = "GlobalMod.create_Netfile: Starting Junction portion of NetFile..."
        'pan frmNetLayer.Refresh

        Dim pcount As Long
        Dim dctReservedNodes As Dictionary(Of Object, Object)  '[042506] hyu: dictionary for reserved nodes: key=PSRCJunctId, item=Emme2Node
        Dim dctWeaveNodes As Dictionary(Of Object, Object)     '[051106] hyu: dictionary for weave nodes not physically in the ScenarioJunct layer
        Dim dctWeaveNodes2 As Dictionary(Of Object, Object)    '[052007] hyu: dictionary for junctions and corresponding weave nodes
        Dim lPSRCJctID As Long
        Dim dctJctsCor As Dictionary(Of Object, Object)

        dctJctsCor = New Dictionary(Of Object, Object)
        dctReservedNodes = New Dictionary(Of Object, Object)
        dctWeaveNodes = New Dictionary(Of Object, Object)
        dctWeaveNodes2 = New Dictionary(Of Object, Object)

        '[050407]hyu
        Dim dctJcts As Dictionary(Of Object, Object), dctJoints As Dictionary(Of Object, Object), dctEdgeID As Dictionary(Of Object, Object), lMaxEdgeOID As Long
        Dim pNewFeat As IFeature
        dctJcts = New Dictionary(Of Object, Object)
        '     dctJoints = New Dictionary
        '     dctEdgeID = New Dictionary

        '    calcEdgesAtJoint m_edgeShp, m_junctShp, g_INode, g_JNode, dctEdges, dctJcts, dctEdgeID, lMaxEdgeOID
        '
        WriteLogLine("start getting all nodes " & Now())
        pcount = 1
        Do Until pFeat Is Nothing 'loop through features in intermediate junction shapefile
            pPoint = pFeat.Shape
            '[042706] pan:  added parkandride node subtype 7 to test-they need a * also
            If (pFeat.Value(Subtype) = 6) Or (pFeat.Value(Subtype) = 7) Then
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

            If lPSRCJctID <> pFeat.Value(lSFld) Then
                If Not dctReservedNodes.ContainsKey(CStr(lPSRCJctID)) Then dctReservedNodes.Add(CStr(lPSRCJctID), CStr(pFeat.Value(lSFld)))
            End If

            '************************************
            'jaf--removed block 1 for readability.  see removed_block_1.txt
            '************************************

            PrintLine(7, tempString)
            PrintLine(1, tempString)
            PrintLine(4, tempString)
            PrintLine(5, tempString)
            PrintLine(6, tempString)

            'pStatusBar.StepProgressBar
            pFeat = pFeatCursor.NextFeature
        Loop 'end of loop through every feature in intermediate junction shapefile



        'jaf--let's keep the user updated
        g_frmNetLayer.lblStatus.Text = "GlobalMod.create_Netfile: Finished creating Junction portion of Netfile"
        'pan frmNetLayer.Refresh

        Dim NextLine As String
        Dim lanecap As Double, length As Double
        Dim oneIndex As Long, cnt As Long
        Dim fromI As Boolean
        oneIndex = m_edgeShp.FindField("Oneway")

        'jaf--this appears to commence the links portion of the buildfiles
        PrintLine(17, "t links init")
        PrintLine(11, "t links init")
        PrintLine(14, "t links init")
        PrintLine(15, "t links init")
        PrintLine(16, "t links init")

        Dim tempstringAm As String, tempstringM As String, tempstringPm As String, tempstringE As String, tempstringN As String


        '       pathname = "c:\MODULEcreateweave.txt"
        '       Open pathname For Append As #2
        WriteLogLine("Model Input create weave links ")

        'jaf--let's keep the user updated
        g_frmNetLayer.lblStatus.Text = "GlobalMod.create_Netfile: Starting Edge portion of NetFile..."
        'pan frmNetLayer.Refresh

        '[051507]hyu: no need of cursor
        upIndex = m_edgeShp.FindField("Updated1")
        lSFld = m_edgeShp.FindField("PSRCEdgeID") 'this used to get out of tblmode
        shpindex = m_edgeShp.FindField("shptype")

        '     pFeatCursor = m_edgeShp.Search(Nothing, False)
        '     pFeat = pFeatCursor.NextFeature
        '    upIndex = pFeat.Fields.FindField("Updated1")
        '    lSFld = pFeat.Fields.FindField("PSRCEdgeID") 'this used to get out of tblmode
        '    shpindex = pFeat.Fields.FindField("shptype")

        Dim Wprint As Boolean
        Dim intWays As Integer
        Dim uindex As Integer
        Dim lanes As Long
        Dim direction As String
        Dim t As Integer, l As Integer
        Dim lflag As String
        Dim mode As String
        Dim linkType As String
        Dim addweave As Integer
        Dim laneFld As String
        Dim pos As Long
        Dim Leng
        Dim newstr As String
        Dim rString As String
        Dim rLocation As Long
        Dim vdf As Long
        Dim lancapfld As String
        Dim intFacType As Integer

        '[033106] hyu: declare a lookup dictionary for edges in the ModeAttributes table.  This dictionary replaces the
        'query for row count in the ModeAttributes table.
        Dim dctMAtt As Dictionary(Of Object, Object), dctEdges As Dictionary(Of Long, Feature)
        Dim bEOF As Boolean, lEdgeCt As Long, sEdgeOID As String, sEdgeID As String

        WriteLogLine("start getting all edges and mode attributes " & Now())
        getMAtts(m_edgeShp, pTblMode, g_PSRCEdgeID, dctEdges, dctMAtt)

        WriteLogLine("start writing links " & Now())
        pWSedit.StartEditing(False)
        pWSedit.StartEditOperation()

        bEOF = False
        lEdgeCt = 0

        If dctEdges.Count = 0 Then Exit Sub

        Do Until bEOF
            '[050407] hyu

            pFeat = dctEdges.Item(lEdgeCt)
            lEdgeCt = lEdgeCt + 1
            If lEdgeCt = dctEdges.Count Then bEOF = True
            WriteLogLine("edge count=" & lEdgeCt & " " & Now())
            '        sEdgeOID = dctEdges.Item(lEdgeCt)
            '        If Trim(sEdgeOID) <> "" Then
            '             pFeat = m_edgeShp.GetFeature(val(sEdgeOID))
            '
            '        Else
            '            Do Until Trim(sEdgeOID) <> ""
            '                lEdgeCt = lEdgeCt + 1
            '                If lEdgeCt = dctEdges.count Then
            '                    bEOF = True
            '                    Exit Do
            '                End If
            '                sEdgeOID = dctEdges.Item(lEdgeCt)
            '            Loop
            '
            '            If bEOF Then Exit Do
            '        End If

            'Do Until pFeat Is Nothing 'loop throught all the features in the intermediate edge shapefile
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
                    cnt = 1
                    fromI = True
                Case 1
                    'one way from J
                    cnt = 1
                    fromI = False
                Case 2
                    'two way
                    cnt = 2
                    fromI = True
            End Select

            uindex = m_edgeShp.FindField("UseEmmeN")
            If fVerboseLog Then WriteLogLine("useemmen is index " & CStr(uindex))
            For i = 0 To cnt - 1
                tempString = ""
                nodes = ""
                'jaf--this block writes the inode and jnode values in appropriate direction
                'jaf--(columns inode and jnode)
                '[033006] hyu: problem - IJ for two-way is not considered
                If i = 0 Then
                    If fromI = True Then
                        '                    Tindex = pFeat.Fields.FindField("INode")
                        '                    If IsDBNull(pFeat.value(uindex)) Or pFeat.value(uindex) = 0 Or pFeat.value(uindex) = 2 Then
                        '
                        '                        If IsDBNull(pFeat.value(Tindex)) Then
                        '
                        '                            largestjunct = largestjunct + 1
                        '                            nodes = nodes + CStr(largestjunct + m_Offset)
                        '                        Else
                        '                            nodes = nodes + CStr(pFeat.value(Tindex) + m_Offset)
                        '                        End If
                        '                    Else
                        '                        nodes = nodes + CStr(pFeat.value(Tindex))
                        '                    End If
                        '
                        '                    Tindex = pFeat.Fields.FindField("JNode")
                        '                    If IsDBNull(pFeat.value(uindex)) Or pFeat.value(uindex) = 0 Or pFeat.value(uindex) = 1 Then
                        '                        '[040406] hyu: add judgement for JNode is null
                        '                       If IsDBNull(pFeat.value(Tindex)) Then
                        '                            largestjunct = largestjunct + 1
                        '                            nodes = nodes + " " + CStr(largestjunct + m_Offset)
                        '                        Else
                        '                            nodes = nodes + " " + CStr(pFeat.value(Tindex) + m_Offset)
                        '                        End If
                        ''                        nodes = nodes + " " + CStr(pFeat.value(Tindex) + m_Offset)
                        '                    Else
                        '                        nodes = nodes + " " + CStr(pFeat.value(Tindex))
                        '                    End If
                        '[042506] hyu: change the way to get correct node id.
                        Tindex = pFeat.Fields.FindField("INode")
                        If dctReservedNodes.ContainsKey(CStr(pFeat.Value(Tindex))) Then
                            nodes = nodes + CStr(dctReservedNodes.Item(CStr(pFeat.Value(Tindex))))
                        Else
                            nodes = nodes + CStr(pFeat.Value(Tindex) + m_Offset)
                        End If

                        Tindex = pFeat.Fields.FindField("JNode")
                        If dctReservedNodes.ContainsKey(CStr(pFeat.Value(Tindex))) Then
                            nodes = nodes + " " + CStr(dctReservedNodes.Item(CStr(pFeat.Value(Tindex))))
                        Else
                            nodes = nodes + " " + CStr(pFeat.Value(Tindex) + m_Offset)
                        End If

                    Else
                        '                    Tindex = pFeat.Fields.FindField("JNode")
                        '                    If IsDBNull(pFeat.value(uindex)) Or pFeat.value(uindex) = 0 Or pFeat.value(uindex) = 1 Then
                        '                       nodes = nodes + CStr(pFeat.value(Tindex) + m_Offset)
                        '                    Else
                        '                        nodes = nodes + CStr(pFeat.value(Tindex))
                        '                    End If
                        '                    Tindex = pFeat.Fields.FindField("INode")
                        '                    If IsDBNull(pFeat.value(uindex)) Or pFeat.value(uindex) = 0 Or pFeat.value(uindex) = 2 Then
                        '                       nodes = nodes + " " + CStr(pFeat.value(Tindex) + m_Offset)
                        '                    Else
                        '                        nodes = nodes + " " + CStr(pFeat.value(Tindex))
                        '                    End If
                        '[042506] hyu: change the way to get correct node id.
                        Tindex = pFeat.Fields.FindField("JNode")
                        If dctReservedNodes.ContainsKey(CStr(pFeat.Value(Tindex))) Then
                            nodes = nodes + CStr(dctReservedNodes.Item(CStr(pFeat.Value(Tindex))))
                        Else
                            nodes = nodes + CStr(pFeat.Value(Tindex) + m_Offset)
                        End If

                        Tindex = pFeat.Fields.FindField("INode")
                        If dctReservedNodes.ContainsKey(CStr(pFeat.Value(Tindex))) Then
                            nodes = nodes + " " + CStr(dctReservedNodes.Item(CStr(pFeat.Value(Tindex))))
                        Else
                            nodes = nodes + " " + CStr(pFeat.Value(Tindex) + m_Offset)
                        End If

                    End If
                Else
                    '                Tindex = pFeat.Fields.FindField("JNode")
                    '                If IsDBNull(pFeat.value(uindex)) Or pFeat.value(uindex) = 0 Or pFeat.value(uindex) = 1 Then
                    '                       nodes = nodes + CStr(pFeat.value(Tindex) + m_Offset)
                    '                    Else
                    '                        nodes = nodes + CStr(pFeat.value(Tindex))
                    '                    End If
                    '                Tindex = pFeat.Fields.FindField("INode")
                    '                 If pFeat.value(uindex) = 0 Or pFeat.value(uindex) = 2 Then
                    '                       nodes = nodes + " " + CStr(pFeat.value(Tindex) + m_Offset)
                    '                    Else
                    '                        nodes = nodes + " " + CStr(pFeat.value(Tindex))
                    '                    End If
                    '[042506] hyu: change the way to get correct node id.
                    Tindex = pFeat.Fields.FindField("JNode")
                    If dctReservedNodes.ContainsKey(CStr(pFeat.Value(Tindex))) Then
                        nodes = nodes + CStr(dctReservedNodes.Item(CStr(pFeat.Value(Tindex))))
                    Else
                        nodes = nodes + CStr(pFeat.Value(Tindex) + m_Offset)
                    End If

                    Tindex = pFeat.Fields.FindField("INode")
                    If dctReservedNodes.ContainsKey(CStr(pFeat.Value(Tindex))) Then
                        nodes = nodes + " " + CStr(dctReservedNodes.Item(CStr(pFeat.Value(Tindex))))
                    Else
                        nodes = nodes + " " + CStr(pFeat.Value(Tindex) + m_Offset)
                    End If
                End If

                'jaf--this block writes the length in miles--6 ASCII CHARS MAX LEN!
                'jaf--(column len(miles))
                tempAttrib = nodes
                nodes = "a " + nodes
                If fVerboseLog Then WriteLogLine("nodes string=" & nodes)

                '[032706] hyu: add a judgement before getting Tindex
                If strLayerPrefix = "SDE" Then
                    Tindex = pFeat.Fields.FindField("SHAPE.len") 'SDE
                Else
                    Tindex = pFeat.Fields.FindField("Shape_Length") 'PGDB
                End If

                '[040406] hyu: Length info can be null
                '            length = pFeat.value(Tindex)
                '            length = length * 0.00018939393939
                If IsDBNull(pFeat.Value(Tindex)) Then
                    length = 0.015
                Else
                    length = pFeat.Value(Tindex)
                    length = length * 0.00018939393939
                End If

                tempString = tempString + " " + Left(CStr(length), 6)

                'jaf--lane type is derived from projects, if any
                'jaf--assume this all works for now
                '[033006] hyu: get correct link attributes as following
                'a real link from TransRefEdges: modeAttributes (pTC)
                'a project link: pTblLine
                'a event link: pEvtLine
                Tindex = pFeat.Fields.FindField("Scen_Link") 'this links the attribute file (which gets included in the punch) back tot he intermediate shp
                tempAttrib = tempAttrib + " " + CStr(pFeat.Value(Tindex))

                pQFt = New QueryFilter
                pQFt.WhereClause = "PSRCEdgeID = " + CStr(pFeat.Value(lSFld))

                '[033106] hyu: change to lookup a dictionary before doing the query
                If Not dctMAtt.ContainsKey(CStr(pFeat.Value(lSFld))) Then
                    '             pTC = pTblMode.Search(pQFt, False)
                    '            If (pTblMode.rowcount(pQFt) = 0) Then 'should always be a match in tblmode
                    'jaf--modified to give more info
                    WriteLogLine("DATA ERROR: Edge " & CStr(pFeat.Value(lSFld)) & " is missing in ModeAttributes--GlobalMod.create_NetFile")
                    'move to next one
                Else
                    '[040406] hyu: change to lookup a dictionary
                    pMRow = dctMAtt.Item(CStr(pFeat.Value(lSFld)))
                    '                 pTC = pTblMode.Search(pQFt, False)
                    '                 pMrow = pTC.NextRow

                    If (pFeat.Value(upIndex) = "Yes") Then
                        'jaf--not sure what the Update field signifies
                        If (pFeat.Value(shpindex) <> "Event") Then
                            'get future in service attributes from tblLine
                            tempIndex = pFeat.Fields.FindField("prjRte")
                            tstring = pFeat.Value(tempIndex)
                            pQFt = New QueryFilter
                            pQFt.WhereClause = "projRteID = " + tstring
                            pTC = pTblLine.Search(pQFt, False)
                            If (pTblLine.RowCount(pQFt) = 0) Then
                                'jaf--modified to give more info
                                WriteLogLine("DATA ERROR: project " & CStr(pFeat.Value(lSFld)) & " is missing in tblLine--GlobalMod.create_NetFile")
                            Else
                                pRow = pTC.NextRow
                            End If
                        Else
                            tempIndex = pFeat.Fields.FindField("prjRte")
                            tstring = pFeat.Value(tempIndex)
                            pQFt = New QueryFilter
                            pQFt.WhereClause = "projRteID = " + tstring
                            pTC = pEvtLine.Search(pQFt, False)
                            If (pEvtLine.RowCount(pQFt) = 0) Then
                                'jaf--modified to give more info
                                WriteLogLine("DATA ERROR: project " & CStr(pFeat.Value(lSFld)) & " is missing in evtLine--GlobalMod.create_NetFile")
                            Else
                                pRow = pTC.NextRow
                            End If
                        End If
                    Else
                        pRow = pMRow
                    End If 'now  which table to pull the attributes from

                    If fVerboseLog Then WriteLogLine("setting direction")

                    For l = 0 To 2 'for now just do gp and hov and tr
                        Wprint = False

                        If (i = 0 And fromI = True) Then
                            direction = "IJ"
                            If l = 0 Then lflag = " 1 " ' flag for IJ
                            If l = 1 Then lflag = " 3 " ' flag for IJHOV
                            If l = 2 Then lflag = " 5 " ' flag for IJTR
                            'If l = 3 Then lflag =  " 7 "  ' flag for IJHOV
                        ElseIf (i = 1 Or fromI = False) Then
                            direction = "JI"
                            If l = 0 Then lflag = " 2 " ' flag for JI
                            If l = 1 Then lflag = " 4 " 'flag for JIHOV
                            If l = 2 Then lflag = " 6 " ' flag for IJTR
                            'If l = 3 Then lflag = " 8 "  ' flag for IJHOV
                        End If

                        If l = 0 Then
                            tempString = nodes + tempString
                            nodes = ""
                        End If

                        'jaf--this block writes the modes string, which has defaults if not supplied
                        'jaf--(column modes)

                        Tindex = pFeat.Fields.FindField("Modes")
                        If (pFeat.Value(Tindex) = " " Or IsDBNull(pFeat.Value(Tindex)) = True Or pFeat.Value(Tindex) = "") Then
                            'jaf--default value of modes string
                            '[051507] hyu default modes
                            Select Case l
                                Case 0
                                    mode = "ashi"
                                Case 1  'tr
                                    mode = "ab"
                                Case 2  'hov

                                Case 3  'tk
                                    mode = "auv"
                            End Select
                            '                        Select Case l
                            '                          Case 0 'gp
                            '                             mode = "ashi"
                            '                          Case 2 'tk
                            '                            mode = "auv"
                            '                          Case 3 'tr
                            '                            mode = "ab"
                            '                        End Select
                        Else
                            'supplied value of modes string
                            mode = CStr(pFeat.Value(Tindex))
                        End If

                        'jaf--this block writes the linktype number (between 1-99, used to flag screen lines)
                        'jaf--(column linktype)
                        'jaf--default to 90 for now, we'll have to figure out how to capture this in GeoDB

                        Tindex = pFeat.Fields.FindField("LinkType")
                        If (IsDBNull(pFeat.Value(Tindex)) = False And pFeat.Value(Tindex) > 0) Then
                            linkType = pFeat.Value(Tindex)
                        Else
                            linkType = "90"
                        End If

                        'jaf--all times of day buildfiles get same info up to this point
                        'assign so far a iNode, jNode, length, modes, linkType if not hov
                        'If Not l = 1 Then tempString = tempString + " " + mode + " " + linktype
                        If l <> 2 Then tempString = tempString + " " + mode + " " + linkType
                        'else if hov then mode depends on value in lanes field
                        tempstringAm = tempString
                        tempstringM = tempString
                        tempstringPm = tempString
                        tempstringE = tempString
                        tempstringN = tempString

                        pstring = "" 'reset for new shadow link type
                        wString = ""
                        astring = ""

                        addweave = 0

                        For t = 0 To 4
                            'NOW assign: a iNode, jNode, length, modes, linkType, #lane
                            lanes = 0
                            If l = 0 Or l = 2 Then 'If l < 2 Then
                                laneFld = direction + "Lanes" + lType(l) + timePd(t)
                            Else
                                laneFld = direction + "Lanes" + lType(l)
                            End If

                            Tindex = pRow.Fields.FindField(laneFld)
                            If (IsDBNull(pRow.Value(Tindex)) = False And pRow.Value(Tindex) > 0) Then
                                lanes = pRow.Value(Tindex)
                            Else
                                Tindex = pMRow.Fields.FindField(laneFld)
                                If (IsDBNull(pMRow.Value(Tindex)) = False And pMRow.Value(Tindex) > 0) Then
                                    lanes = pMRow.Value(Tindex)

                                Else
                                    If l = 0 Then
                                        lanes = 1
                                        'If t = 0 Then MsgBox "no lane value " + tempstringAm
                                    End If
                                End If
                            End If

                            If fVerboseLog Then WriteLogLine("checking lanes")
                            If lanes > 0 Then
                                If l = 0 Then
                                    If t = 0 Then
                                        tempstringAm = tempstringAm + " " + CStr(lanes)
                                        tempAttrib = tempAttrib + " " + lflag
                                    End If
                                    If t = 1 Then tempstringM = tempstringM + " " + CStr(lanes)
                                    If t = 2 Then tempstringPm = tempstringPm + " " + CStr(lanes)
                                    If t = 3 Then tempstringE = tempstringE + " " + CStr(lanes)
                                    If t = 4 Then tempstringN = tempstringN + " " + CStr(lanes)
                                End If

                                If l > 0 Then
                                    Wprint = True
                                    If fVerboseLog Then WriteLogLine("go to weave")

                                    nodes = ""
                                    '

                                    'createWeaveLink(lType(l), pFeat, pRow, pstring, wString, nodes, astring, timePd(t), lanes, direction, dctReservedNodes, dctWeaveNodes, dctWeaveNodes2, dctJcts, dctEdges)

                                    '[033106] hyu: edit should start earlier
                                    '                                pWSEdit.StartEditing False
                                    '                                pWSEdit.StartEditOperation
                                    If SplitM = True Then
                                        '[050407] hyu: dictionary updated in the split function
                                        '[040406] hyu: update the edge dictionary
                                        'dctEdges.Add CStr(spID), CStr(spID)

                                        If HaveSplits = False Then
                                            HaveSplits = True
                                            splitquery = "FID =" + CStr(spID)
                                        Else
                                            splitquery = splitquery + "And FID =" + CStr(spID)
                                        End If

                                        FileClose(11)
                                        FileClose(17)
                                        FileClose(14)
                                        FileClose(15)
                                        FileClose(16)

                                        'now copy temp ouput to built file
                                        FileOpen(11, pathnameAm, OpenMode.Input, OpenAccess.ReadWrite, OpenShare.Default)
                                        FileOpen(17, pathnameM, OpenMode.Input, OpenAccess.ReadWrite, OpenShare.Default)
                                        FileOpen(14, pathnamePm, OpenMode.Input, OpenAccess.ReadWrite, OpenShare.Default)
                                        FileOpen(15, pathnameE, OpenMode.Input, OpenAccess.ReadWrite, OpenShare.Default)
                                        FileOpen(16, pathnameNi, OpenMode.Input, OpenAccess.ReadWrite, OpenShare.Default)


                                        'Open pathnameAm For Input As #11
                                        'Open pathnameM For Input As #17
                                        'Open pathnamePm For Input As #14
                                        'Open pathnameE For Input As #15
                                        'Open pathnameNi For Input As #16

                                        Leng = LOF(11)

                                        NextLine = InputString(11, Leng)
                                        pos = InStr(NextLine, orgstr)

                                        If pos > -1 Then
                                            newstr = Replace(NextLine, orgstr, replstr)
                                        End If
                                        pos = InStr(newstr, orgstr2)

                                        If pos > -1 Then
                                            newstr = Replace(newstr, orgstr2, replstr2)
                                        End If

                                        FileClose(11)
                                        FileOpen(11, pathnameAm, OpenMode.Output, OpenAccess.ReadWrite, OpenShare.Default)
                                        'Open pathnameAm For Output As #11
                                        PrintLine(11, newstr)
                                        Leng = LOF(17)

                                        NextLine = InputString(17, Leng)
                                        pos = InStr(NextLine, orgstr)
                                        If pos > -1 Then
                                            newstr = Replace(NextLine, orgstr, replstr)
                                        End If
                                        pos = InStr(newstr, orgstr2)

                                        If pos > -1 Then
                                            newstr = Replace(newstr, orgstr2, replstr2)
                                        End If
                                        FileClose(17)
                                        FileOpen(17, pathnameM, OpenMode.Output, OpenAccess.ReadWrite, OpenShare.Default)
                                        'Open pathnameM For Output As #17
                                        PrintLine(17, newstr)
                                        Leng = LOF(14)

                                        NextLine = InputString(14, Leng)
                                        pos = InStr(NextLine, orgstr)

                                        If pos > -1 Then
                                            newstr = Replace(NextLine, orgstr, replstr)
                                        End If
                                        pos = InStr(newstr, orgstr2)

                                        If pos > -1 Then
                                            newstr = Replace(newstr, orgstr2, replstr2)
                                        End If

                                        FileClose(14)
                                        FileOpen(14, pathnameAm, OpenMode.Output, OpenAccess.ReadWrite, OpenShare.Default)
                                        'Open pathnamePm For Output As #14
                                        PrintLine(14, newstr)
                                        Leng = LOF(15)

                                        NextLine = InputString(15, Leng)
                                        pos = InStr(NextLine, orgstr)

                                        If pos > -1 Then
                                            newstr = Replace(NextLine, orgstr, replstr)
                                        End If
                                        pos = InStr(newstr, orgstr2)

                                        If pos > -1 Then
                                            newstr = Replace(newstr, orgstr2, replstr2)
                                        End If

                                        FileClose(15)
                                        FileOpen(15, pathnameE, OpenMode.Output, OpenAccess.ReadWrite, OpenShare.Default)
                                        'Open pathnameE For Output As #15
                                        Print(15, newstr)
                                        Leng = LOF(16)

                                        NextLine = Input(Leng, 16)
                                        pos = InStr(NextLine, orgstr)

                                        If pos > -1 Then
                                            newstr = Replace(NextLine, orgstr, replstr)
                                        End If
                                        pos = InStr(newstr, orgstr2)

                                        If pos > -1 Then
                                            newstr = Replace(newstr, orgstr2, replstr2)
                                        End If

                                        FileClose(16)
                                        FileOpen(16, pathnameNi, OpenMode.Output, OpenAccess.ReadWrite, OpenShare.Default)
                                        'Open pathnameNi For Output As #16
                                        Print(16, newstr)
                                        SplitM = False 'fixed it in all files
                                    End If

                                    If t = 0 Then
                                        '[041806] hyu: missing @lid (real link's ID) of the shadow link
                                        'tempAttrib = nodes + " " + lflag
                                        tempAttrib = nodes + " " + CStr(pFeat.Value(pFeat.Fields.FindField("Scen_Link"))) + " " + lflag
                                        If fVerboseLog Then WriteLogLine(tempAttrib)
                                    End If

                                    pFeat.Store()
                                    '[033006] hyu: no need to stop editing here
                                    '                                pWSEdit.StopEditOperation
                                    '                                pWSEdit.StopEditing True

                                    If l <> 2 Then 'If Not l = 1 Then
                                        lanes = 1
                                        If t = 0 Then tempstringAm = "a " + nodes + " " + Left(CStr(length), 6) + " " + tempstringAm + " " + CStr(lanes)
                                        If t = 1 Then tempstringM = "a " + nodes + " " + Left(CStr(length), 6) + " " + tempstringM + " " + CStr(lanes)
                                        If t = 2 Then tempstringPm = "a " + nodes + " " + Left(CStr(length), 6) + " " + tempstringPm + " " + CStr(lanes)
                                        If t = 3 Then tempstringE = "a " + nodes + " " + Left(CStr(length), 6) + " " + tempstringE + " " + CStr(lanes)
                                        If t = 4 Then tempstringN = "a " + nodes + " " + Left(CStr(length), 6) + " " + tempstringN + " " + CStr(lanes)
                                    Else
                                        Select Case lanes
                                            Case 2
                                                mode = "ahijb"
                                            Case 3
                                                mode = "aijb"
                                        End Select
                                        lanes = 1
                                        If t = 0 Then tempstringAm = "a " + nodes + " " + Left(CStr(length), 6) + " " + " " + mode + " " + linkType + " " + CStr(lanes)
                                        If t = 1 Then tempstringM = "a " + nodes + " " + Left(CStr(length), 6) + " " + " " + mode + " " + linkType + " " + CStr(lanes)
                                        If t = 2 Then tempstringPm = "a " + nodes + " " + Left(CStr(length), 6) + " " + " " + mode + " " + linkType + " " + CStr(lanes)
                                        If t = 3 Then tempstringE = "a " + nodes + " " + Left(CStr(length), 6) + " " + " " + mode + " " + linkType + " " + CStr(lanes)
                                        If t = 4 Then tempstringN = "a " + nodes + " " + Left(CStr(length), 6) + " " + " " + mode + " " + linkType + " " + CStr(lanes)
                                    End If

                                    If Not pstring = "" Then
                                        If t = 0 Then
                                            Print(1, pstring) 'print for each time timperiod
                                            If Not astring = "" Then Print(3, astring)
                                            If fVerboseLog Then WriteLogLine("ASTRING* " + astring)
                                        End If
                                        If t = 1 Then Print(7, pstring)
                                        If t = 2 Then Print(4, pstring)
                                        If t = 3 Then Print(5, pstring)
                                        If t = 4 Then Print(6, pstring)
                                    End If

                                End If
                            End If

                            If fVerboseLog Then WriteLogLine("checking VDFunc")

                            'NOW assign: a iNode, jNode, length, modes, linkType, #lane, functional class
                            tempString = ""
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
                                    tempString = tempString + " 1 "
                                End If
                            End If

                            If vdf > 0 Then
                                If linkType = "32" Then
                                    If IsDBNull(rLocation) Or rLocation = 0 Then

                                        If t = 0 Or t = 1 Or t = 3 Then tempString = tempString + " " + CStr(vdf - 30)
                                        If t = 2 Then tempString = tempString + " " + CStr(vdf)
                                        If t = 4 Then tempString = tempString + " " + CStr(vdf - 20)
                                    Else

                                        If t = 3 Then vdf = vdf + 10
                                        If t = 4 Then vdf = vdf + 20
                                        tempString = tempString + " " + CStr(vdf)
                                    End If
                                Else
                                    If t = 3 Then vdf = vdf + 10
                                    If t = 4 Then vdf = vdf + 20
                                    tempString = tempString + " " + CStr(vdf)
                                End If
                            End If

                            lanecap = 0

                            'NOW assign: a iNode, jNode, length, modes, linkType, #lane, functional class, capacity
                            If l = 0 Or l = 2 Then 'If l < 2 Then
                                Tindex = pRow.Fields.FindField(direction + "LaneCap" + lType(l))
                            Else
                                Tindex = pRow.Fields.FindField(direction + "LaneCap" + lType(0))
                            End If
                            If Not IsDBNull(pRow.Value(Tindex)) And pRow.Value(Tindex) > 0 Then
                                tempString = tempString + " " + CStr(pRow.Value(Tindex))
                            Else
                                If l = 0 Or l = 2 Then 'If l < 2 Then
                                    Tindex = pMRow.Fields.FindField(direction + "LaneCap" + lType(l))
                                Else
                                    Tindex = pMRow.Fields.FindField(direction + "LaneCap" + lType(0))
                                End If
                                If Not IsDBNull(pMRow.Value(Tindex)) And pMRow.Value(Tindex) > 0 Then
                                    tempString = tempString + " " + CStr(pMRow.Value(Tindex))
                                Else
                                    tempString = tempString + " 1800 "
                                End If

                            End If

                            Tindex = pRow.Fields.FindField(direction + "FFS")
                            If (IsDBNull(pRow.Value(Tindex)) = False And pRow.Value(Tindex) > 0) Then
                                'tempstring = tempstring + " " + CStr(pRow.value(Tindex))
                                'jaf--converted to FormatNumber 12.02.03
                                'tempstring = tempstring + " " + FormatNumber(pRow.value(Tindex), 2, vbTrue)
                                strFFT = FormatNumber(pRow.Value(Tindex), 2, vbTrue)
                                If strFFT = "0.00" Then strFFT = "0.01"
                                tempString = tempString + " " + strFFT
                            Else
                                Tindex = pMRow.Fields.FindField(direction + "FFS")
                                If (IsDBNull(pMRow.Value(Tindex)) = False And pMRow.Value(Tindex) > 0) Then
                                    strFFT = FormatNumber(pMRow.Value(Tindex), 2, vbTrue)
                                    If strFFT = "0.00" Then strFFT = "0.01"
                                    tempString = tempString + " " + strFFT
                                Else
                                    If l = 0 Then
                                        tempString = tempString + " 30 "
                                    Else
                                        tempString = tempString + " 60 "
                                    End If
                                End If
                            End If

                            If t = 0 Then tempstringAm = tempstringAm + " " + tempString
                            If t = 1 Then tempstringM = tempstringM + " " + tempString
                            If t = 2 Then tempstringPm = tempstringPm + " " + tempString
                            If t = 3 Then tempstringE = tempstringE + " " + tempString
                            If t = 4 Then tempstringN = tempstringN + " " + tempString

                            'jaf--write out the Functional Class as model Facility Type

                            Tindex = pFeat.Fields.FindField("FunctionalClass") 'this is changing to E2FacType
                            If Tindex < 0 Then
                                'ERROR--No FacilityType field!!!
                                intFacType = 0

                                'write out error message
                                WriteLogLine("DATA ERROR: GlobalMod.create_NetFile: edge pFeat.OID=" & CStr(pFeat.OID) & " has no Functional field, using 0")
                                g_frmNetLayer.lblStatus.Text = "GlobalMod.create_NetFile: edge pFeat.OID=" & CStr(pFeat.OID) & " has no Functional field, using 0"
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

                            If Not wString = "" Then
                                If fVerboseLog Then WriteLogLine("wstring=" & wString)
                                If t = 0 Then Print(11, wString)
                                If t = 1 Then Print(17, wString)
                                If t = 2 Then Print(14, wString)
                                If t = 3 Then Print(15, wString)
                                If t = 4 Then Print(16, wString)
                            End If

                            If lanes > 0 Then
                                If t = 0 Then Print(11, tempstringAm + " " + CStr(intFacType))
                                If t = 1 Then Print(17, tempstringM + " " + CStr(intFacType))
                                If t = 2 Then Print(14, tempstringPm + " " + CStr(intFacType))
                                If t = 3 Then Print(15, tempstringE + " " + CStr(intFacType))
                                If t = 4 Then Print(16, tempstringN + " " + CStr(intFacType))
                            End If
                        Next t

                        If l = 0 Or Wprint = True Then
                            Print(3, tempAttrib)
                            tempAttrib = ""
                        End If
                        tempString = ""
                    Next l
                End If
            Next i

            '        '[040406] hyu:
            '        lEdgeCt = lEdgeCt + 1
            '        If lEdgeCt = dctEdges.count Then bEOF = True
            '         pFeat = pFeatCursor.NextFeature

            pStatusBar.StepProgressBar()
        Loop 'end of loop through every feature in intermediate edge shapefile
        'now add an additional links from splitting due to weave link case 2

        pWSedit.StopEditOperation()
        pWSedit.StopEditing(True)

        If HaveSplits = True Then
            addSplits_NetFile()
        End If

        '[041406] hyu: code moved to a seperate procedure
        '    Close #11
        '    Close #17
        '    Close #14
        '    Close #15
        '    Close #16
        'now copy temp ouput to built file
        WriteLogLine("copying temp output files to final build files")
        CreateFinalNetfiles(pathnameN, filenameN)
        '    Do Until EOF(11)
        '      Line Input #11, NextLine
        '      If Trim(NextLine) <> "" Then Print #1, NextLine
        '    Loop
        '    Do Until EOF(17)
        '      Line Input #17, NextLine
        '      If Trim(NextLine) <> "" Then Print #7, NextLine
        '    Loop
        '    Do Until EOF(14)
        '      Line Input #14, NextLine
        '      If Trim(NextLine) <> "" Then Print #4, NextLine
        '    Loop
        '    Do Until EOF(15)
        '      Line Input #15, NextLine
        '      If Trim(NextLine) <> "" Then Print #5, NextLine
        '    Loop
        '    Do Until EOF(16)
        '      Line Input #16, NextLine
        '      If Trim(NextLine) <> "" Then Print #6, NextLine
        '    Loop
        '


        'jaf--let's keep the user updated
        g_frmNetLayer.lblStatus.Text = "GlobalMod.create_Netfile: Finished creating Edge portion of Netfile"
        'pan frmNetLayer.Refresh
        'MsgBox "Finished creating Edge portion of Netfile", , "GlobalMod.create_Netfile"
        pStatusBar.HideProgressBar()

        '    Close #11
        '    Close #17
        '    Close #14
        '    Close #15
        '    Close #16
        '    Close #1
        '    Close #7
        '    Close #3
        '    Close #4
        '    Close #5
        '    Close #6
        '    'Close #2
        '    'Close #9

        pFeat = Nothing
        pWorkspace = Nothing
        pTblLine = Nothing
        ' ptblPoint = Nothing
        pTblMode = Nothing
        pEvtLine = Nothing
        pevtPoint = Nothing
        pRow = Nothing

        dctEdges.Clear()
        dctMAtt.Clear()
        dctReservedNodes.Clear()
        dctWeaveNodes.Clear()
        dctWeaveNodes2.Clear()
        dctEdges = Nothing
        dctMAtt = Nothing
        dctReservedNodes = Nothing
        dctWeaveNodes = Nothing
        dctWeaveNodes2 = Nothing

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
        'Close #2
        'Close #9
        FileClose(11)
        FileClose(17)
        FileClose(14)
        FileClose(15)
        FileClose(16)
        FileClose(1)
        FileClose(7)
        FileClose(3)
        FileClose(4)
        FileClose(5)
        FileClose(6)
        pFeat = Nothing
        pWorkspace = Nothing
        pTblLine = Nothing
        ' ptblPoint = Nothing
        pTblMode = Nothing
        pEvtLine = Nothing
        pevtPoint = Nothing
        pRow = Nothing

        If Not dctEdges Is Nothing Then dctEdges.Clear()
        If Not dctMAtt Is Nothing Then dctMAtt.Clear()
        dctReservedNodes.Clear()
        dctWeaveNodes.Clear()
        dctEdges = Nothing
        dctMAtt = Nothing
        dctReservedNodes = Nothing
        dctWeaveNodes = Nothing

    End Sub

Public Sub addSplits_NetFile()
  'on error GoTo eh

        WriteLogLine("called addSplits_NetFile")
        WriteLogLine("------------------------")

   Dim strFFT As String    'holds formatted FFT number
    strFFT = ""

    Dim pStatusBar As IStatusBar
    Dim pProgbar As IStepProgressor
     pStatusBar = m_App.StatusBar
     pProgbar = pStatusBar.ProgressBar

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

    Dim pWorkspace As IWorkspace
     pWorkspace = get_Workspace
    Dim pTblLine As ITable
    'Dim ptblPoint As ITable no longer using this table
    Dim pTblMode As ITable


    'get tables needed for the attribute data need to send to emme/2
     pTblMode = get_TableClass(m_layers(2))
     Dim pTC As ICursor, pTCt As ICursor
    Dim pRow As IRow
    Dim pQFt As IQueryFilter, pQF As IQueryFilter
     pQF = New QueryFilter
    pQF.WhereClause = "SplitHOV = 'HOV' Or SplitTR = 'TR' Or SplitTK = 'TK'"
    Dim tempAttrib As String
    Dim pFeatCursor As IFeatureCursor

    Dim pFeat As IFeature
    Dim pFlds As IFields
    Dim lSFld As Long, upIndex As Long, tempIndex As Long
    Dim Tindex As Long
    Dim i As Long

    Dim count As Long

    Dim lanecap As Double, length As Double
    Dim oneIndex As Long, cnt As Long, shpindex As Long
    Dim fromI As Boolean
    oneIndex = m_edgeShp.FindField("Oneway")

    Dim tempstringAm As String, tempstringM As String, tempstringPm As String, tempstringE As String, tempstringN As String

    'jaf--let's keep the user updated
        g_frmNetLayer.lblStatus.Text = "GlobalMod.addSplits_NetFile: Starting Edge portion of NetFile..."
    'pan frmNetLayer.Refresh

     pFeatCursor = m_edgeShp.Search(pQF, True)

     pFeat = pFeatCursor.NextFeature
    upIndex = pFeat.Fields.FindField("Updated1")
    lSFld = pFeat.Fields.FindField("PSRCEdgeID") 'this used to get out of tblmode
    shpindex = pFeat.Fields.FindField("shptype")
    Dim Wprint As Boolean

     Do Until pFeat Is Nothing 'loop throught all the features in the intermediate edge shapefile
      SplitM = False
            'MsgBox "pfeat oid " + CStr(pFeat.OID)
        'jaf--create the basic links for one-way or two-way facilities based on field Oneway:
        '       0 = one way from I
        '       1 = one way from J
        '       2 = two way
        Dim intWays As Integer
        intWays = 2
            If IsDBNull(pFeat.Value(oneIndex)) Then intWays = 2 Else intWays = pFeat.Value(oneIndex)

        Select Case intWays
            Case 0
                'one way from I
                cnt = 1
                fromI = True
            Case 1
                'one way from J
                cnt = 1
                fromI = False
            Case 2
                'two way
                cnt = 2
                fromI = True
        End Select
               Dim tempString As String, nodes As String
        Dim uindex As Integer
        uindex = m_edgeShp.FindField("UseEmmeN")
        For i = 0 To cnt - 1
            tempString = ""
            nodes = ""
            'jaf--this block writes the inode and jnode values in appropriate direction
            'jaf--(columns inode and jnode)
            If i = 0 Then
                If fromI = True Then
                    Tindex = pFeat.Fields.FindField("INode")
                    If pFeat.value(uindex) = 0 Or pFeat.value(uindex) = 2 Then
                       nodes = nodes + CStr(pFeat.value(Tindex) + m_Offset)
                    Else
                        nodes = nodes + CStr(pFeat.value(Tindex))
                    End If
                    Tindex = pFeat.Fields.FindField("JNode")
                    If pFeat.value(uindex) = 0 Or pFeat.value(uindex) = 1 Then
                       nodes = nodes + " " + CStr(pFeat.value(Tindex) + m_Offset)
                    Else
                        nodes = nodes + " " + CStr(pFeat.value(Tindex))
                    End If
                Else
                    Tindex = pFeat.Fields.FindField("JNode")
                    If pFeat.value(uindex) = 0 Or pFeat.value(uindex) = 1 Then
                       nodes = nodes + CStr(pFeat.value(Tindex) + m_Offset)
                    Else
                        nodes = nodes + CStr(pFeat.value(Tindex))
                    End If
                    Tindex = pFeat.Fields.FindField("INode")
                     If pFeat.value(uindex) = 0 Or pFeat.value(uindex) = 2 Then
                       nodes = nodes + " " + CStr(pFeat.value(Tindex) + m_Offset)
                    Else
                        nodes = nodes + " " + CStr(pFeat.value(Tindex))
                    End If
                End If
            Else
                Tindex = pFeat.Fields.FindField("JNode")
                If pFeat.value(uindex) = 0 Or pFeat.value(uindex) = 1 Then
                       nodes = nodes + CStr(pFeat.value(Tindex) + m_Offset)
                    Else
                        nodes = nodes + CStr(pFeat.value(Tindex))
                    End If
                Tindex = pFeat.Fields.FindField("INode")
                 If pFeat.value(uindex) = 0 Or pFeat.value(uindex) = 2 Then
                       nodes = nodes + " " + CStr(pFeat.value(Tindex) + m_Offset)
                    Else
                        nodes = nodes + " " + CStr(pFeat.value(Tindex))
                    End If
            End If

            'jaf--this block writes the length in miles--6 ASCII CHARS MAX LEN!
            'jaf--(column len(miles))

            nodes = "a " + nodes
            '[040406] hyu: add a judgement before getting Tindex
            If strLayerPrefix = "SDE" Then
                Tindex = pFeat.Fields.FindField("SHAPE.len")
            Else
                Tindex = pFeat.Fields.FindField("SHAPE_length")
            End If
            length = pFeat.value(Tindex)
            length = length * 0.00018939393939
            tempString = tempString + " " + Left(CStr(length), 6)
            
            
            
            'jaf--lane type is derived from projects, if any
            'jaf--assume this all works for now
            Dim lanes As Long
            Tindex = pFeat.Fields.FindField("Scen_Link") 'this links the attribute file (which gets included in the punch) back tot he intermediate shp

             Dim direction As String
           Dim t As Integer, l As Integer
          l = 0 ' just gp

             Wprint = False
           Dim lflag As String
            If (i = 0 And fromI = True) Then
              direction = "IJ"
               lflag = " 1 "      ' flag for IJ

            ElseIf (i = 1 Or fromI = False) Then
                direction = "JI"
                 lflag = " 2 "      ' flag for JI

           End If
            If l = 0 Then
              tempString = nodes + tempString
              nodes = ""
            End If
                 pQFt = New QueryFilter
                pQFt.WhereClause = "PSRCEdgeID = " + CStr(pFeat.value(lSFld))
                 pTC = pTblMode.Search(pQFt, False)

             'jaf--this block writes the modes string, which has defaults if not supplied
            'jaf--(column modes)
            Dim mode As String
            Tindex = pFeat.Fields.FindField("Modes")
                If (pFeat.Value(Tindex) = " " Or IsDBNull(pFeat.Value(Tindex)) = True Or pFeat.Value(Tindex) = "") Then
                    'jaf--default value of modes string
                    mode = "ashi"

                Else
                    'supplied value of modes string
                    mode = CStr(pFeat.Value(Tindex))
                End If
            'jaf--this block writes the linktype number (between 1-99, used to flag screen lines)
            'jaf--(column linktype)
            'jaf--default to 90 for now, we'll have to figure out how to capture this in GeoDB
            Dim linkType As String
            Tindex = pFeat.Fields.FindField("LinkType")
                If (IsDBNull(pFeat.Value(Tindex)) = False And pFeat.Value(Tindex) > 0) Then
                    linkType = pFeat.Value(Tindex)
                Else
                    linkType = "90"
                End If
            'jaf--all times of day buildfiles get same info up to this point
            'assign so far a iNode, jNode, length, modes, linkType if not hov
            If Not l = 1 Then tempString = tempString + " " + mode + " " + linkType
            'else if hov then mode depends on value in lanes field
            tempstringAm = tempString
            tempstringM = tempString
            tempstringPm = tempString
            tempstringE = tempString
            tempstringN = tempString
         'just need the current attributes stored in modeAttributes table
           If (pTblMode.rowcount(pQFt) = 0) Then
                    Print(Err.ToString, "This edge missing from mode attributes: PSRCEdgeID = " + CStr(pFeat.Value(lSFld)))
                Else
                     pRow = pTC.NextRow
                End If

            Dim laneFld As String
            For t = 0 To 4
             'NOW assign: a iNode, jNode, length, modes, linkType, #lane
              laneFld = direction + "Lanes" + lType(l) + timePd(t)

             Tindex = pRow.Fields.FindField(laneFld)
                    If (IsDBNull(pRow.Value(Tindex)) = False And pRow.Value(Tindex) > 0) Then
                        lanes = pRow.Value(Tindex)
                    Else
                        lanes = 1
                    End If
              If t = 0 Then tempstringAm = tempstringAm + " " + CStr(lanes)
             If t = 1 Then tempstringM = tempstringM + " " + CStr(lanes)
                                If t = 2 Then tempstringPm = tempstringPm + " " + CStr(lanes)
                                If t = 3 Then tempstringE = tempstringE + " " + CStr(lanes)
                                If t = 4 Then tempstringN = tempstringN + " " + CStr(lanes)

          ' NOW assign: a iNode, jNode, length, modes, linkType, #lane, functional class
                            tempString = ""
                            Tindex = pRow.Fields.FindField(direction + "VDFunc") 'this will some have by timepd
                             Dim rString As String
                             rString = "r"

                     Dim rLocation As Long

                             rLocation = InStr(mode, rString)
                             Dim vdf As Long
                             vdf = 0
                    If (IsDBNull(pRow.Value(Tindex)) = False And pRow.Value(Tindex) > 0) Then
                        vdf = pRow.Value(Tindex)
                    Else
                        tempString = tempString + " 1 "

                    End If
                            If vdf > 0 Then
                                If linkType = "32" Then
                            If IsDBNull(rLocation) Or rLocation = 0 Then

                                If t = 0 Or t = 1 Or t = 3 Then tempString = tempString + " " + CStr(vdf - 30)
                                If t = 2 Then tempString = tempString + " " + CStr(vdf)
                                If t = 4 Then tempString = tempString + " " + CStr(vdf - 20)
                            Else

                                If t = 3 Then vdf = vdf + 10
                                If t = 4 Then vdf = vdf + 20
                                tempString = tempString + " " + CStr(vdf)
                            End If
                                Else
                                 If t = 3 Then vdf = vdf + 10
                                 If t = 4 Then vdf = vdf + 20
                                 tempString = tempString + " " + CStr(vdf)
                                End If
                            End If

                            lanecap = 0
                            Dim lancapfld As String
                            'NOW assign: a iNode, jNode, length, modes, linkType, #lane, functional class, capacity
                            If l < 2 Then
                              Tindex = pRow.Fields.FindField(direction + "LaneCap" + lType(l))
                            Else
                              Tindex = pRow.Fields.FindField(direction + "LaneCap" + lType(0))
                            End If
                    If Not IsDBNull(pRow.Value(Tindex)) And pRow.Value(Tindex) > 0 Then
                        tempString = tempString + " " + CStr(pRow.Value(Tindex))
                    Else
                        tempString = tempString + " 1800 "

                    End If

                            Tindex = pRow.Fields.FindField(direction + "FFS")
                    If (IsDBNull(pRow.Value(Tindex)) = False And pRow.Value(Tindex) > 0) Then
                        'tempstring = tempstring + " " + CStr(pRow.value(Tindex))
                        'jaf--converted to FormatNumber 12.02.03
                        'tempstring = tempstring + " " + FormatNumber(pRow.value(Tindex), 2, vbTrue)
                        strFFT = FormatNumber(pRow.Value(Tindex), 2, vbTrue)
                        If strFFT = "0.00" Then strFFT = "0.01"
                        tempString = tempString + " " + strFFT
                    Else
                        tempString = tempString + " 30 "

                    End If
                            If t = 0 Then tempstringAm = tempstringAm + " " + tempString
                            If t = 1 Then tempstringM = tempstringM + " " + tempString
                            If t = 2 Then tempstringPm = tempstringPm + " " + tempString
                            If t = 3 Then tempstringE = tempstringE + " " + tempString
                            If t = 4 Then tempstringN = tempstringN + " " + tempString

          'jaf--write out the Functional Class as model Facility Type
            Dim intFacType As Integer
            Tindex = pFeat.Fields.FindField("Functional") 'this is changing to E2FacType
            If Tindex < 0 Then
                'ERROR--No FacilityType field!!!
                intFacType = 0

                'write out error message
                        g_frmNetLayer.lblStatus.Text = "GlobalMod.create_NetFile: edge pFeat.OID=" & CStr(pFeat.OID) & " has no Functional field, using 0"
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

    If lanes > 0 Then
                        If t = 0 Then Print(11, tempstringAm + " " + CStr(intFacType))
                        If t = 1 Then Print(17, tempstringM + " " + CStr(intFacType))
                        If t = 2 Then Print(14, tempstringPm + " " + CStr(intFacType))
                        If t = 3 Then Print(15, tempstringE + " " + CStr(intFacType))
                        If t = 4 Then Print(16, tempstringN + " " + CStr(intFacType))

          End If
        Next t
             tempString = ""

        Next i
         pFeat = pFeatCursor.NextFeature
        pStatusBar.StepProgressBar
    Loop 'end of loop through every feature in intermediate edge shapefile

    'jaf--let's keep the user updated
        g_frmNetLayer.lblStatus.Text = "GlobalMod.create_Netfile: Finished creating Edge portion of Netfile"
    'pan frmNetLayer.Refresh
    'MsgBox "Finished creating Edge portion of Netfile", , "GlobalMod.create_Netfile"
    pStatusBar.HideProgressBar

     pFeat = Nothing
     pWorkspace = Nothing

     pTblMode = Nothing

     pRow = Nothing

    Exit Sub
eh:

        CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.addSplits_NetFile")
        MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.addSplits_NetFile")

  'MsgBox Err.Description, vbExclamation, "GlobalMod.addSplits_NetFile"

   pFeat = Nothing
   pWorkspace = Nothing
   pTblMode = Nothing
   pRow = Nothing

End Sub


    Public Sub create_TurnFile(ByVal pathnameN As String, ByVal filenameN As String, ByVal pInterLayerE As IFeatureLayer, ByVal modelYEAR As Long, ByVal pTurnLayer As IFeatureLayer)
        'this has to be called by createScenarioShapefile
        'pan problem identified 10-11-05 not limited to chosen subnetwork not fixed yet
        'on error GoTo eh
        ' Dim pathname As String
        ' pathname = "c:\createTurn_ERROR.txt"
        '  Open pathname For Output As #11
        '  Print #11, "Create_TurnFile error log"
        WriteLogLine("")
        WriteLogLine("===================================")
        WriteLogLine("create_TurnFile started " & Now())
        WriteLogLine("===================================")
        WriteLogLine("")

        Dim runtype(5) As String
        runtype(0) = "am"
        runtype(1) = "md"
        runtype(2) = "pm"
        runtype(3) = "ev"
        runtype(4) = "ni"

        'Dim pStatusBar As IStatusBar
        'Dim pProgbar As IStepProgressor
        'pStatusBar = m_App.StatusBar
        'pProgbar = pStatusBar.ProgressBar

        Dim thepath(5) As String
        Dim tempfield As String
        Dim tempString As String
        Dim attrib(4) As String

        thepath(0) = pathnameN + "\" + runtype(0) + "_" + filenameN
        thepath(1) = pathnameN + "\" + runtype(1) + "_" + filenameN
        thepath(2) = pathnameN + "\" + runtype(2) + "_" + filenameN
        thepath(3) = pathnameN + "\" + runtype(3) + "_" + filenameN
        thepath(4) = pathnameN + "\" + runtype(4) + "_" + filenameN
        'thepath(4) = pathnameN + "\" + filenameN + "TurnPenN.txt"

        WriteLogLine("Done with init, opening output files")

        FileClose(1, 2, 3, 4, 5)
        FileOpen(1, thepath(0), OpenMode.Output, OpenAccess.Default, OpenShare.Shared)
        FileOpen(2, thepath(1), OpenMode.Output, OpenAccess.Default, OpenShare.Shared)
        FileOpen(3, thepath(2), OpenMode.Output, OpenAccess.Default, OpenShare.Shared)
        FileOpen(4, thepath(3), OpenMode.Output, OpenAccess.Default, OpenShare.Shared)
        FileOpen(5, thepath(4), OpenMode.Output, OpenAccess.Default, OpenShare.Shared)

        'Open thepath(0) For Output As 1
        'Open thepath(1) For Output As 2
        'Open thepath(2) For Output As 3
        'Open thepath(3) For Output As 4
        'Open thepath(4) For Output As 5

        PrintLine(1, "c Exported Turn Penality File from ArcMap /Emme/2 Interface for Am: ") '+ CStr(MyDate)
        PrintLine(1, "t turns init")
        PrintLine(2, "c Exported Turn Penality File from ArcMap /Emme/2 Interface for Midday: ") '+ CStr(MyDate)
        PrintLine(2, "t turns init")
        PrintLine(3, "c Exported Turn Penality File from ArcMap /Emme/2 Interface for PM: ") '+ CStr(MyDate)
        PrintLine(3, "t turns init")
        PrintLine(4, "c Exported Turn Penality File from ArcMap /Emme/2 Interface for Evening: ") '+ CStr(MyDate)
        PrintLine(4, "t turns init")
        PrintLine(5, "c Exported Turn Penality File from ArcMap /Emme/2 Interface for Night: ") '+ CStr(MyDate)
        PrintLine(5, "t turns init")

        WriteLogLine("Done writing file headers, commencing layer setup")

        Dim pFeatLayerE As IFeatureLayer 'pTurnLayer As IFeatureLayer
        ' pFeatLayerE = get_MasterNetwork(0) 'need to use intermediate edge because events might split edges
        pFeatLayerE = pInterLayerE
        ' pTurnLayer = get_FeatureLayer(m_layers(12)) 'Turn Movements
        Dim pQFturn As IQueryFilter, pQF As IQueryFilter
        Dim pFeature As IFeature, pFeatTurn As IFeature, pFeatE As IFeature
        Dim FromN As Long, ToN As Long
        Dim Frome As Long, Toe As Long
        Dim i As Long

        Dim pFeatCursor As IFeatureCursor
        Dim pFCturn As IFeatureCursor, pFCedge As IFeatureCursor, pFCtemp As IFeatureCursor
        Dim pFilter As ISpatialFilter

        pFeatCursor = m_junctShp.Search(Nothing, False)
        pFeature = pFeatCursor.NextFeature

        Dim index As Long
        Dim idIndex As Long

        Dim lastFrom As Long, lastto As Long

        'Do Until pFeature Is Nothing


        lastFrom = 0
        lastto = 0
        'now get it in turnmovements- should return more than one enter
        pQFturn = New QueryFilter
        '    pQFturn.WhereClause = "PSRCJunctID = " + CStr(pFeature.value(idIndex))
        'If (pTurnLayer.FeatureClass.featurecount(pQFturn) > 0) Then
        'MsgBox "number of turns" + CStr(pTurnLayer.FeatureClass.featurecount(Nothing))

        '[041106] hyu: narrow cursor to only modelyear records
        pQFturn.WhereClause = "InServiceDate<=" & modelYEAR & " AND OutServiceDate>" & modelYEAR
        pFCturn = pTurnLayer.Search(pQFturn, False)
        ' pFCturn = pTurnLayer.Search(Nothing, False)
        pFeatTurn = pFCturn.NextFeature
        Dim pGeometry As IGeometry

        WriteLogLine("Done with layer setup and query, commencing iterations")

        Do Until pFeatTurn Is Nothing

            idIndex = pTurnLayer.FeatureClass.FindField("PSRCJunctID") 'this is PSRC_junctID
            'pan 11-15-05 this should include m_Offset --1200
            '[042606] pan: ok so now I put it in
            'tempstring = "a " + CStr(pFeatTurn.value(idIndex)) + " "
            tempString = "a " + CStr(pFeatTurn.Value(idIndex) + m_Offset) + " "
            pFilter = New SpatialFilter
            ' pGeometry = pFeatTurn.Shape
            ' With pFilter
            '      .Geometry = pGeometry
            '     .GeometryField = pFeatLayerE.FeatureClass.ShapeFieldName 'pTurnLayer.FeatureClass.ShapeFieldName
            '     .SpatialRel = esriSpatialRelContains 'esriSpatialRelIntersects
            ' End With
            '  pFCedge = pFeatLayerE.Search(pFilter, False)
            'MsgBox pFeatLayerE.FeatureClass.featurecount(pFilter)
            Frome = 0
            Toe = 0
            '  pFeatE = pFCedge.NextFeature
            ''Do Until pFeatE Is Nothing
            '  index = pFeatE.Fields.FindField("INode")
            '  If (pFeatE.value(index) = pFeature.value(idIndex)) Then
            'MsgBox "found toedge"
            '       Toe = pFeatE.OID
            '   End If
            '   index = pFeatE.Fields.FindField("JNode")
            '   If (pFeatE.value(index) = pFeature.value(idIndex)) Then
            'MsgBox "find fromedge"
            '      Frome = pFeatE.OID
            '    End If
            '     pFeatE = pFCedge.NextFeature

            'Loop
            idIndex = pFeatTurn.Fields.FindField("ToEdgeID")
            Toe = pFeatTurn.Value(idIndex)
            idIndex = pFeatTurn.Fields.FindField("FrEdgeID")
            Frome = pFeatTurn.Value(idIndex)

            'Dim temp As Long
            'if (Frome <> 0 Or Toe <> 0) Then
            'need to get spatial filter that intersects!
            '    pFilter = New SpatialFilter
            '   With pFilter
            '        .Geometry = pFeatTurn.ShapeCopy
            '       .GeometryField = pFeatLayerE.FeatureClass.ShapeFieldName 'pTurnLayer.FeatureClass.ShapeFieldName
            '       .SpatialRel = esriSpatialRelIntersects
            '  End With
            '   pFCedge = pFeatLayerE.Search(pFilter, False)
            'MsgBox pFeatLayerE.FeatureClass.featurecount(pFilter)
            '   pFeatE = pFCedge.NextFeature
            'Do Until pFeatE Is Nothing
            '   index = pFeatE.Fields.FindField("INode")
            '   If (pFeatE.value(index) = pFeature.value(idIndex) And Toe = 0) Then
            'MsgBox "found toedge"
            '       Toe = pFeatE.OID

            '  End If
            '   index = pFeatE.Fields.FindField("JNode")
            '   If (pFeatE.value(index) = pFeature.value(idIndex) And Frome = 0) Then
            '      Frome = pFeatE.OID
            '   End If

            '    pFeatE = pFCedge.NextFeature

            'Loop
            'End If
            If (Frome <> 0 And Toe <> 0) Then 'its in scenario
                'getfrom node
                ' pFeatE = pFeatLayerE.FeatureClass.GetFeature(Frome)
                pQF = New QueryFilter
                pQF.WhereClause = "PSRCEdgeID = " + CStr(Frome)
                If (pFeatLayerE.FeatureClass.FeatureCount(pQF) < 1) Then
                    WriteLogLine("DATA ERROR: PSRCEdgeID = " + CStr(Frome) + " is missing from TransRefEdges--create_TurnFile")
                Else
                    pFCedge = pFeatLayerE.Search(pQF, False)
                    pFeatE = pFCedge.NextFeature
                    index = pFeatE.Fields.FindField("INode")
                    FromN = pFeatE.Value(index)

                    idIndex = pTurnLayer.FeatureClass.FindField("PSRCJunctID")
                    If (FromN = pFeatTurn.Value(idIndex)) Then
                        index = pFeatE.Fields.FindField("JNode")
                        FromN = pFeatE.Value(index)
                    End If

                    'get to node
                    ' pFeatE = pFeatLayerE.FeatureClass.GetFeature(Toe)
                    pQF = New QueryFilter
                    pQF.WhereClause = "PSRCEdgeID = " + CStr(Toe)
                    If (pFeatLayerE.FeatureClass.FeatureCount(pQF) < 1) Then
                        WriteLogLine("DATA ERROR: PSRCEdgeID = " + CStr(Toe) + " is missing from TransRefEdges--create_TurnFile")
                    Else
                        pFCedge = pFeatLayerE.Search(pQF, False)
                        pFeatE = pFCedge.NextFeature
                        index = pFeatE.Fields.FindField("JNode")
                        ToN = pFeatE.Value(index)
                        If (ToN = pFeatTurn.Value(idIndex)) Then
                            index = pFeatE.Fields.FindField("INode")
                            ToN = pFeatE.Value(index)
                        End If
                        '[042606] pan:  added m_offset
                        tempString = tempString + CStr((FromN) + m_Offset) + " " + CStr((ToN) + m_Offset) + " "
                        lastFrom = FromN
                        lastto = ToN

                        'MsgBox tempstring
                        For i = 0 To 4

                            tempfield = "Function" + runtype(i)
                            index = pFeatTurn.Fields.FindField(tempfield)
                            If (IsDBNull(pFeatTurn.Value(index)) = True) Then
                                attrib(i) = " 0 "
                            Else
                                attrib(i) = CStr(pFeatTurn.Value(index)) + " "
                            End If

                            tempfield = "user1" + runtype(i)
                            index = pFeatTurn.Fields.FindField(tempfield)
                            If (IsDBNull(pFeatTurn.Value(index)) = True) Then
                                attrib(i) = attrib(i) + " 0 "
                            Else
                                attrib(i) = attrib(i) + CStr(pFeatTurn.Value(index)) + " "
                            End If
                            tempfield = "user2" + runtype(i)
                            index = pFeatTurn.Fields.FindField(tempfield)
                            If (IsDBNull(pFeatTurn.Value(index)) = True) Then
                                attrib(i) = attrib(i) + " 0 "
                            Else
                                attrib(i) = attrib(i) + CStr(pFeatTurn.Value(index)) + " "
                            End If
                            tempfield = "user3" + runtype(i)
                            index = pFeatTurn.Fields.FindField(tempfield)
                            If (IsDBNull(pFeatTurn.Value(index)) = True) Then
                                attrib(i) = attrib(i) + " 0 "
                            Else
                                attrib(i) = attrib(i) + CStr(pFeatTurn.Value(index)) + " "
                            End If
                        Next i
                        attrib(0) = tempString + attrib(0)

                        PrintLine(1, attrib(0))
                        attrib(1) = tempString + attrib(1)
                        PrintLine(2, attrib(1))
                        attrib(2) = tempString + attrib(2)
                        PrintLine(3, attrib(2))
                        attrib(3) = tempString + attrib(3)
                        PrintLine(4, attrib(3))
                        attrib(4) = tempString + attrib(4)
                        PrintLine(5, attrib(4))

                    End If
                End If
            End If

            ' tempstring = "a " + CStr(pFeature.value(idIndex)) + " "
            pFeatTurn = pFCturn.NextFeature
            'Loop
            'Else
            ' MsgBox "Could not find the turn penality for: " + CStr(pFeature.value(idIndex))
            'End If

            ' pFeature = pFeatCursor.NextFeature
            System.Runtime.InteropServices.Marshal.ReleaseComObject(pFCedge)
        Loop
        'System.Runtime.InteropServices.Marshal.ReleaseComObject(pFeatCursor)
        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(pFeatCursor)
        'SEC: 04/20/2009
        'The following code creates turn penalties to prevent traffic from going directly to a Freeway/Highway HOV Lane from an on-ramp or some other
        'facility. Best example is a ramp to freeway link where the vehicle needs to get on
        'the highway before switching over to the HOV lanes. Same applies to Non HOV only /direct access ramps that have
        ' HOV lanes. The following code should not be applied to HOV only facilities and direct access ramps.
        WriteLogLine("Getting HOV Turn Penalties " & Now())
        Dim pHOVQF As IQueryFilter
        Dim pJNodeQF As IQueryFilter
        Dim pINodeQF As IQueryFilter
        Dim pEdgeIDQF As IQueryFilter
        Dim pModeCursor As ICursor
        Dim pRow As IRow

        Dim pHOVFC As IFeatureCursor
        Dim pHOVFeat As IFeature
        Dim indexFacType As Long
        Dim indexINode As Long
        Dim indexJNode As Long
        Dim indexIWeaveNode As Long
        Dim indexEdgeID As Long
        Dim indexJWeaveNode As Long
        indexFacType = pInterLayerE.FeatureClass.FindField("FacilityType")
        indexINode = pInterLayerE.FeatureClass.FindField("INode")
        indexIWeaveNode = pInterLayerE.FeatureClass.FindField("HOV_I")
        indexJWeaveNode = pInterLayerE.FeatureClass.FindField("HOV_J")
        indexJNode = pInterLayerE.FeatureClass.FindField("JNode")
        indexEdgeID = pInterLayerE.FeatureClass.FindField("PSRCEdgeID")

        'Using pComReleaser As New ComReleaser

        pHOVQF = New QueryFilter
        'pComReleaser.ManageLifetime(pHOVQF)
        'Only want to do this to Freeways, State Routes, etc that have HOV lanes and are OnewayIJ, not reversibles
        pHOVQF.WhereClause = "HOV_I > 0 AND FacilityType < 4 And Oneway = 0"
        pHOVFC = pInterLayerE.Search(pHOVQF, False)
        'pComReleaser.ManageLifetime(pInterLayerE)
        pHOVFeat = pHOVFC.NextFeature




        Try


            Do Until pHOVFeat Is Nothing
                Dim pEdgeID
                Dim pHOVEdgeID As Long
                Dim pHOVfeatFacType As Integer
                Dim pHOVINode As Long
                Dim pINode As Long
                Dim pJNode As Long
                Dim pHOVIWeaveNode As Long
                Dim pJNodeFeatCurs As IFeatureCursor
                Dim pJNodeFeat As IFeature


                'Dim pTblMode As ITable
                'pTblMode = get_TableClass(m_layers(2))
                'get the edge id
                pHOVEdgeID = pHOVFeat.Value(indexEdgeID)

                'get the HOV edge's Facility Type
                pHOVfeatFacType = pHOVFeat.Value(indexFacType)
                pHOVIWeaveNode = pHOVFeat.Value(indexIWeaveNode)
                'want to find edges connecting to the HOVEdge, which are any edges with a Jnode = to the HOVEdge's INode:
                pHOVINode = pHOVFeat.Value(indexINode)
                pJNodeQF = New QueryFilter
                pJNodeQF.WhereClause = "JNode = " & pHOVINode
                pJNodeFeatCurs = pInterLayerE.Search(pJNodeQF, False)
                pJNodeFeat = pJNodeFeatCurs.NextFeature

                'For ramps to highways
                Do Until pJNodeFeat Is Nothing
                    pEdgeID = pJNodeFeat.Value(indexEdgeID)
                    'make sure this edge is GP:
                    pEdgeIDQF = New QueryFilter
                    pEdgeIDQF.WhereClause = "PSRCEdgeID = " & pEdgeID
                    pModeCursor = gTblModeAtt.Search(pEdgeIDQF, False)
                    pRow = pModeCursor.NextRow
                    'looking for a different facility type, e.g. ramp and must NOT be an HOV only facility (we want them to be able to access the hov system directly in most cases)
                    If pJNodeFeat.Value(indexFacType) <> pHOVfeatFacType And pHOVfeatFacType <> 18 And pHOVfeatFacType <> 19 Then
                        'get the connected feature's Inode
                        pINode = pJNodeFeat.Value(indexINode)

                        'write out the penalty in the build file
                        tempString = "a " + CStr(pHOVINode + m_Offset) + " " + CStr(pINode + m_Offset) + " " + CStr(pHOVIWeaveNode + m_Offset) + " 0 0 0 0"
                        PrintLine(1, tempString)
                        PrintLine(2, tempString)
                        PrintLine(3, tempString)
                        PrintLine(4, tempString)
                        PrintLine(5, tempString)


                        'Also, check to see if Ramp has HOV. If so, we need to build a turn penality from Ramps HOV to Freeway's HOV
                        'Direct Access are/should be excluded by checking for GP lanes
                        If Not IsDBNull(pJNodeFeat.Value(indexJWeaveNode)) Then
                            If pJNodeFeat.Value(indexJWeaveNode) > 0 Then
                                tempString = "a " + CStr(pJNodeFeat.Value(indexJWeaveNode) + m_Offset) + " " + CStr(pJNodeFeat.Value(indexIWeaveNode) + m_Offset) + " " + CStr(pHOVFeat.Value(indexJWeaveNode) + m_Offset) + " 0 0 0 0"
                                PrintLine(1, tempString)
                                PrintLine(2, tempString)
                                PrintLine(3, tempString)
                                PrintLine(4, tempString)
                                PrintLine(5, tempString)

                            End If
                        End If

                    End If




                    pJNodeFeat = pJNodeFeatCurs.NextFeature
                Loop
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(pJNodeFeatCurs)
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(pModeCursor)
                ' End If

                'Dim pEdgeID
                'Dim pHOVEdgeID As Long
                'Dim pHOVfeatFacType As Integer
                Dim pHOVJNode As Long
                'Dim pINode As Long
                'Dim pJNode As Long
                Dim pHOVJWeaveNode As Long
                Dim pINodeFeatCurs As IFeatureCursor
                Dim pINodeFeat As IFeature


                'Dim pTblMode As ITable
                'Set pTblMode = get_TableClass(m_layers(2))
                'get the edge id
                'pHOVEdgeID = pHOVFeat.value(indexEdgeID)

                'get the HOV edge's Facility Type
                'pHOVfeatFacType = pHOVFeat.value(indexFacType)
                pHOVJWeaveNode = pHOVFeat.Value(indexJWeaveNode)
                'want to find edges connecting to the HOVEdge, which are any edges with a Jnode = to the HOVEdge's INode:
                pHOVJNode = pHOVFeat.Value(indexJNode)
                pINodeQF = New QueryFilter
                pINodeQF.WhereClause = "INode = " & pHOVJNode
                pINodeFeatCurs = pInterLayerE.Search(pINodeQF, True)
                pINodeFeat = pINodeFeatCurs.NextFeature


                'for highways to ramps:
                Do Until pINodeFeat Is Nothing
                    pEdgeID = pINodeFeat.Value(indexEdgeID)
                    'make sure this edge is GP:
                    pEdgeIDQF = New QueryFilter
                    pEdgeIDQF.WhereClause = "PSRCEdgeID = " & pEdgeID
                    'pModeCursor = gTblModeAtt.Search(pEdgeIDQF, False)
                    'pRow = pModeCursor.NextRow
                    'looking for a different facility type, e.g. ramp and must NOT be an HOV only facility (we want them to be able to access the hov system directly in most cases)
                    If pINodeFeat.Value(indexFacType) <> pHOVfeatFacType And pHOVfeatFacType <> 18 And pHOVfeatFacType <> 19 Then
                        'get the connected feature's Inode
                        pJNode = pINodeFeat.Value(indexJNode)

                        'write out the penalty in the build file
                        tempString = "a " + CStr(pHOVJNode + m_Offset) + " " + CStr(pHOVJWeaveNode + m_Offset) + " " + CStr(pJNode + m_Offset) + " " + " 0 0 0 0"
                        PrintLine(1, tempString)
                        PrintLine(2, tempString)
                        PrintLine(3, tempString)
                        PrintLine(4, tempString)
                        PrintLine(5, tempString)


                        'Also, check to see if Ramp has HOV. If so, we need to build a turn penality from Ramps HOV to Freeway's HOV
                        'Direct Access are/should be excluded by checking for GP lanes
                        If Not IsDBNull(pINodeFeat.Value(indexIWeaveNode)) Then
                            If pINodeFeat.Value(indexIWeaveNode) > 0 Then

                                tempString = "a " + CStr(pINodeFeat.Value(indexIWeaveNode) + m_Offset) + " " + CStr(pHOVFeat.Value(indexIWeaveNode) + m_Offset) + " " + CStr(pINodeFeat.Value(indexJWeaveNode) + m_Offset) + " " + " 0 0 0 0"
                                PrintLine(1, tempString)
                                PrintLine(2, tempString)
                                PrintLine(3, tempString)
                                PrintLine(4, tempString)
                                PrintLine(5, tempString)

                            End If
                        End If

                    End If




                    pINodeFeat = pINodeFeatCurs.NextFeature
                Loop
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(pINodeFeatCurs)













                pHOVFeat = pHOVFC.NextFeature
            Loop
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(pHOVFC)
            'Marshal.FinalReleaseComObject(pHOVFC)
        Catch ex As Exception
            MessageBox.Show(ex.ToString)

        End Try
        'End Using




        WriteLogLine("done writing, closing files")
        'pStatusBar.HideProgressBar()
        FileClose(1)
        FileClose(2)
        FileClose(3)
        FileClose(4)
        FileClose(5)
        'Close #11
        pFeatLayerE = Nothing
        pTurnLayer = Nothing
        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(pFCturn)
        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(pFCedge)
        'System.Runtime.InteropServices.Marshal.ReleaseComObject(pFCtemp)
        GC.Collect()



        WriteLogLine("FINISHED create_TurnFile at " & Now())
        Exit Sub

eh:
        CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.create_TurnFile")
        MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.create_TurnFile")
        'Print #11, Err.Description, vbExclamation, "TurnFile"
        FileClose(1)
        FileClose(2)
        FileClose(3)
        FileClose(4)
        FileClose(5)
        'Close #11
        pFeatLayerE = Nothing
        pTurnLayer = Nothing
        pFeature = Nothing

    End Sub

Public Sub create_ScenarioRow()
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
        pWorkspaceEdit.StartEditOperation
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
    If (pScenTable.rowcount(Nothing) > 0) Then
         pTC = pScenTable.Search(Nothing, False)
         pRow = pTC.NextRow
        Do Until pRow Is Nothing
            If (pRow.value(index) > count) Then
                count = pRow.value(index)
            End If
             pRow = pTC.NextRow
        Loop
    End If
    m_ScenarioId = count + 1
    pNewRow.value(index) = m_ScenarioId

    index = pScenTable.FindField("Title")
    If (index = -1) Then
            WriteLogLine("Could not find field 'Title' in tblModelScenario")
        'Exit Sub
    End If
    pNewRow.value(index) = m_title
    index = pScenTable.FindField("Description")
    If (index = -1) Then
            WriteLogLine("Could not find field 'Description' in tblModelScenario")
        'Exit Sub
    End If
    pNewRow.value(index) = m_descript
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
    pNewRow.value(index) = m_Offset

    pNewRow.Store
     pNewRow = Nothing
     pScenTable = Nothing

    If (strLayerPrefix = "SDE") Then
      pWorkspaceEdit.StopEditOperation
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

Public Sub create_ProjectScenRows()
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
        pWorkspaceEdit.StartEditOperation

     pInScenTable = get_TableClass(m_layers(15)) 'tblprojectsinscenarios
     pTblLine = get_TableClass(m_layers(3))
   '  ptblPoint = get_TableClass("tblPointProjects", pWs)

    Dim index As Long, dbindex As Long
    Dim count As Long, j As Long
    Dim tempprj As ClassPrjSelect
    Dim tempString As String


    For j = 1 To prjselectedCol.count
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

        If (pTblLine.rowcount(pQF) > 0) Then
           ' MsgBox "found it in line"
             pNewRow = pInScenTable.CreateRow
            pNewRow.value(index) = m_ScenarioId
            
             pTC = pTblLine.Search(pQF, False)
             pRow = pTC.NextRow

            index = pInScenTable.FindField("projID_Ver")
            If (index = -1) Then
                    WriteLogLine("Could not find field 'projID_Ver' in tblProjectsInScenario")
                'Exit Sub
            Else
                pNewRow.value(index) = pRow.value(pRow.Fields.FindField("projID_Ver"))
            End If
            '[080808] SEC: Here, the index is pointing to the right field but wrong table! Should be pInScenTable not
            'pTblLine
            'index = pTblLine.FindField("projDBS")
            index = pInScenTable.FindField("projDBS")
            If (dbindex = -1) Then
                    WriteLogLine("Could not find field 'projDBS' in tblLineProjects")
                'Exit Sub
            Else
            
                pNewRow.value(index) = pRow.value(pRow.Fields.FindField("projDBS"))
            End If
'             pTC = pTblLine.Search(pQF, False)
'             pRow = pTC.NextRow
            
            index = pInScenTable.FindField("projID")
            If (index = -1) Then
                    WriteLogLine("Could not find field 'projID' in tblProjectsInScenario")
                'Exit Sub
            Else
                pNewRow.value(index) = pRow.value(pRow.Fields.FindField("projID"))
            End If
            
            index = pInScenTable.FindField("projRteID")
            If (index = -1) Then
                    WriteLogLine("Could not find field 'projRteID' in tblProjectsInScenario")
                'Exit Sub
            Else
                pNewRow.value(index) = tempString
            End If
            
            index = pInScenTable.FindField("version")
            If (index = -1) Then
                    WriteLogLine("Could not find field 'version' in tblProjectsInScenario")
                'Exit Sub
            Else
                pNewRow.value(index) = pRow.value(pRow.Fields.FindField("version"))
            End If
            
            index = pInScenTable.FindField(g_InSvcDate)
            If (index = -1) Then
                    WriteLogLine("Could not find field '" & g_InSvcDate & "' in tblProjectsInScenario")
                'Exit Sub
            Else
                pNewRow.value(index) = pRow.value(pTblLine.FindField(g_InSvcDate))
            End If
        'ElseIf (ptblPoint.rowcount(pQF) > 0) Then
           'deleted code 01-19-05
            pNewRow.Store
        Else
        '[021506] hyu: shouldn 't the pNewRow be deleted, if no project found in tbl_LineProjects?
                WriteLogLine("Could not find the project in tbl_LineProjects " + CStr(tempprj.PrjId))
        End If
       
    Next j

    For j = 1 To evtselectedCol.count
         pQF = New QueryFilter
         tempprj = evtselectedCol.Item(j)
        tempString = tempprj.PrjId
        'MsgBox tempstring
        '[100207] hyu: search on "projRteID" instead of "projID_Ver"
        'pQF.WhereClause = "projID_Ver = '" + tempString + "'"
        pQF.WhereClause = "projRteID = " + tempString
        index = pInScenTable.FindField("Scenario_ID")
         pNewRow = pInScenTable.CreateRow
        pNewRow.value(index) = m_ScenarioId
        'now just been saved in tblModelScenario
       ' index = pInScenTable.FindField("Title")
       ' pNewRow.value(index) = m_title

        If (pTblLine.rowcount(pQF) > 0) Then
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
             pNewRow.value(index) = tempString
'            End If
        'ElseIf (ptblPoint.rowcount(pQF) > 0) Then
            'deleted code 01-19-05
        Else
                WriteLogLine("Could not find the project in tbl_LineProjects" + CStr(tempprj.PrjId))
        End If
       pNewRow.Store
    Next j

    'If (strLayerPrefix = "SDE") Then
       pWorkspaceEdit.StopEditOperation
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

'Public Sub Open_Shapefiles(edgepath As String, edgename As String) 'didn't really need junction shapefile, junctname As String, junctpath As String)
Public Sub Open_Shapefiles() '(edgepath As String, edgename As String) 'didn't really need junction shapefile, junctname As String, junctpath As String)
   'called by frmOutNetSelect in the Output Prep Module
   'jaf--it appears that frmOutNetSelect stays visible until this routine completes
   'jaf--
   '    edgepath is path to intermediate edge shapefile
   '    edgename is filename of intermediate edge shapefile

   'on error GoTo eh

   'jaf--let user know what's going on
'   frmOutNetSelect.lblStatus = "GlobalMod.Open_Shapefiles: checking " & edgename
   'pan frmOutNetSelect.Refresh
        WriteLogLine("called Open_Shapefiles")
        WriteLogLine("----------------------")

   Dim filename As String
   Dim extension As String
   Dim ext As String
   ext = "."
   Dim stringlen As Long, position As Long

   Dim pFact As IWorkspaceFactory
   Dim pWorkspace As IWorkspace
    
    '[hyu 091906]: get workspace from the loaded data
     pWorkspace = getWorkspaceOfWorkVersion 'setSDEWorkspace
    
   Dim pFeatWorkspace As IFeatureWorkspace
    pFeatWorkspace = pWorkspace
   
    '[092106 hyu] comment out this block. not usefule.
'   stringlen = Len(edgename)
'   position = InStrRev(edgename, ext, , vbTextCompare)
'   filename = Left(edgename, (position - 1))
'   extension = Right(edgename, (stringlen - position))

   'jaf--m_edgeShp is handle to edge intermediate shapefile
'    m_edgeShp = pFeatWorkspace.OpenFeatureClass(filename)
     m_edgeShp = pFeatWorkspace.OpenFeatureClass(m_layers(22))
    
   'get scenario id from the edge intermediate shapefile (have to get a feature first)
   Dim index As Long
   index = m_edgeShp.FindField("ScenarioID")
   Dim pFC As IFeatureCursor
   Dim pFeature As IFeature
   
   '[hyu 092106] direct call the first feature. search for nothing is poor at performance.
   Dim i As Long
   'jaf-- feature cursor pFC to everything in the edge int. shapefile
    pFC = m_edgeShp.Search(Nothing, False)
    pFeature = pFC.NextFeature

   'jaf-- scenarioID from first feature
   m_ScenarioId = pFeature.value(index)
   'MsgBox m_ScenarioId
   'need to union any split features
   'pan--This union operation is the inverse of thinning prior to creating Net file?
   index = m_edgeShp.FindField("PSRC_E2ID")

'   Dim pQF As IQueryFilter
'   Dim pUFeat As IFeature
'   Dim pFCU As IFeatureCursor
'   Dim pGeom As IGeometry
'   Dim m_Bag As IGeometryCollection
'   Dim pPolyline As IPolyline
'   Dim pLine As ILine
'   Dim pTopoOp As ITopologicalOperator
'   Dim pNewFeature As IFeature
'   Dim pFields As IFields
'   Dim FieldCount As Long
'   Dim pField As IField
'
'   'jaf--loop through all features in intermediate edge shapefile
'   Do Until pfeature Is Nothing
'       pQF = New QueryFilter
'      If (pfeature.value(Index) <> 0) Then
'         'jaf--feature has a nonzero PSRC_E2ID which is a good thing
'         pQF.WhereClause = "PSRC_E2ID = " + CStr(pfeature.value(Index))
'         If (m_edgeShp.featurecount(pQF) > 1) Then ' need to union
'            'jaf--more than one edge with same PSRC_E2ID: Melinda says need to union
''**************************************************************************
''jaf--why are we doing this?
''**************************************************************************
'            'MsgBox "Unioning int. shp edges with PSRC_E2ID=" & CStr(pFeature.value(index)), , "GlobalMod.Open_Shapefiles"
'            'MsgBox "unioning"
'
'            'jaf--let user know what's going on
'            frmOutNetSelect.lblStatus = "GlobalMod.Open_Shapefiles: Unioning int. shp edges with PSRC_E2ID=" & CStr(pfeature.value(Index))
'            'pan frmOutNetSelect.Refresh
'
'             m_Bag = New GeometryBag
'            'jaf-- featurecursor pFCU to all edges with this PSRC_E2ID
'             pFCU = m_edgeShp.Search(pQF, False)
'            'jaf--loop through all edges that need unioning
'             pUFeat = pFCU.NextFeature
'            Do Until pUFeat Is Nothing
'               m_Bag.AddGeometry pUFeat.Shape
'                pUFeat = pFCU.NextFeature
'            Loop
'             pPolyline = New Polyline
'             pTopoOp = pPolyline
'            pTopoOp.ConstructUnion m_Bag
'            'jaf--create the unioned feature
'             pNewFeature = m_edgeShp.CreateFeature
'             pNewFeature.Shape = pPolyline
'             pFields = pfeature.Fields
'            'jaf--copy field values
'            For FieldCount = 0 To pFields.FieldCount - 1
'                pField = pFields.field(FieldCount)
'               'If pField.Editable Then
'               If Not pField.Type = esriFieldTypeOID And Not pField.Type = esriFieldTypeGeometry Then
'                 pNewFeature.value(FieldCount) = pfeature.value(FieldCount)
'               End If
'               'End If
'            Next FieldCount
'            pNewFeature.Store
'            'jaf--now delete the old features used to compose the union
'             pFCU = Nothing
'             pUFeat = Nothing
'             pFCU = m_edgeShp.Search(pQF, False)
'             pUFeat = pFCU.NextFeature
'            Do Until pUFeat Is Nothing
'               pUFeat.Delete
'                pUFeat = pFCU.NextFeature
'            Loop
'         End If   '(m_edgeShp.featurecount(pQF) > 1)
'      End If   '(pFeature.value(index) <> 0)
'       pfeature = pFC.NextFeature
'   Loop
'   'don't need to open junction intermediate file
'   ' pFact = New ShapefileWorkspaceFactory
'   ' pWorkspace = pFact.OpenFromFile(junctpath, 0)
'   ' pFeatWorkspace = pWorkspace
'
'   'stringlen = Len(junctname)
'   'position = InStrRev(junctname, ext, , vbTextCompare)
'   'filename = Left(junctname, (position - 1))
'   'extension = Right(junctname, (stringlen - position))
'
'   ' m_junctSHP = pFeatWorkspace.OpenFeatureClass(filename)
'   'MsgBox m_junctSHP.featurecount(Nothing)
'
'   'Open_NetFile 0 not using this approach now 9/9/2003
'
'   'join_Emme2toTransRef
'    pfeature = Nothing
'    pUFeat = Nothing
'    pNewFeature = Nothing
'
   edgeshpCnt = m_edgeShp.featurecount(Nothing)
   'open_CrystalReport
   'new way
   GlobalMod.Open_NetFile
'    Read_Netfile
    
   Exit Sub
eh:

        CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.Open_Shapefiles")
        MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.Open_Shapefiles")

   'MsgBox Err.Description, vbExclamation, "GlobalMod.Open_Shapefiles"
'    pfeature = Nothing
'    pUFeat = Nothing
'    pNewFeature = Nothing

End Sub

Public Sub Open_NetFile()
   'the newest routine to open the emme/2 outputs 11/25/03
   'each TransRefEdge can have up to 4 rows of output- IJ, JI, HOVIJ, and HOVJI
   'Scen_Link_ID is used to connect these results with intermediate file
   'this reads in the exported textfields to get the attributes contained in this file
   'and then to get the values and add to tblModelResultstemp
   'on error GoTo eh

   'jaf--let user know what's going on
        g_frmOutNetSelect.lblStatus.text = "GlobalMod.Open_Netfile: starting..."
   'pan frmOutNetSelect.Refresh
        WriteLogLine("called Open_NetFile")
        WriteLogLine("-------------------")

   Dim pStatusBar As IStatusBar
   Dim pProgbar As IStepProgressor
    pStatusBar = m_App.StatusBar
    pProgbar = pStatusBar.ProgressBar

   Dim thepath(5) As String
   thepath(0) = NetOutfilename
   thepath(1) = NetOutfilenameMidday
   thepath(2) = NetOutfilenamePM
   thepath(3) = NetOutfilenameEvening
   thepath(4) = NetOutfilenameNight

   Dim pRMTable As ITable
   Dim pWkSpFactory As IWorkspaceFactory
   Dim pFeatWorkspace As IFeatureWorkspace
   Dim pWkSp As IWorkspace
   Dim pWkSpDS As IDataset
   Dim pWkSpName As IWorkspaceName
   Dim pOutDSName As IDatasetName
        Dim pFields As IFields, pFieldsRM As IFields
   Dim pFields2 As IFields
   Dim pDataset As IDataset
   Dim pDSName As IDatasetName
   Dim pDS As IDataset
   Dim pED As IEnumDataset
   Dim pStTab2 As IStandaloneTable
   Dim pTable2 As ITable, pNewTable As ITable
   Dim pCursor2 As ICursor
   Dim pRowBuf As IRowBuffer
   Dim pRow As IRowBuffer, pRow2 As IRow  ' IRow
   Dim pName As iName
   Dim pStTab As IStandaloneTable
   Dim pStTabColl As IStandaloneTableCollection
   Dim pQF As IQueryFilter
   Dim pCursor As ICursor
    Dim pWorkspaceEdit As IWorkspaceEdit
    
   '++  up the output dataset name.
   '++ New .dbf file will be created in c:\temp
   ' pfields = pTab.Fields
'    pWkSpFactory = New ShapefileWorkspaceFactory
'    pWkSp = pWkSpFactory.OpenFromFile("c:\temp", 0)
   
    pWkSp = getWorkspaceOfWorkVersion 'get_Workspace
    pRMTable = get_TableClass(m_layers(17)) 'tblModelResults
    pFieldsRM = pRMTable.Fields
    pFields = pFieldsRM

    pFeatWorkspace = pWkSp
    pOutDSName = New tablename
   pOutDSName.Name = m_layers(18)  'tblModelResultstemp= use the original tblModelResults to be the bases for creating this file
    pOutDSName.WorkspaceName = pWkSpName

   '++ Check if .dbf table already exists: if yes, delete it.
        pED = pWkSp.Datasets(esriDatasetType.esriDTTable)
    pDS = pED.Next
        Debug.Print("start deleting table: " & pOutDSName.Name & " at " & Now())
        WriteLogLine("start deleting table: " & pOutDSName.Name & " at " & Now())
   
    pWorkspaceEdit = pWkSp 'pWorkspace
        If pWorkspaceEdit.IsBeingEdited = True Then pWorkspaceEdit.StopEditing(False)
   
   Do Until pDS Is Nothing
      If UCase(pDS.Name) = UCase(pOutDSName.Name) Then
         pDS.Delete
         'MsgBox "deleted tblmodel"
         Exit Do
      End If
       pDS = pED.Next
   Loop
        WriteLogLine("End deleting table " & pOutDSName.Name & ": " & pOutDSName.Name & " at " & Now())
        Debug.Print("End deleting table " & pOutDSName.Name & ": " & pOutDSName.Name & " at " & Now())
   
   '++ Create a new .dbf table
   'MsgBox pOutDSName.name
   'jaf--let user know what's going on
   
        g_frmOutNetSelect.lblStatus.text = "GlobalMod.Open_Netfile: creating new DBF table " & pOutDSName.Name
    'pan frmOutNetSelect.Refresh
        pFeatWorkspace.CreateTable(pOutDSName.Name, pFields, Nothing, Nothing, "")
    
    '++ Create a new stand alone table object to represent the .dbf table
     pStTab2 = getStandaloneTable(pOutDSName.Name)
    If pStTab2 Is Nothing Then
         pStTab2 = New StandaloneTable
     '   pFeatWorkspace.CreateTable pOutDSName.name, pFields, Nothing, Nothing, ""
        pStTab2.Table = pFeatWorkspace.openTable(pOutDSName.Name)
    End If
     pTable2 = pStTab2.Table
     pFields2 = pTable2.Fields
    'MsgBox "Created table 2"
    'the possible attribute fields in the crystal reports
    Dim MyChar As String
    Dim mymodel(15) As modelexport
    mymodel(0).textstring = "inode"
    mymodel(0).position = 999 'if position of attribut field not found, then value stay 999
    mymodel(1).textstring = "jnode"
    mymodel(1).position = 999
    mymodel(2).textstring = "@sov"
    mymodel(2).position = 999
    mymodel(3).textstring = "@hov2"
    mymodel(3).position = 999
    mymodel(4).textstring = "@hov3"
    mymodel(4).position = 999
    mymodel(5).textstring = "@vpool"
    mymodel(5).position = 999
    mymodel(6).textstring = "@light"
    mymodel(6).position = 999
    mymodel(7).textstring = "@mediu"
    mymodel(7).position = 999
    mymodel(8).textstring = "@heavy"
    mymodel(8).position = 999
    mymodel(9).textstring = "@voltr"
    mymodel(9).position = 999
     mymodel(10).textstring = "@brdtr"
    mymodel(10).position = 999
    mymodel(11).textstring = "@alttr"
    mymodel(11).position = 999
    mymodel(12).textstring = "@timtr"
    mymodel(12).position = 999

   mymodel(13).textstring = "@lid"
   mymodel(13).position = 999
   mymodel(14).textstring = "@ltype"
   mymodel(14).position = 999
   'get atrribute field postion because each row is proceeded by the attribute field names and then the values for that row

   Dim Myint As Integer
   Dim MyDbl As Double
   Dim templine(100) As String
   Dim TextLIne As String
   Dim compareString As String

   Dim runtype(5) As String
   runtype(0) = "AM"
   runtype(1) = "MD"
   runtype(2) = "PM"
   runtype(3) = "EV"
   runtype(4) = "NI"
   Dim tempString As String
   tempString = "Reading in Model Results for "

   Dim newName As String
   Dim length As Long, position As Long
   Dim i As Long, j As Long
   Dim pos As Long, firstpos As Long
   Dim fieldcnt As Long
   Dim pos2 As Long
   Dim nextString() As String
   Dim index As Long
   Dim rowcount As Long
   Dim file As Long, filecnt As Long
   Dim firstrun As Boolean
        Dim dctEdges As Dictionary(Of Object, Object)
   
   firstrun = False
   filecnt = 0
        dctEdges = New Dictionary(Of Object, Object)
    
    
        Debug.Print("Start editing to read punch files at: " & Now())
        WriteLogLine("Start editing to read punch files at: " & Now())

'  pWorkspaceEdit.StartEditing False
'  pWorkspaceEdit.StartEditOperation
  
   For file = 0 To 4
        If (thepath(file) <> "") Then
                Debug.Print("Start reading punch file [" & thepath(file) & "] at " & Now())
                WriteLogLine("Start reading punch file [" & thepath(file) & "] at " & Now())
            
                FileOpen(1, thepath(file), OpenMode.Input, OpenAccess.ReadWrite, OpenShare.Default)
                'Open thepath(file) For Input As 1   ' Open file.
         'jaf--let user know what's going on
                g_frmOutNetSelect.lblStatus.text = "GlobalMod.Open_Netfile: processing file " & thepath(file)
         'pan frmOutNetSelect.Refresh

         pProgbar.position = 0
         'tempString = tempString + runtype(file)
                pStatusBar.ShowProgressBar(tempString + runtype(file), 0, edgeshpCnt * 2, 1, True)

                InputString(1, TextLIne)   ' Read line into variable.
         'pan--let the user know what's going on
                g_frmOutNetSelect.lblStatus.text = "GlobalMod.Open_Netfile: processing line " & TextLIne
         'pan frmOutNetSelect.Refresh

         Trim (TextLIne)
         nextString = split(TextLIne, " ", , vbTextCompare)
         'nextString = split(TextLIne, vbTab, , vbTextCompare)

         firstpos = 999
         Dim colcount As Long
         colcount = 0

         'pan number of fields plus blank separators limited to 49
         For i = 0 To 49
            If Len(nextString(i)) > 2 Then
               templine(colcount) = nextString(i)
               'MsgBox nextString(i)
                        g_frmOutNetSelect.lblStatus.text = "GlobalMod.Open_Netfile: processing templine(colcount) " & templine(colcount)
               'pan frmOutNetSelect.Refresh
               colcount = colcount + 1
            End If
            If (nextString(i) = "result") Then
               Exit For
            Else

            End If
         Next i
         
         '[092007] hyu: The model output file may have columns in different order, _
                'it's safe to establish a correspondence between the column in the output _
                'file and the array index of the mymodel() variable.
            
         Dim col() As Integer   'variable for the above correspondence
         ReDim col(0 To colcount - 1)
         For j = 0 To colcount - 1
            col(j) = -1
            For i = 0 To 14
               compareString = mymodel(i).textstring '+ runtype(file)
               'MsgBox templine(j)
               If (templine(j) = compareString) Then
                  mymodel(i).position = j + 1
                  col(j) = i
                  'MsgBox compareString + " " + CStr(mymodel(i).position)
               End If
            Next i
         Next j

         fieldcnt = colcount
         rowcount = 0

        Dim FieldName As String, direction As String
        Dim value As Double
        Dim dct
                
        Do While Not EOF(1)   ' Loop until end of file.
            pos = 1
            
                    InputString(1, TextLIne)
            TextLIne = Trim(TextLIne)
            
            'the first column may be Emme2 action
            If LCase(Left(TextLIne, 1)) = "c" Then
            'skip this line -- comment line
            
            Else
                If Not IsNumeric(Left(TextLIne, 1)) Then
                'the first column is an Emme2 action, skip it.
                    TextLIne = Trim(Mid(TextLIne, 2, Len(TextLIne) - 1))
                End If
                
                j = 0   'j is the index of column
                pos = 0
                mymodel(col(0)).rowvalue = ""
                For i = 1 To Len(TextLIne)
                    If Mid(TextLIne, i, 1) <> " " Then
                        pos = pos + 1
                    Else
                        If pos <> 0 Then
                            If col(j) <> -1 Then    'when the column corresponding to a mymodel(?).
                                mymodel(col(j)).rowvalue = Mid(TextLIne, i - pos, pos)
                            End If
                            pos = 0
                            j = j + 1
                            'initialize the next mymodel() value to "".
                            If j < colcount And col(j) <> -1 Then mymodel(col(j)).rowvalue = ""
                        End If
                    End If
                Next i
                
'            End If
'            For j = 0 To 14
'                '[hyu 092106] judge the position value before call the input # function
'                If mymodel(j).position <> 999 Then
'                    Input #1, value
'                    mymodel(j).rowvalue = CStr(value)
'                End If
'
'            Next j
'            pos = pos + 1
'            Input #1, value   'get the result value

            '[hyu 100306]: exclude weave link (@ltype=0)
                If mymodel(14).rowvalue = "0" Then
                
    '            If (firstrun = False) Then
                Else
                    'still have to check if already created
                            If Not dctEdges.ContainsKey(CStr(mymodel(13).rowvalue)) Then
                                pRow = pTable2.CreateRowBuffer
                                '                     pRow2 = pTable2.CreateRow
                                '                     pRow = pRow2
                                dctEdges.Add(CStr(mymodel(13).rowvalue), pRow)

                            Else
                                pRow = dctEdges.Item(CStr(mymodel(13).rowvalue))
                            End If
                    
    '                 pQF = New QueryFilter
    '                pQF.WhereClause = "Scen_Link_ID = " + mymodel(13).rowvalue
    '                'MsgBox "Scen_Link_ID = " + mymodel(13).rowvalue
    '
    '                If (pTable2.rowcount(pQF) = 0) Then
    '                     pRow = pTable2.CreateRow
                        
                        index = pTable2.FindField("Scenario_ID")
                        If (index <> -1) Then
                           pRow.value(index) = m_ScenarioId
                        End If
                        index = pTable2.FindField("Scen_Link_ID")
                        If (index <> -1) Then
                           pRow.value(index) = Val(mymodel(13).rowvalue)
                        End If
                        If mymodel(14).rowvalue = "1" Or mymodel(14).rowvalue = "3" Or mymodel(14).rowvalue = "5" Then
                           direction = "IJ"
                        Else
                           direction = "JI"
                        End If
                        If (mymodel(2).position <> 999) Then
                           FieldName = direction + "VolSOV" + runtype(file)
                           index = pTable2.Fields.FindField(FieldName)
                           If (index <> -1) Then
                              pRow.value(index) = CDbl(mymodel(2).rowvalue)
                              'MsgBox fieldname
                              'jaf--let user know what's going on
                                    g_frmOutNetSelect.lblStatus.text = "GlobalMod.Open_Netfile: mymodel(2) processing field " & FieldName
                              'pan frmOutNetSelect.Refresh
                           End If
                        End If
                        If (mymodel(3).position <> 999) Then
                          FieldName = direction + "VolHOV2" + runtype(file)
                          index = pTable2.Fields.FindField(FieldName)
                          If (index <> -1) Then
                              pRow.value(index) = CDbl(mymodel(3).rowvalue)
                              'MsgBox prow.value(index)
                              'jaf--let user know what's going on
                                    g_frmOutNetSelect.lblStatus.text = "GlobalMod.Open_Netfile: mymodel(3) processing pRow.value(index) " & CStr(pRow.Value(index))
                              'pan frmOutNetSelect.Refresh
                          End If
                        End If
                        If (mymodel(4).position <> 999) Then
                          FieldName = direction + "VolHOV3" + runtype(file)
                          index = pTable2.Fields.FindField(FieldName)
                          If (index <> -1) Then
                             pRow.value(index) = CDbl(mymodel(4).rowvalue)
                          End If
                        End If
                        If (mymodel(5).position <> 999) Then
                          FieldName = direction + "VolVanpool" + runtype(file)
                          index = pTable2.Fields.FindField(FieldName)
                          If (index <> -1) Then
                             pRow.value(index) = CDbl(mymodel(5).rowvalue)
                          End If
                        End If
                        If (mymodel(6).position <> 999) Then
                           FieldName = direction + "VolLtTruck" + runtype(file)
                          index = pTable2.Fields.FindField(FieldName)
                          If (index <> -1) Then
                             pRow.value(index) = CDbl(mymodel(6).rowvalue)
                          End If
    
                        End If
                        If (mymodel(7).position <> 999) Then
                          FieldName = direction + "VolMedTruck" + runtype(file)
                          index = pTable2.Fields.FindField(FieldName)
                          If (index <> -1) Then
                             pRow.value(index) = CDbl(mymodel(7).rowvalue)
                          End If
                        End If
                        If (mymodel(8).position <> 999) Then
                          FieldName = direction + "VolHvyTruck" + runtype(file)
                          index = pTable2.Fields.FindField(FieldName)
                          If (index <> -1) Then
                             pRow.value(index) = CDbl(mymodel(8).rowvalue)
                          End If
                        End If
                        If (mymodel(9).position <> 999) Then
                          FieldName = direction + "VolTransit" + runtype(file)
                          index = pTable2.Fields.FindField(FieldName)
                          If (index <> -1) Then
                             pRow.value(index) = CDbl(mymodel(9).rowvalue)
                          End If
                        End If
                        If (mymodel(10).position <> 999) Then
                          FieldName = direction + "AlgtTransit" + runtype(file)
                          index = pTable2.Fields.FindField(FieldName)
                          If (index <> -1) Then
                             pRow.value(index) = CDbl(mymodel(10).rowvalue)
                          End If
                        End If
                        If (mymodel(11).position <> 999) Then
                          FieldName = direction + "BrdTransit" + runtype(file)
                          index = pTable2.Fields.FindField(FieldName)
                          If (index <> -1) Then
                             pRow.value(index) = CDbl(mymodel(11).rowvalue)
                          End If
                        End If
                        If (mymodel(12).position <> 999) Then
                          FieldName = direction + "TravelTimeTransit" + runtype(file)
                          index = pTable2.Fields.FindField(FieldName)
                          If (index <> -1) Then
                             pRow.value(index) = CDbl(mymodel(12).rowvalue)
                          End If
                        End If
                        
                        'jaf--let user know what's going on
                            g_frmOutNetSelect.lblStatus.text = "GlobalMod.Open_Netfile: done with firstrun = false"
                        'pan frmOutNetSelect.Refresh
                        
                        rowcount = rowcount + 1
                    End If

                End If   '(firstrun = False)
                pStatusBar.StepProgressBar
            Loop  'While Not EOF(1)

            pStatusBar.HideProgressBar
                FileClose(1)   ' FileClose file.

            'jaf--let user know what's going on
                g_frmOutNetSelect.lblStatus.text = "GlobalMod.Open_Netfile: done with file " & thepath(file)
            'pan frmOutNetSelect.Refresh

            If (filecnt = 0) Then
                firstrun = True
                filecnt = 1
            End If
        End If   '(thepath(file) <> "")
            Debug.Print("End reading punch file [" & thepath(file) & "] at " & Now())
            WriteLogLine("End reading punch file [" & thepath(file) & "] at " & Now())
    
    Next file
    pProgbar.position = 0
        pStatusBar.ShowProgressBar("Inserting rows...", 0, dctEdges.Count, 1, True)
    
        Debug.Print("Inserting  rows to " & pStTab2.Name & " starts at: " & Now())
        WriteLogLine("Inserting rows to " & pStTab2.Name & " starts at: " & Now())
    
    Dim pIns As ICursor
     pIns = pTable2.Insert(True)
    For i = 0 To dctEdges.count - 1
            pRow = dctEdges.Item(i)
            pIns.InsertRow(pRow)
        pStatusBar.StepProgressBar
    Next i
    pStatusBar.HideProgressBar
        Debug.Print("End Inserting rows at: " & Now())
        WriteLogLine("End Inserting rows at: " & Now())
    
        dctEdges.Clear()
    dctEdges = Nothing
   
'Debug.Print "Start to save reading punch files at: " & Now()
'pWorkspaceEdit.StopEditOperation
'pWorkspaceEdit.StopEditing True
'Debug.Print "End saving at: " & Now()

'   '++ Add the new sorted table to map document
'    pStTab = New StandaloneTable
'    pStTab.Table = pTable2 ' pNewTable
'    pStTabColl = m_Map
'   pStTabColl.AddStandaloneTable pStTab

   m_Doc.UpdateContents
        Debug.Print("Start join_Emme2toScenarioEdge at " & Now())
   GlobalMod.join_Emme2toTransRef
        Debug.Print("End join_Emme2toScenarioEdge at " & Now())
        CloseLogFile("Retireval succeeded!")
   Exit Sub
eh:

        CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.open_NetFile")
        MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.open_NetFile")

  'MsgBox Err.Description + " open_NetFile"
End Sub



Public Sub join_tblModelResultstoEmme2(pFCNew As IFeatureClass)
   'MsgBox "joining tables"
   'called by join_EMME2toTransRef to join the tblModel to spatial joined Emme/2 shapefiles to TransRefedges (pFCNew)

   'will have a default table to keep using over and over for this routine
   'on error GoTo eh

   'jaf--let user know what's going on
        g_frmOutNetSelect.lblStatus.text = "GlobalMod.join_tblModelResultstoEmme2: commencing join of tables"
   'pan frmOutNetSelect.Refresh
        WriteLogLine("called join_tblModelResultstoEmme2")
        WriteLogLine("----------------------------------")

   joinname = "ModelResults_TransRef" 'InputBox("Enter the output shapefile name:", "Join by ModelResults to TransRefEdges", "ModelResult_TR")

   Dim pMRTable As ITable
   Dim pWS As IWorkspace
    pWS = get_Workspace
        Check_if_exist(pWS, joinname)

    pMRTable = get_TableClass(m_layers(18)) 'tblModelResultstemp
   Dim pMemRelClassFact As IMemoryRelationshipClassFactory
    pMemRelClassFact = New MemoryRelationshipClassFactory
   Dim pRelClass As IRelationshipClass
   'a table join is actually a relations ship class so need to create one
        pRelClass = pMemRelClassFact.Open(joinname, pFCNew, "Scen_Link", pMRTable, "Scen_Link_ID", "forward", "backward", esriRelCardinality.esriRelCardinalityOneToMany)

   ' ++ Perform the join
   Dim pRelQueryTableFact As IRelQueryTableFactory
   Dim pRelQueryTab As ITable
    pRelQueryTableFact = New RelQueryTableFactory
   'get the relationship class table
    pRelQueryTab = pRelQueryTableFact.Open(pRelClass, True, Nothing, Nothing, "", True, True)

    m_FinalJoin = pRelQueryTab 'since the primary table object was a feature class have the 'Shape' field to make a New feture class
   ' ++ Print the fields
   Dim pCursor As ICursor
    pCursor = pRelQueryTab.Search(Nothing, True)
   Dim pField As IField
   Dim pFields As IFields
   Dim intI As Long, intJ As Long

    pFields = pCursor.Fields
   intI = pFields.FieldCount - 1
   For intJ = 0 To intI
      pField = pFields.field(intJ)
     'MsgBox pField.name
            g_frmRender.ComboBox1.Items.Add(pField.Name)
   Next intJ
   Dim pRow As IRow
   Dim pTC As ICursor
   ' pTC = pRelQueryTab.Search(Nothing, False)
   'add final object created from the joins to the map display so can use the renderer
   Dim i As Long
   Dim pFeatLayer As IFeatureLayer
    pFeatLayer = New FeatureLayer
    pFeatLayer.FeatureClass = m_FinalJoin
   'MsgBox m_FinalJoin.AliasName
   pFeatLayer.Name = joinname

        m_Map.AddLayer(pFeatLayer)

   'jaf--let user know what's going on
        g_frmOutNetSelect.lblStatus.text = "GlobalMod.join_tblModelResultstoEmme2: finished join of tables"
   'pan frmOutNetSelect.Refresh

        g_frmLocateEdg_Jun.Close()
        g_frmRender.Show()
   Exit Sub
eh:

        CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.join_tblModelResultstoEmme2")
        MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.join_tblModelResultstoEmme2")

  'MsgBox Err.Description, vbExclamation, "join_tblModelResultstoEmme2"
End Sub


Public Sub join_Emme2toTransRef()
   'MsgBox "joining"
   'called by Open_Export
   'jaf--actually, seems to be called by OpenNetfile
   'this spatially join the intermediate shapefiles to the TransRefEdges
   'on error GoTo eh

   'jaf--let user know what's going on
        g_frmOutNetSelect.lblStatus.text = "GlobalMod.join_Emme2toTransRef: commencing join of int. shapefile to TransRefEdges"
   'pan frmOutNetSelect.Refresh
        WriteLogLine("called join_Emme2toTransRef")
        WriteLogLine("----------------------------------")

   ' Get the output workspace name - make it
   ' the same as the input polygon layer in the join
   Dim pTRedgeLayer As IFeatureLayer
        pTRedgeLayer = get_MasterNetwork(0, m_Map)

    Dim pDataset As IDataset
   Dim pWkSpDataset As IWorkspace
    pWkSpDataset = get_Workspace
   Dim pWkSpName As IWorkspaceName
    pDataset = pWkSpDataset 'pTRedgeLayer.FeatureClass

   ' pWkSpDataset = pTRedgeLayer.FeatureClass.FeatureDataset.Workspace
    pWkSpName = pDataset.FullName

   ' create the name object for the output join by location shapefile
   Dim strOutName As String
   Dim pFCName As IFeatureClassName
   Dim pOutDSName As IDatasetName
   Dim pName As iName
   Dim pDS As IDataset
   Dim pED As IEnumDataset

   'will create a default feture class to use over and over for  this routine instead of the code below
   strOutName = "Emme2_TransRef" 'InputBox("Enter the output shapefile name:", "Join by EMME/2-edges to TransRef", "Emme2_TR")
        Check_if_exist(pWkSpDataset, strOutName)
    pFCName = New FeatureClassName
   With pFCName
            .FeatureType = esriFeatureType.esriFTSimple
     .ShapeFieldName = "Shape"
            .ShapeType = esriGeometryType.esriGeometryPolyline
   End With

    pOutDSName = pFCName
   With pOutDSName
     .Name = strOutName
      .WorkspaceName = pWkSpName
   End With

    ' If (MsgBox("The shapefile already exists, try another name?", vbOKCancel) = vbOK) Then
       'strOutName = InputBox("Enter the output shapefile name:", "Join by EMME/2-edges to TransRef", "Emme2_TR")
    pName = pOutDSName

   ' Do a join by location that joins the attributes of the
   ' first point contained within each polygon.
   Dim pSpJoin As ISpatialJoin
   Dim pFCNew As IFeatureClass
    pSpJoin = New SpatialJoin

   'With pSpJoin
   '   .JoinTable = m_edgeShp
  '    .SourceTable = pTRedgeLayer.FeatureClass
  '   .LeftOuterJoin = True
   'End With

   ' setting maxMapDist to 0 means that only points within
   ' each each polygon will be considered
   ' pFCNew = pSpJoin.JoinNearest(pName, 0)
   ' pFCNew = pSpJoin.JoinWithin(pName)
   'now just getting m_edgeshp since it is an overlay deal
    pFCNew = m_edgeShp

   ' Create a new layer and add it to the Map
   If Not pFCNew Is Nothing Then
     Dim pNewFLayer As IFeatureLayer
      pNewFLayer = New FeatureLayer
      pNewFLayer.FeatureClass = pFCNew
     'pNewFLayer.name = "Sample Join by Location"
     'm_Map.AddLayer pNewFLayer

      m_Doc.UpdateContents
   End If
    pSpJoin = Nothing
'   frmLocateEdg_Jun.Hide
   
        Debug.Print("Start join_tblModelResultstoEmme2 at " & Now())
        GlobalMod.join_tblModelResultstoEmme2(pFCNew)
        Debug.Print("End join_tblModelResultstoEmme2 at " & Now())

   'jaf--let user know what's going on
        g_frmOutNetSelect.lblStatus.text = "GlobalMod.join_Emme2toTransRef: finished join of int. shapefile to TransRefEdges"
   'pan frmOutNetSelect.Refresh

   Exit Sub

eh:

        CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.join_Emme2toTransRef")
        MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.join_Emme2toTransRef")

   'MsgBox "JoinByLocation_Nearest: " & Err.Number & " " & Err.Description


End Sub

    Public Function Check_if_exist(ByVal pWkspace As IWorkspace, ByVal Name As String) As Boolean
        'on error GoTo eh
        '[031506] hyu: change the code to improve performance.

        Dim pEnumDSNm As IEnumDatasetName
        Dim pDSNm As IDatasetName
        Dim pName As IName
        Dim pDS As IDataset
        Dim bExist As Boolean
        bExist = False

        pEnumDSNm = pWkspace.DatasetNames(esriDatasetType.esriDTFeatureClass)
        pDSNm = pEnumDSNm.Next
        Do Until pDSNm Is Nothing
            If UCase(pDSNm.Name) = UCase(Name) Then

                bExist = True
                '             pName = pDSNm
                '             pDS = pName.Open
                '            pDS.Delete
                Exit Do
            End If
            pDSNm = pEnumDSNm.Next
        Loop

        Check_if_exist = False
        Dim pTbl As IFeatureClass
        If bExist Then
            pName = pDSNm
            pTbl = pName.Open

            deleteAllRows(pWkspace, pTbl)
            Check_if_exist = True
        End If



        pName = Nothing
        pDSNm = Nothing
        pEnumDSNm = Nothing
        Exit Function

        '  '++ Check to see if shapefile already exists
        '  Dim nm2 As String
        '  Dim pED As IEnumDataset
        '  Dim pDS As IDataset
        '   pED = pWkspace.Datasets(esriDTAny)
        '   pDS = pED.Next
        '
        '  'Get the first dataset for the wkspace
        '  'Check_for_shapefile = False
        '  Do Until pDS Is Nothing
        ' 'MsgBox pDS.name
        '    If pDS.name = name Then
        '      pDS.Delete
        '      WriteLogLine "GlobalMod.Check_if_exist: deleted" & name
        '      Exit Do
        '    End If
        '     pDS = pED.Next
        '  Loop
        '
        '  Exit Sub
        '
eh:
        CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.Check_if_exist")
        MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.Check_if_exist")

    End Function


Public Sub loadScenarios()
'called by frmNetOutSelect for loading the scenario ids

'on error GoTo eh
'pan cmbofiletype problematic now use CommonDialog1
    'frmOutNetSelect.cmbofiletype.AddItem "Emme/2 Files"
    'frmOutNetSelect.CommonDialog1.filetype.AddItem "Emme/2 Files"
       'frmOutNetSelect.cmbofiletype.AddItem "All Files"
       'frmOutNetSelect.cmbofiletype.ListIndex = 0

        'SEC 052410_commented out the following
        'g_frmOutNetSelect.lblFile.Caption = "Name of network file: "
        'g_frmOutNetSelect.CommonDialog1.Filter = "*.rpt; *.Out; *.txt"

        'frmOutNetSelect.Height = 6990
       'function call here to load scenario
        'g_frmOutNetSelect.filetype = 0

    Dim pScenTable As ITable
    Dim pRow As IRow, pNewRow As IRow
    Dim pTC As ICursor

    Dim pWS As IWorkspace
     pWS = get_Workspace
     pScenTable = get_TableClass(m_layers(16)) 'tblModelScenario

     pTC = pScenTable.Search(Nothing, False)
     pRow = pTC.NextRow
    Dim index As Long
    Dim value As Long

    index = pRow.Fields.FindField("Title")
    'frmOutNetSelect.txttitle = pRow.value(index)
    index = pRow.Fields.FindField("Description")
    'frmOutNetSelect.txtdes = pRow.value(index)
    Do Until pRow Is Nothing 'not using this method of getting the scenario id 9/23/2003
        index = pRow.Fields.FindField("Scenario_ID")
        value = pRow.value(index)
        'frmOutNetSelect.Combo1.AddItem CStr(value)
         pRow = pTC.NextRow
   Loop
   'frmOutNetSelect.Combo1.ListIndex = 0
        g_frmOutNetSelect.Show()


   Exit Sub

eh:
        CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.LoadScenarios")
        MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.LoadScenarios")
  'MsgBox "LoadScenario: " & Err.Number & " " & Err.Description
End Sub

Public Sub add_DataTotblModelResults()
'not used 9/9/2003- see open_Export
  'on error GoTo eh
        WriteLogLine("called add_DataTotblModelResults")
        WriteLogLine("--------------------------------")

    Dim pModelResults As ITable
    Dim pRow As IRow, pNewRow As IRow
    Dim pTC As ICursor

    Dim pWS As IWorkspace
     pWS = get_Workspace
     pModelResults = get_TableClass("tblModelResults2")

    Dim pFC As IFeatureCursor
    Dim pFeature As IFeature
     pFC = m_edgeShp.Search(Nothing, False)
     pFeature = pFC.NextFeature

    Dim index As Long
    index = m_edgeShp.Fields.FindField("ScenarioID")

    Dim sID As Long
    sID = pFeature.value(index)

    Dim e2Index As Long, e2Indexf As Long
    Dim count As Long
    index = pModelResults.FindField("Scenario_ID")
    e2Index = pModelResults.FindField("PSRC_E2ID")
    count = 0
    e2Indexf = m_edgeShp.Fields.FindField("PSRC_E2ID")

    Do Until pFeature Is Nothing

         pNewRow = pModelResults.CreateRow
        pNewRow.value(index) = sID
        pNewRow.value(e2Index) = pFeature.value(e2Indexf)
        pNewRow.Store
         pFeature = pFC.NextFeature
    Loop

    Exit Sub

eh:
        CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.add_DataTotblModelResults")
        MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.add_DataTotblModelResults")
End Sub









Public Sub RemoveEvents()
'can't create project routes if have events on them!
'on error GoTo eh
        WriteLogLine("called RemoveEvents")
        WriteLogLine("-------------------")

   Dim pFeatLayer As IFeatureLayer
   Dim i As Long
   Dim pLayer As ILayer
   Dim pFLayer As IFeatureLayer
   Dim pFeatClass As IFeatureClass

   Dim count As Long
   count = 0
   m_Map.ClearSelection
   For i = 0 To m_Map.LayerCount - 1 'first loop to see how many event route source have
    If TypeOf m_Map.Layer(i) Is IFeatureLayer Then
       pFLayer = m_Map.Layer(i)
       pFeatClass = pFLayer.FeatureClass

      If Not pFeatClass Is Nothing Then
                    If TypeOf pFeatClass Is ESRI.ArcGIS.Location.IRouteEventSource Then

                        count = count + 1
                    End If
      End If
    End If

  Next i
  Dim j As Long
  If (count > 0) Then
   For j = 0 To count - 1 'now go get them and remove them
    For i = 0 To m_Map.LayerCount - 1
      If TypeOf m_Map.Layer(i) Is IFeatureLayer Then
         pFLayer = m_Map.Layer(i)
         pFeatClass = pFLayer.FeatureClass

        If Not pFeatClass Is Nothing Then
                            If TypeOf pFeatClass Is ESRI.ArcGIS.Location.IRouteEventSource Then

                                m_Map.DeleteLayer(pFLayer) 'just removes them
                                Exit For
                            End If
        End If
      End If

    Next i

   Next j
  End If

  m_Doc.ActiveView.Refresh
  m_Map.ClearSelection
   Exit Sub
eh:
        CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.RemoveEvents")
        MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.RemoveEvents")
  'MsgBox Err.Description, vbInformation, "ArchivingNetwork"
   pFLayer = Nothing
End Sub





    Public Function WriteTollInfoOLD(ByVal objFile As Object, ByVal pRow As IRow, ByVal strLineI As String, ByVal strLineJ As String, ByVal strTOD As String, ByVal intOneway As Integer)
        '       0 = one way from I
        '       1 = one way from J
        '       2 = two way
        Dim strLine2 As String
        With pRow

            Select Case intOneway
                Case 0

                    strLine2 = strLineI & .value(.Fields.FindField("IJTollSOV" & strTOD)) & " " _
                        & .value(.Fields.FindField("IJTollHOV2" & strTOD)) & " " _
                        & .value(.Fields.FindField("IJTollHOV3" & strTOD)) & " " _
                        & .value(.Fields.FindField("IJTollTrkLt" & strTOD)) & " " _
                        & .value(.Fields.FindField("IJTollTrkMed" & strTOD)) & " " _
                        & .value(.Fields.FindField("IJTollTrkHvy" & strTOD))
                    objFile.printline(strLine2)
                Case 1

                    strLine2 = strLineJ & .value(.Fields.FindField("JITolSOV" & strTOD)) & " " _
                        & .value(.Fields.FindField("JITollHOV2" & strTOD)) & " " _
                        & .value(.Fields.FindField("JITollHOV3" & strTOD)) & " " _
                        & .value(.Fields.FindField("JITollTrkLt" & strTOD)) & " " _
                        & .value(.Fields.FindField("JITollTrkMed" & strTOD)) & " " _
                        & .value(.Fields.FindField("JITolTrkHvy" & strTOD))
                    objFile.printline(strLine2)
                Case 2
                    strLine2 = strLineI & .value(.Fields.FindField("IJTollSOV" & strTOD)) & " " _
                       & .value(.Fields.FindField("IJTollHOV2" & strTOD)) & " " _
                       & .value(.Fields.FindField("IJTollHOV3" & strTOD)) & " " _
                       & .value(.Fields.FindField("IJTollTrkLt" & strTOD)) & " " _
                       & .value(.Fields.FindField("IJTollTrkMed" & strTOD)) & " " _
                       & .value(.Fields.FindField("IJTollTrkHvy" & strTOD))
                    objFile.printline(strLine2)

                    strLine2 = strLineJ & .value(.Fields.FindField("JITollSOV" & strTOD)) & " " _
                       & .value(.Fields.FindField("JITollHOV2" & strTOD)) & " " _
                       & .value(.Fields.FindField("JITollHOV3" & strTOD)) & " " _
                       & .value(.Fields.FindField("JITollTrkLt" & strTOD)) & " " _
                       & .value(.Fields.FindField("JITollTrkMed" & strTOD)) & " " _
                       & .value(.Fields.FindField("JITollTrkHvy" & strTOD))
                    objFile.printline(strLine2)
            End Select
        End With
    End Function

    Public Sub createParkRideFile(pathName As String, fileName As String)
        'on error GoTo eh
        WriteLogLine("")
        WriteLogLine("=======================================")
        WriteLogLine("createParkRideFile started " & Now())
        WriteLogLine("=======================================")
        WriteLogLine("")

        Dim strOutFile As String
        strOutFile = pathName + "\" + fileName

        'get Junction layer
        'Dim pFL As IFeatureLayer
        Dim pFClsJct As IFeatureClass
        pFClsJct = m_junctShp
        Dim i As Long

        'Get tblModelScenario table
        Dim pWS As IWorkspace
        pWS = get_Workspace()

        Dim pTbl As ITable
        pTbl = get_TableClass(m_layers(16)) 'tblModelScenario

        'get field indices
        Dim fldSceID1 As Long, fldNodeID As Long, fldCap As Long
        Dim fldSceID2 As Long, fldTitle As Long

        With pFClsJct
            fldSceID1 = .FindField("ScenarioID")
            fldNodeID = .FindField("Scen_Node") '("Emme2nodeID")
            fldCap = .FindField("P_RStalls")
        End With

        With pTbl
            fldSceID2 = .FindField("Scenario_ID")
            fldTitle = .FindField("Title")
        End With

        WriteLogLine("Done with field setup, creating P&R Buildfile")

        'Create output file
        Dim fs As Object, F As Object
        fs = CreateObject("Scripting.FileSystemObject")

        'Write header
        Dim strLine As String
        F = fs.CreateTextFile(strOutFile)
        F.WriteLine("t matrices")
        F.WriteLine("m matrix=md" & """" & "pnrcap" & """")

        WriteLogLine("Created P&R Buildfile and wrote header, starting selection")

        'Select Park and Ride
        Dim psort As ITableSort, pFilt As IQueryFilter
        Dim pCs_Jct As ICursor, pRow_Jct As IRow
        Dim pCs_Sce As ICursor, pRow_Sce As IRow
        Dim strSceID_pre As String, strSceID As String

        pFilt = New QueryFilter
        psort = New TableSort
        pFilt.WhereClause = "JunctionType=7"
        With psort
            .Fields = "Scen_Node" ' "ScenarioID"
            .Ascending("Scen_Node") = True '("ScenarioID") = False
            .QueryFilter = pFilt
            .Table = pFClsJct
            .Sort(Nothing)
            pCs_Jct = .Rows
        End With

        WriteLogLine("Finished selection & sort, starting iterations")
        ' pCs = pFClsJct.Search(pFilt, False)
        pRow_Jct = pCs_Jct.NextRow
        strSceID_pre = ""
        'MsgBox "Before Do createParkRideFile"
        Do Until pRow_Jct Is Nothing
            'MsgBox "In Do until pRow_Jct Is Nothing createParkRideFile"
            'MsgBox "pRow_Jct.value(fldsceID1)=" & pRow_Jct.value(fldSceID1)
            'pan problem if strSceID is null if SDE
            If Not IsDBNull(pRow_Jct.Value(fldSceID1)) Then
                strSceID = pRow_Jct.Value(fldSceID1)
                If fVerboseLog Then WriteLogLine("row ScenID=" & strSceID)
                'when the cursor gets to a different scenarioID.
                If strSceID_pre <> strSceID Then
                    pFilt.WhereClause = "Scenario_ID=" & strSceID
                    pCs_Sce = pTbl.Search(pFilt, False)
                    pRow_Sce = pCs_Sce.NextRow
                    strLine = "c ScenarioID=" & strSceID

                    'when the scenario is a new scenario
                    If pRow_Sce Is Nothing Then
                        strLine = strLine & "'Scenario_1' P&R capacities"
                        F.WriteLine(strLine)
                    Else 'not a new scenario
                        strLine = strLine & "'" & pRow_Sce.Value(fldTitle) & "' P&R capacities"
                    End If
                End If
                strSceID_pre = strSceID
            End If

            strLine = " all " & pRow_Jct.Value(fldNodeID) & ":" & pRow_Jct.Value(fldCap)
            F.WriteLine(strLine)

            pRow_Jct = pCs_Jct.NextRow
        Loop

        F.Close()
        pFilt = Nothing
        pRow_Sce = Nothing
        pRow_Jct = Nothing
        pCs_Sce = Nothing
        pCs_Jct = Nothing
        pFClsJct = Nothing
        pTbl = Nothing

        psort = Nothing
        WriteLogLine("FINISHED createParkRideFile at " & Now())
        Exit Sub
eh:
        CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.createParkRideFile")
        MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.createParkRideFile")
        'MsgBox Err.Description, vbExclamation, "createParkRideFile"
    End Sub

    Public Sub createTollsFile(ByVal pathName As String, ByVal filename As String, ByVal dctReservedNodes As Dictionary(Of Object, Object), ByVal app As IApplication, ByVal projRteFL As IFeatureLayer)
        'creates Toll buildfile(s) for Emme2
        'on error GoTo eh
        Try

            Dim strEdge As String, strTblModeTolls As String, strTblLineTolls As String
            Dim strOutFile1 As String, strOutFile2 As String, strOutFile3 As String
            Dim strOutFile4 As String, strOutFile5 As String

            WriteLogLine("")
            WriteLogLine("======================================")
            WriteLogLine("createTollsFile started " & Now())
            WriteLogLine("======================================")
            WriteLogLine("")

            strEdge = "test3edge"
            strTblModeTolls = m_layers(9) 'ModeTolls this is a attribute relationship with TransRefEdges
            ' strTblLineTolls = m_layers(11) 'tblLineTolls Events this is a spatial relationship
            strOutFile1 = pathName + "\" + filename + "_TollAM.txt"
            strOutFile2 = pathName + "\" + filename + "_TollMD.txt"
            strOutFile3 = pathName + "\" + filename + "_TollPM.txt"
            strOutFile4 = pathName + "\" + filename + "_TollEV.txt"
            strOutFile5 = pathName + "\" + filename + "_TollNT.txt"

            'get Edge layer
            'Dim pFL As IFeatureLayer

            Dim pFClsEdge As IFeatureClass
            pFClsEdge = m_edgeShp

            'Get tblModelScenario table
            'Dim pTblMode As ITable ' pTblLine As ITable- 4/20/05 using this as an event
            Dim evtMode As IFeatureClass
            Dim pGeomEvt As IGeometry
            Dim pWS As IWorkspace
            Dim pFWS As IFeatureWorkspace
            GC.Collect()

            pWS = get_Workspace()
            'pFWS = pWS

            'pTblMode = g_FWS.OpenTable(strTblModeTolls)

            'pTblMode = get_TableClass(strTblModeTolls)
            ' pTblLine = get_TableClass(strTblLineTolls, pws)

            'stefan**************************8
            'evtMode = get_RouteEventSource(strTblLineTolls, app)

            'get field indices
            Dim fldEdgeID1 As Long, fldINode As Long, fldJNode As Long, fldOneWay As Long
            Dim fldUpdate As Long
            Dim fldEdgeID2 As Long
            Dim fldPrjRte1 As Long, fldPrjRte2 As Long



            WriteLogLine("getting field indices to edge shpFile")
            With pFClsEdge
                fldEdgeID1 = .FindField("PSRCEdgeID")
                fldOneWay = .FindField("Oneway")
                fldPrjRte1 = .FindField("prjRte")
                fldUpdate = .FindField("Updated1")
                fldINode = .FindField("INode")
                fldJNode = .FindField("JNode")
            End With

            WriteLogLine("getting field index to tblModeTolls")
            With g_TblMode
                fldEdgeID2 = .FindField("PSRCEdgeID")
            End With

            WriteLogLine("getting field index to tblLineTolls")

            'stefan*****************************
            'With evtMode
            ''fldPrjRte2 = .FindField("projRteID")
            'End With

            WriteLogLine("Done getting layers, now creating output files")

            'Create output file
            Dim fs As Object, f1 As Object, f2 As Object, f3 As Object, f4 As Object, f5 As Object
            fs = CreateObject("Scripting.FileSystemObject")

            'Write header
            Dim strLineI As String, strLineJ As String
            f1 = fs.CreateTextFile(strOutFile1)
            f2 = fs.CreateTextFile(strOutFile2)
            f3 = fs.CreateTextFile(strOutFile3)
            f4 = fs.CreateTextFile(strOutFile4)
            f5 = fs.CreateTextFile(strOutFile5)

            f1.WriteLine("headers")
            f2.WriteLine("headers")
            f3.WriteLine("headers")
            f4.WriteLine("headers")
            f5.WriteLine("headers")

            WriteLogLine("Done creating output files and writing headers, starting modeTolls scan")

            Dim psort As ITableSort, pFilt As IQueryFilter
            Dim pFCs_Edge As IFeatureCursor, pFtEdge As IFeature
            Dim pCs_Mode As ICursor, pRow_Mode As IRow
            Dim pCs_Line As IFeatureCursor, pRow_Line As IRow
            Dim strEdgeID_pre As String, strEdgeID As String
            Dim evtFeat As IFeature
            'System.Runtime.InteropServices.Marshal.ReleaseComObject(pFCs_Edge)
            'get scenarioid
            pFCs_Edge = pFClsEdge.Search(Nothing, False)
            pFtEdge = pFCs_Edge.NextFeature
            m_ScenarioId = pFtEdge.Value(pFtEdge.Fields.FindField("ScenarioID"))
            pFCs_Edge = Nothing
            pFtEdge = Nothing

            'Future new toll projects
            Dim intOneway As Integer, lEdgeID As Long
            Dim pTblPrjInSce As ITable
            Dim pCs_Prj As ICursor, pRow_Prj As IRow
            Dim pFClsPrjRte As IFeatureClass
            Dim pFCsRte As IFeatureCursor, pFtPrj As IFeature
            Dim pFilt2 As ISpatialFilter
            Dim lPrjRteID As Long
            Dim dir As Integer, dirPrj As Integer
            Dim dct As Dictionary(Of Object, Object)
            Dim iNode As Long, JNode As Long, iUseEmmeN As Integer
            'Dim pSort As ITableSort

            psort = New TableSort
            pFilt = New QueryFilter
            pFilt2 = New SpatialFilter
            dct = New Dictionary(Of Object, Object)


            pFClsPrjRte = projRteFL.FeatureClass
            pTblPrjInSce = get_TableClass(m_layers(15))

            pFilt.WhereClause = "Scenario_ID =" & m_ScenarioId
            With psort
                .Fields = g_InSvcDate
                .Ascending(g_InSvcDate) = True
                .QueryFilter = pFilt
                .Table = pTblPrjInSce
                .Sort(Nothing)
                pCs_Prj = .Rows
            End With
            pCs_Prj = pTblPrjInSce.Search(pFilt, False)
            pRow_Prj = pCs_Prj.NextRow
            Dim pRelOp As IRelationalOperator

            Do Until pRow_Prj Is Nothing
                lPrjRteID = pRow_Prj.Value(pRow_Prj.Fields.FindField("projRteID"))
                pFilt.WhereClause = "projRteID=" & lPrjRteID
                pCs_Line = evtMode.Search(pFilt, False)
                evtFeat = pCs_Line.NextFeature

                pFCsRte = pFClsPrjRte.Search(pFilt, False)
                pFtPrj = pFCsRte.NextFeature
                If Not evtFeat Is Nothing Then

                    With pFilt2
                        .Geometry = evtFeat.ShapeCopy
                        .GeometryField = evtMode.ShapeFieldName
                        .SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects
                    End With

                    pFCs_Edge = pFClsEdge.Search(pFilt2, False)
                    pFtEdge = pFCs_Edge.NextFeature
                    Do Until pFtEdge Is Nothing
                        pRelOp = pFtEdge.Shape
                        lEdgeID = pFtEdge.Value(fldEdgeID1)

                        If Not (pRelOp.Touches(evtFeat.Shape) Or pRelOp.Crosses(evtFeat.Shape)) Then
                            iNode = pFtEdge.Value(fldINode)
                            JNode = pFtEdge.Value(fldJNode)
                            iUseEmmeN = pFtEdge.Value(pFtEdge.Fields.FindField("UseEmmeN"))

                            Select Case iUseEmmeN
                                Case 0
                                    iNode = iNode + m_Offset
                                    JNode = JNode + m_Offset
                                Case 1
                                    iNode = dctReservedNodes.Item(iNode)
                                    JNode = JNode + m_Offset
                                Case 2
                                    JNode = dctReservedNodes.Item(JNode)
                                    iNode = iNode + m_Offset
                                Case 3
                                    iNode = dctReservedNodes.Item(iNode)
                                    JNode = dctReservedNodes.Item(JNode)
                            End Select

                            strLineI = iNode & " " & JNode & " "
                            strLineJ = JNode & " " & iNode & " "
                            intOneway = pFtEdge.Value(fldOneWay)


                            dirPrj = 0
                            dir = projectAndEdgeDirectionRelationship(pFtPrj, pFtEdge, dirPrj)
                            'Use the lineTolls values
                            WriteTollInfo(f1, evtFeat, strLineI, strLineJ, "AM", intOneway, dir)
                            WriteTollInfo(f2, evtFeat, strLineI, strLineJ, "MD", intOneway, dir)
                            WriteTollInfo(f3, evtFeat, strLineI, strLineJ, "PM", intOneway, dir)
                            WriteTollInfo(f4, evtFeat, strLineI, strLineJ, "EV", intOneway, dir)
                            WriteTollInfo(f5, evtFeat, strLineI, strLineJ, "NI", intOneway, dir)
                            dct.Add(lEdgeID, 1)
                        End If

                        pFtEdge = pFCs_Edge.NextFeature
                    Loop
                End If
                pRow_Prj = pCs_Prj.NextRow
            Loop

            'Toll attributes:     'Loop through all rows in modeTolls table
            pCs_Mode = g_TblMode.Search(Nothing, False)
            pRow_Mode = pCs_Mode.NextRow

            Do Until pRow_Mode Is Nothing
                'Get the PSRCEdgeID
                strEdgeID = pRow_Mode.Value(fldEdgeID2)
                pFilt.WhereClause = "PSRCEdgeID=" & strEdgeID
                pFCs_Edge = pFClsEdge.Search(pFilt, False)
                pFtEdge = pFCs_Edge.NextFeature
                If pFtEdge Is Nothing Then
                Else
                    intOneway = pFtEdge.Value(fldOneWay)
                    iNode = pFtEdge.Value(fldINode)
                    JNode = pFtEdge.Value(fldJNode)
                    iUseEmmeN = pFtEdge.Value(pFtEdge.Fields.FindField("UseEmmeN"))

                    Select Case iUseEmmeN
                        Case 0
                            iNode = iNode + m_Offset
                            JNode = JNode + m_Offset
                        Case 1
                            iNode = dctReservedNodes.Item(iNode)
                            JNode = JNode + m_Offset
                        Case 2
                            JNode = dctReservedNodes.Item(JNode)
                            iNode = iNode + m_Offset
                        Case 3
                            iNode = dctReservedNodes.Item(iNode)
                            JNode = dctReservedNodes.Item(JNode)
                    End Select

                    strLineI = iNode & " " & JNode & " "
                    strLineJ = JNode & " " & iNode & " "
                    lEdgeID = pFtEdge.Value(fldEdgeID1)
                    If Not dct.ContainsKey(CType(lEdgeID, Integer)) Then

                        '            If pFtEdge.value(fldUpdate) = "Yes" Then 'with a tip project
                        'handled later with future projects
                        '? is it possible that a toll project and other non-toll project exist on the same edge?

                        '              '[082707] jaf:  changed mispelt field name to prRteID
                        '
                        '
                        '                pFilt.WhereClause = "projRteID=" & pFtEdge.value(fldPrjRte1)
                        '                 pCs_Line = evtMode.Search(pFilt, False)
                        '                 evtFeat = pCs_Line.NextFeature
                        '
                        '                If evtFeat Is Nothing Then    'didn't find the prjRteID in tblLineTolls
                        '                    'Use the modeTolls values
                        WriteTollInfo(f1, pRow_Mode, strLineI, strLineJ, "AM", intOneway)
                        WriteTollInfo(f2, pRow_Mode, strLineI, strLineJ, "MD", intOneway)
                        WriteTollInfo(f3, pRow_Mode, strLineI, strLineJ, "PM", intOneway)
                        WriteTollInfo(f4, pRow_Mode, strLineI, strLineJ, "EV", intOneway)
                        WriteTollInfo(f5, pRow_Mode, strLineI, strLineJ, "NI", intOneway)
                        '
                        '                Else    'find the prjRteID in tblLineTolls
                        '
                        '                    'Use the lineTolls values
                        '                    WriteTollInfo f1, evtFeat, strLineI, strLineJ, "AM", intOneway
                        '                    WriteTollInfo f2, evtFeat, strLineI, strLineJ, "MD", intOneway
                        '                    WriteTollInfo f3, evtFeat, strLineI, strLineJ, "PM", intOneway
                        '                    WriteTollInfo f4, evtFeat, strLineI, strLineJ, "EV", intOneway
                        '                    WriteTollInfo f5, evtFeat, strLineI, strLineJ, "NI", intOneway
                        '                End If
                        '
                        '            Else    'w/o a tip project
                        '                'use the modeTolls values
                        '                WriteTollInfo f1, pRow_Mode, strLineI, strLineJ, "AM", intOneway
                        '                WriteTollInfo f2, pRow_Mode, strLineI, strLineJ, "MD", intOneway
                        '                WriteTollInfo f3, pRow_Mode, strLineI, strLineJ, "PM", intOneway
                        '                WriteTollInfo f4, pRow_Mode, strLineI, strLineJ, "EV", intOneway
                        '                WriteTollInfo f5, pRow_Mode, strLineI, strLineJ, "NI", intOneway
                    End If

                End If
                pRow_Mode = pCs_Mode.NextRow
            Loop

            '    WriteLogLine "Done with modeTolls, scanning tblLineTolls"


            '     pCs_Line = evtMode.Search(Nothing, False)
            '
            '     evtFeat = pCs_Line.NextFeature
            '    Do Until evtFeat Is Nothing
            '    'pFilt.WhereClause = "prjRte='" & evtFeat.value(fldPrjRte2) & "'"
            '        With pFilt2
            '             .Geometry = evtFeat.ShapeCopy
            '            .GeometryField = evtMode.ShapeFieldName
            '            .SpatialRel = esriSpatialRelIntersects
            '        End With
            '         pFCs_Edge = pFClsEdge.Search(pFilt2, False)
            '         pFtEdge = pFCs_Edge.nextfeature
            '
            '        If pFtEdge Is Nothing Then
            '        Else
            '            strLineI = pFtEdge.value(fldINode) & " " & pFtEdge.value(fldJNode) & " "
            '             strLineJ = pFtEdge.value(fldJNode) & " " & pFtEdge.value(fldINode) & " "
            '            intOneway = pFtEdge.value(fldOneWay)
            '             pFilt = New QueryFilter
            '            pFilt.WhereClause = "PSRCEdgeID=" & pFtEdge.value(pFCs_Edge.FindField("PSRCEdgeID"))
            '             pCs_Mode = pTblMode.Search(pFilt, False)
            '            If pCs_Mode Is Nothing Then
            '                'Use the lineTolls values
            '                WriteTollInfo f1, evtFeat, strLineI, strLineJ, "AM", intOneway
            '                WriteTollInfo f2, evtFeat, strLineI, strLineJ, "MD", intOneway
            '                WriteTollInfo f3, evtFeat, strLineI, strLineJ, "PM", intOneway
            '                WriteTollInfo f4, evtFeat, strLineI, strLineJ, "EV", intOneway
            '                WriteTollInfo f5, evtFeat, strLineI, strLineJ, "NI", intOneway
            '            End If
            '        End If
            '
            '         evtFeat = pCs_Line.NextFeature
            '      Loop

            f1.Close()
            f2.Close()
            f3.Close()
            f4.Close()
            f5.Close()
            If Not pCs_Line Is Nothing Then

                System.Runtime.InteropServices.Marshal.ReleaseComObject(pCs_Line)
            End If
            If Not pCs_Mode Is Nothing Then

                System.Runtime.InteropServices.Marshal.ReleaseComObject(pCs_Mode)
            End If
            If Not pFCs_Edge Is Nothing Then

                System.Runtime.InteropServices.Marshal.ReleaseComObject(pFCs_Edge)
            End If

            pFilt2 = Nothing
            pFilt = Nothing
            evtFeat = Nothing
            pRow_Mode = Nothing
            pFtEdge = Nothing
            pCs_Line = Nothing
            pCs_Mode = Nothing
            pFCs_Edge = Nothing
            evtMode = Nothing
            'pTblMode = Nothing
            pFClsEdge = Nothing
            psort = Nothing
            WriteLogLine("FINISHED createTollsFile at " & Now())
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        Exit Sub
eh:
        CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.createTollsFile")
        MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.createTollsFile")
        'MsgBox Err.Description, vbExclamation, "createTollsFile"
    End Sub
    Public Function WriteTollInfo(ByVal objFile As Object, ByVal pRow As IRow, ByVal strLineI As String, ByVal strLineJ As String, ByVal strTOD As String, ByVal intOneway As Integer, Optional ByVal dir As Integer = 1)
        '       0 = one way from I
        '       1 = one way from J
        '       2 = two way
        'dir = 1: the project route direction is WITH the edge direction
        'dir = -1: the project route direction is AGAINST the edge direction

        'on error GoTo eh
        If fVerboseLog Then WriteLogLine("called WriteTollInfo")
        If fVerboseLog Then WriteLogLine("--------------------")

        Dim strLine2 As String
        With pRow

            Select Case intOneway
                Case 0
                    If dir = 1 Then
                        strLine2 = strLineI & .value(.Fields.FindField("IJTollSOV" & strTOD)) & " " _
                        & .value(.Fields.FindField("IJTollHOV2" & strTOD)) & " " _
                        & .value(.Fields.FindField("IJTollHOV3" & strTOD)) & " " _
                        & .value(.Fields.FindField("IJTollTrkLt" & strTOD)) & " " _
                        & .value(.Fields.FindField("IJTollTrkMed" & strTOD)) & " " _
                        & .value(.Fields.FindField("IJTollTrkHvy" & strTOD))
                    Else
                        strLine2 = strLineI & .value(.Fields.FindField("JITolSOV" & strTOD)) & " " _
                        & .value(.Fields.FindField("JITollHOV2" & strTOD)) & " " _
                        & .value(.Fields.FindField("JITollHOV3" & strTOD)) & " " _
                        & .value(.Fields.FindField("JITollTrkLt" & strTOD)) & " " _
                        & .value(.Fields.FindField("JITollTrkMed" & strTOD)) & " " _
                        & .value(.Fields.FindField("JITolTrkHvy" & strTOD))
                    End If
                    objFile.WriteLine(strLine2)
                Case 1
                    If dir = 1 Then
                        strLine2 = strLineJ & .value(.Fields.FindField("JITolSOV" & strTOD)) & " " _
                        & .value(.Fields.FindField("JITollHOV2" & strTOD)) & " " _
                        & .value(.Fields.FindField("JITollHOV3" & strTOD)) & " " _
                        & .value(.Fields.FindField("JITollTrkLt" & strTOD)) & " " _
                        & .value(.Fields.FindField("JITollTrkMed" & strTOD)) & " " _
                        & .value(.Fields.FindField("JITolTrkHvy" & strTOD))
                    Else
                        strLine2 = strLineJ & .value(.Fields.FindField("IJTollSOV" & strTOD)) & " " _
                        & .value(.Fields.FindField("IJTollHOV2" & strTOD)) & " " _
                        & .value(.Fields.FindField("IJTollHOV3" & strTOD)) & " " _
                        & .value(.Fields.FindField("IJTollTrkLt" & strTOD)) & " " _
                        & .value(.Fields.FindField("IJTollTrkMed" & strTOD)) & " " _
                        & .value(.Fields.FindField("IJTollTrkHvy" & strTOD))
                    End If
                    objFile.WriteLine(strLine2)
                Case 2
                    If dir = 1 Then
                        strLine2 = strLineI & .value(.Fields.FindField("IJTollSOV" & strTOD)) & " " _
                        & .value(.Fields.FindField("IJTollHOV2" & strTOD)) & " " _
                        & .value(.Fields.FindField("IJTollHOV3" & strTOD)) & " " _
                        & .value(.Fields.FindField("IJTollTrkLt" & strTOD)) & " " _
                        & .value(.Fields.FindField("IJTollTrkMed" & strTOD)) & " " _
                        & .value(.Fields.FindField("IJTollTrkHvy" & strTOD))

                        objFile.WriteLine(strLine2)

                        strLine2 = strLineJ & .value(.Fields.FindField("JITollSOV" & strTOD)) & " " _
                        & .value(.Fields.FindField("JITollHOV2" & strTOD)) & " " _
                        & .value(.Fields.FindField("JITollHOV3" & strTOD)) & " " _
                        & .value(.Fields.FindField("JITollTrkLt" & strTOD)) & " " _
                        & .value(.Fields.FindField("JITollTrkMed" & strTOD)) & " " _
                        & .value(.Fields.FindField("JITollTrkHvy" & strTOD))

                        objFile.WriteLine(strLine2)
                    Else
                        strLine2 = strLineJ & .value(.Fields.FindField("IJTollSOV" & strTOD)) & " " _
                        & .value(.Fields.FindField("IJTollHOV2" & strTOD)) & " " _
                        & .value(.Fields.FindField("IJTollHOV3" & strTOD)) & " " _
                        & .value(.Fields.FindField("IJTollTrkLt" & strTOD)) & " " _
                        & .value(.Fields.FindField("IJTollTrkMed" & strTOD)) & " " _
                        & .value(.Fields.FindField("IJTollTrkHvy" & strTOD))

                        objFile.WriteLine(strLine2)

                        strLine2 = strLineI & .value(.Fields.FindField("JITollSOV" & strTOD)) & " " _
                        & .value(.Fields.FindField("JITollHOV2" & strTOD)) & " " _
                        & .value(.Fields.FindField("JITollHOV3" & strTOD)) & " " _
                        & .value(.Fields.FindField("JITollTrkLt" & strTOD)) & " " _
                        & .value(.Fields.FindField("JITollTrkMed" & strTOD)) & " " _
                        & .value(.Fields.FindField("JITollTrkHvy" & strTOD))
                        objFile.WriteLine(strLine2)
                    End If
            End Select
        End With

        Exit Function
eh:
        CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.WriteTollInfo")
        MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.WriteTollInfo")
    End Function

Public Sub OpenLogFile(filepath As String, filename As String)
   'called at startup to initialize a generic log file for ALL modules
   'modules should call GlobalMod.WriteLogLine(message as string) to write to log
   'on error GoTo eh
   
        'If MsgBox("Do you want full details in the log (may be a huge file!)?", vbYesNo, "Layer Source") = vbYes Then fVerboseLog = True Else fVerboseLog = False
        'If MessageBox.Show("Do you want full details in the log (mayb be a huge file!)?", "Log?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then

        'End If
        'filepath = Application.Templates.Item(1)
        'filepath = Left(filepath, InStrRev(filepath, "\") - 1)
        FileClose(255)
        FileOpen(255, filepath & "\" & filename, OpenMode.Output, OpenAccess.Default, OpenShare.Shared)
        'Open filepath & "\" & filename For Output Shared As 255
        m_fLogOpen = True
        m_RunStartTime = Now()
        Print(255, "PSRC Combined VBA Extensions Log 091806 Run time=" & m_RunStartTime)
        'Print 2, "  MTP DSN=" & strDSNmtp()
        'Print 2, "  TIP DSN=" & strDSNtip()
        Print(255, "----------------------------------------------------------")
        MsgBox("An application flow/error log file has been created at: " & filepath & "\" & filename & " . Consult this file in event of error.", , "Log file created")
        Exit Sub
eh:
        MsgBox("Error #" & Err.Number & ", " & Err.Description, , "GlobalMod.OpenLogFile")
        GlobalMod.CloseLogFile("GlobalMod.OpenLogFile: eh closed log file due to " & Err.Description)
        Err.Clear()
    End Sub
Public Sub WriteLogLine(ByVal strOutput As String)
   'writes strOutput as a line to the debug log file
   On Error GoTo eh
   'check to see if log file is open and if so, write
   If m_fLogOpen Then
            PrintLine(255, strOutput)
   End If

   Exit Sub
eh:
   'error handler
        MsgBox("Error #" & Err.Number & ", " & Err.Description & ".  LOG ENTRIES MAY NOT BE COMPLETE.", vbInformation, "GlobalMod.writeLogLine")
   Err.Clear
End Sub
Public Sub CloseLogFile(strMsg As String)
   'called by error handlers or final program exit to close generic log file
   'on error GoTo eh
   'if we already closed the file, let's not get stuck again
   If Not (m_fLogOpen) Then
      'MsgBox "Log file is already closed but CloseLogFile called with " & strMsg, , "GlobalMod.CloseLogFile"
      Exit Sub
   End If

   If Len(strMsg) > 0 Then
      'we have a closing message to write before closing log file
            Print(255, "Log file closed with message: " & vbCrLf & strMsg)
   End If

   'write footer and close
        Print(255, "----------------------------------------------------------")
        Print(255, "Closed log started " & m_RunStartTime & " at " & Now())
        FileClose(255)
   m_fLogOpen = False

   Exit Sub
eh:
        MsgBox("Error #" & Err.Number & ", " & Err.Description, , "GlobalMod.CloseLogFile")
   Err.Clear
End Sub

Sub SelectByLocation0(pFLayer1 As IFeatureLayer, pFLayer2 As IFeatureLayer, pFSel1 As IFeatureSelection) 'transrefjunction, transrefedge
    ' selects all features in layer(0) that intersect any
    ' selected features in layer(1)

    Dim i As Integer

   'MsgBox "starting"
    pFSel1 = pFLayer1 'transrefjunction
    pFSel1.Clear
    Dim pFSel2 As IFeatureSelection
     pFSel2 = pFLayer2 'transrefedge
    If pFSel2.SelectionSet.count = 0 Then
            MsgBox("GlobalMod.SelectByLocation: nothing selected in " & pFLayer2.Name)
        Exit Sub
    End If

    Dim pFeatIndex As IFeatureIndex
     pFeatIndex = New FeatureIndex
     pFeatIndex.FeatureClass = pFLayer1.FeatureClass
        pFeatIndex.Index(Nothing, GetSelectedExtent(pFLayer2))
    Dim pIQ2 As IIndexQuery2
     pIQ2 = pFeatIndex

    Dim pFCur As IFeatureCursor
        pFSel2.SelectionSet.Search(Nothing, False, pFCur)

    Dim pDict                   'Create a variable
   pDict = CreateObject("Scripting.Dictionary")

    Dim pFeat As IFeature, l As Long
     pFeat = pFCur.NextFeature
    Do Until pFeat Is Nothing
            Dim vOIDs As Object
            pIQ2.IntersectedFeatures(pFeat.Shape, vOIDs)
            If Not vOIDs Is Nothing Then
                For l = 0 To UBound(vOIDs)
                    If Not pDict.ContainsKey(vOIDs(l)) Then
                        pDict.Add(vOIDs(l), 0)
                    End If
                Next l
            End If
         pFeat = pFCur.NextFeature
    Loop
        WriteLogLine("SelectByLocation Geometry bag complete; starting spatial filter at " & Now())
        'frmNetlayer.lblStatus.text = "GlobalMod.create_ScenarioShapefiles: Starting spatial filter"
  'pan 09/08/05 Label status should read select by location
        g_frmNetLayer.lblStatus.Text = "GlobalMod.SelectByLocation: Starting spatial filter"
  'pan frmNetLayer.Refresh

    If pDict.count = 0 Then Exit Sub
   ' MsgBox pDict.count
    Dim lOIDs() As Long
    ReDim lOIDs(pDict.count - 1)
        Dim vKey As Object
    l = 0
    For Each vKey In pDict.Keys
        lOIDs(l) = CLng(vKey)
        l = l + 1
    Next vKey
        pFSel1.SelectionSet.AddList(pDict.count, lOIDs(0))

   ' MsgBox pFSel1.SelectionSet.count
End Sub

Function GetSelectedExtent(pFSel As IFeatureSelection) As IEnvelope
    Dim pFCur As IFeatureCursor
        pFSel.SelectionSet.Search(Nothing, False, pFCur)
    Dim pFeat As IFeature
     pFeat = pFCur.NextFeature
    Dim pEnv As IEnvelope
    Do Until pFeat Is Nothing
        If pEnv Is Nothing Then
             pEnv = pFeat.Shape.Envelope.Envelope
        Else
                pEnv.Union(pFeat.Shape.Envelope)
        End If
         pFeat = pFCur.NextFeature
    Loop
     GetSelectedExtent = pEnv
End Function


Public Sub AddMyIndex(pFC As IFeatureClass, strfieldname As String, indexname As String)
  ' up fields
        Dim pFields As IFields
  Dim pfieldsedit As IFieldsEdit
  Dim pField As IField
  Dim lFld As Long
        pFields = New Fields
   pfieldsedit = pFields
        'pfieldsedit.FieldCount = 1
        'this field already exists
        lFld = pFC.FindField(strfieldname)
'  Exit Sub
   pField = pFC.Fields.field(lFld)
        ' pfieldsedit.field(0) = pField
  Dim pIndex As iIndex
  Dim pIndexEdit As IIndexEdit
   pIndex = New index
   pIndexEdit = pIndex

  With pIndexEdit
            .Fields_2 = pFields
            .Name_2 = indexname
  End With
  'add index to feature class
        pFC.AddIndex(pIndex)

   pFields = Nothing
   pfieldsedit = Nothing
   pField = Nothing
   pIndex = Nothing
   pIndexEdit = Nothing


    End Sub



    Public Sub GetLargestID(ByVal pTable As ITable, ByVal AttField As String, ByRef LargestAtt As Long)

        'on error GoTo eh

        Dim psort As ITableSort
        Dim pCs As ICursor, pRow As IRow
        psort = New TableSort

        'hyu - this block seems redundant, so is comment out.
        '    With pSort
        '         .QueryFilter = Nothing
        '        .Fields = pTable.OIDFieldName
        '        .Ascending(pTable.OIDFieldName) = False
        '         .Table = pTable
        '        .Sort Nothing
        '    End With
        '
        '     pCs = pSort.Rows
        '     pRow = pCs.NextRow

        With psort
            .QueryFilter = Nothing
            .Fields = AttField
            .Ascending(AttField) = False
            .Table = pTable
            .Sort(Nothing)
        End With

        pCs = psort.Rows
        pRow = pCs.NextRow
        LargestAtt = pRow.Value(pTable.FindField(AttField))

        Exit Sub

eh:
        MsgBox("Error: GetLargestID " & vbNewLine & Err.Description)

    End Sub


Public Sub Calculate(pFClass As IFeatureClass, thefield As String, ByRef thevalue As String, theselection As String)
'++ Calculate table field values
'on error GoTo ErrorHandler
    Dim mx As IMxDocument
    Dim pTable As ITable
    Dim pLayer As ILayer
    Dim pStTab As IStandaloneTable
    Dim pCalc As ICalculator
        'Dim fff As IUnknown
    Dim pFld As IField
    Dim pItem As ICommandItem
    Dim pCursor As ICursor
    Dim pfeatcls As IFeatureClass
    'Dim pFeatLayer As IFeatureLayer
    'Dim pEditor As IEditor
    'Dim pID As New UID
    Dim pQFilter As IQueryFilter
    Dim pCallBack As clsCalculator
    Dim rcnt As Long

     pfeatcls = pFClass
    pTable = pfeatcls
  '++ Make sure you are in edit mode
 'pID = "esriEditor.Editor"
' pEditor = m_App.FindExtensionByCLSID(pID)
 'If Not pEditor.EditState = esriStateEditing Then
 'MsgBox "Must be in edit mode to continue"
 'Exit Sub
 'End If

        If pTable.FindField(thefield) = -1 Then MsgBox("There is no field named [" & thefield & "] in the table" & thefield)
      '++ Only calculate for states whose name begins with
      '++ the letter A
     pQFilter = New QueryFilter

    pQFilter.WhereClause = theselection
     pCursor = pTable.Search(pQFilter, False)

    rcnt = pfeatcls.featurecount(pQFilter)

    '++ Simple calculation
     pCalc = New Calculator
     pCallBack = New clsCalculator

'MsgBox rcnt

    pCallBack.Get_Rowset (rcnt)

    With pCalc
         .Cursor = pCursor
         .Callback = pCallBack
        '++ Triple quotes for text string updates   """AA"""
        .Expression = thevalue
        .field = thefield
         '++ display a msgbox if error occurs
        .ShowErrorPrompt = True
    End With

    pCalc.Calculate

    '++ De-reference the cursor and calculator if saving shapefile changes
     pCursor = Nothing
     pCalc = Nothing
     pCallBack = Nothing

 Exit Sub
ErrorHandler:
        MsgBox("Calculate", Err.Number, Err.Description)
End Sub

    Public Function addField(ByVal pFCls As IFeatureClass, ByVal FieldName As String, ByVal fieldType As esriFieldType, Optional ByVal fieldLen As Integer = 0, Optional ByVal defaultV As Object = -999)




        Dim i_NewField As IField
        i_NewField = New Field

        Dim i_NewFieldEdit As IFieldEdit
        i_NewFieldEdit = i_NewField
        Try
            With i_NewFieldEdit
                .Name_2 = FieldName
                '.AliasName = pFieldAliasName
                .Type_2 = fieldType
                '.Editable = pIsEditable
                '.IsNullable = pIsNullable
                .Length_2 = fieldLen
                '        .DefaultValue = 0#
            End With

            pFCls.AddField(i_NewField)
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString)
        End Try
        
















        'Dim pField As IField = New Field
        'With pField
        '.'Type_2 = fieldType
        ' .name_2 = FieldName
        ' If fieldLen > 0 Then .Length_2 = fieldLen
        'If defaultV <> -999 Then .DefaultValue_2 = defaultV
        ' End With
        'pFCls.AddField(pField)



'        MsgBox("Error: globalMod_addField" & vbNewLine & Err.Description)
        'releaseObj:
        'pField = Nothing
        'pFld = Nothing
    End Function

Public Function deleteAllRows0(pWSedit As IWorkspaceEdit, pFCls As IFeatureClass)
    
        WriteLogLine("start deleting all rows: " & Now())
        Debug.Print("start deleting all rows: " & Now())
    Dim pFCS As IFeatureCursor
    
        Dim pFtSet As ISet
    Dim i As Integer
    Dim pFtEdit As IFeatureEdit, pFt As IFeature
        pFtSet = New ESRI.ArcGIS.esriSystem.SetClass
    i = 0
     pFCS = pFCls.Search(Nothing, False)
    
    Dim pDaStats As IDataStatistics
    Dim pStats As IStatisticsResults
    Dim lMax As Long, lMin As Long, l As Long
    Dim pFilt As IQueryFilter2
    
     pDaStats = New DataStatistics
    pDaStats.field = pFCls.OIDFieldName
     pDaStats.Cursor = pFCS
    pDaStats.SimpleStats = True
     pStats = pDaStats.Statistics
    If pFCls.featurecount(Nothing) = 0 Then Exit Function
    lMax = pStats.Maximum
    lMin = pStats.Minimum
    l = lMax
     pFilt = New QueryFilter
    
    'deleting rows in desceding order of OID
    Do Until lMax <= lMin
        If lMax Mod 10000 > 0 Then
        'the first round, usually the largest OID is not a multiplication of 10000
            If lMax > 10000 Then
                l = Val(Left(CStr(lMax), Len(CStr(lMax)) - 4) & "0000")
            Else
                l = 0
            End If
        Else
            l = lMax - 10000
        End If
        
        pFilt.WhereClause = pFCls.OIDFieldName & ">=" & CStr(l)
        
            If pWSedit.IsBeingEdited = False Then pWSedit.StartEditing(True)
        pWSedit.StartEditOperation
         pFCS = pFCls.Search(pFilt, False)
        
         pFt = pFCS.NextFeature
        Do Until pFt Is Nothing
             pFtEdit = pFt
                pFtSet.Add(pFtEdit)
            i = i + 1
            If i = 100 Then
                    pFtEdit.DeleteSet(pFtSet)
                    pFtSet = New ESRI.ArcGIS.esriSystem.SetClass
                i = 0
            End If
             pFt = pFCS.NextFeature
        Loop
            If pFtSet.count > 0 Then pFtEdit.DeleteSet(pFtSet)
        lMax = l
    
        pWSedit.StopEditOperation
            pWSedit.StopEditing(True)
    Loop
     pStats = Nothing
     pDaStats = Nothing
    
     pFtSet = Nothing
     pFtEdit = Nothing
    
        WriteLogLine("finish deleting all rows: " & Now())
        Debug.Print("end deleting all rows: " & Now())
     pFCS = Nothing
    pWSedit.StopEditOperation
        pWSedit.StopEditing(True)
        Debug.Print("finish saving deleting: " & Now())
End Function

Public Function deleteAllRows(pWSedit As IWorkspaceEdit, pTbl As ITable)
    
        WriteLogLine("start deleting all rows: " & Now())
        Debug.Print("start deleting all rows: " & Now())
    Dim pCs As ICursor
    
        Dim pRSet As ISet

    Dim i As Integer
    Dim pREdit As IRowEdit, pR As IRow
        pRSet = New ESRI.ArcGIS.esriSystem.Set
    i = 0
     pCs = pTbl.Search(Nothing, False)
    
    Dim pDaStats As IDataStatistics
    Dim pStats As IStatisticsResults
    Dim lMax As Long, lMin As Long, l As Long
    Dim pFilt As IQueryFilter2
    
     pDaStats = New DataStatistics
    pDaStats.field = pTbl.OIDFieldName
        pDaStats.Cursor = pCs
    pDaStats.SimpleStats = True
     pStats = pDaStats.Statistics
        If pTbl.RowCount(Nothing) = 0 Then Exit Function
        lMax = pStats.Maximum
        lMin = pStats.Minimum
        l = lMax
        pFilt = New QueryFilter

        'deleting rows in desceding order of OID
        Do Until lMax <= lMin
            If lMax Mod 10000 > 0 Then
                'the first round, usually the largest OID is not a multiplication of 10000
                If lMax > 10000 Then
                    l = Val(Left(CStr(lMax), Len(CStr(lMax)) - 4) & "0000")
                Else
                    l = 0
                End If
            Else
                l = lMax - 10000
            End If

            pFilt.WhereClause = pTbl.OIDFieldName & ">=" & CStr(l)

            If pWSedit.IsBeingEdited = False Then pWSedit.StartEditing(True)
            pWSedit.StartEditOperation()
            pCs = pTbl.Search(pFilt, False)

            pR = pCs.NextRow
            Do Until pR Is Nothing
                pREdit = pR
                pRSet.Add(pREdit)
                i = i + 1
                If i = 100 Then
                    pREdit.DeleteSet(pRSet)

                    pRSet = New ESRI.ArcGIS.esriSystem.Set

                    i = 0
                End If
                pR = pCs.NextRow
            Loop
            If pRSet.Count > 0 Then pREdit.DeleteSet(pRSet)
            lMax = l

            pWSedit.StopEditOperation()
            pWSedit.StopEditing(True)
        Loop

        pStats = Nothing
        pDaStats = Nothing

        pRSet = Nothing
        pREdit = Nothing

        WriteLogLine("finish deleting all rows: " & Now())
        Debug.Print("end deleting all rows: " & Now())
        pCs = Nothing
        pWSedit.StopEditOperation()
        pWSedit.StopEditing(True)
        Debug.Print("finish saving deleting: " & Now())
    End Function
    Public Sub getMAtts(ByVal pFClsEdge As IFeatureClass, ByVal pTbl As ITable, ByVal EdgeIDFld As String, ByVal dctEdge As Dictionary(Of Long, Feature), ByVal dctMAtt As Dictionary(Of Object, Object))

        '[062007] jaf:  appears to populate dctEdge and dctMAtt with ScenarioEdge and related modeAttributes records respectively

        Dim dctEdge2 As Dictionary(Of Object, Object)
        dctEdge2 = New Dictionary(Of Object, Object)
        If dctEdge Is Nothing Then dctEdge = New Dictionary(Of Long, Feature)
        If dctMAtt Is Nothing Then dctMAtt = New Dictionary(Of Object, Object)


        Dim pDS As IDataset
        pDS = pTbl
        Debug.Print("getMAtt: " & pDS.Name & " start at " & Now())

        Dim pCs As ICursor, pRow As IRow
        Dim pFCS As IFeatureCursor, pFt As IFeature
        Dim sEdgeID As String, sEdgeOID As String
        Dim fld As Long

        fld = pFClsEdge.FindField(EdgeIDFld)
        pFCS = pFClsEdge.Search(Nothing, False)
        pFt = pFCS.NextFeature
        Do Until pFt Is Nothing
            sEdgeOID = CStr(pFt.OID)
            sEdgeID = CStr(pFt.Value(fld))

            If Not dctEdge2.ContainsKey(sEdgeID) Then dctEdge2.Add(sEdgeID, sEdgeOID)
            '[050407] hyu
            'If Not dctEdge.ContainsKey(sEdgeOID) Then dctEdge.Add sEdgeOID, sEdgeOID
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


    Public Sub CreateFinalNetfiles(ByVal pathnameN As String, ByVal filenameN As String)
        'first merge the contents in the netfile and temporary files to netfile
        'there're blank lines in the nodes part due to splitting when create weavelinks
        'this function is to remove those blank lines.

        Dim NextLine As String
        Dim pathnameAm As String, pathnameM As String, pathnamePm As String, pathnameE As String, pathnameNi As String, pathnameA As String
        Dim TpathnameAm As String, TpathnameM As String, TpathnamePm As String, TpathnameE As String, TpathnameNi As String, TpathnameA As String

        'filename of final netfiles
        pathnameAm = pathnameN + "\" + filenameN + "AM.txt"
        pathnameM = pathnameN + "\" + filenameN + "M.txt"
        pathnamePm = pathnameN + "\" + filenameN + "PM.txt"
        pathnameE = pathnameN + "\" + filenameN + "E.txt"
        pathnameNi = pathnameN + "\" + filenameN + "N.txt"
        pathnameA = pathnameN + "\" + filenameN + "Attr.txt"

        'filename of temporary files
        TpathnameAm = pathnameN + "\" + filenameN + "AMtemp.txt"
        TpathnameM = pathnameN + "\" + filenameN + "Mtemp.txt"
        TpathnamePm = pathnameN + "\" + filenameN + "PMtemp.txt"
        TpathnameE = pathnameN + "\" + filenameN + "Etemp.txt"
        TpathnameNi = pathnameN + "\" + filenameN + "Ntemp.txt"
        TpathnameA = pathnameN + "\" + filenameN + "AttrTemp.txt"

        FileClose(11, 17, 14, 15, 16, 1, 7, 3, 4, 5, 6)

        'open temperary files

        FileOpen(11, TpathnameAm, OpenMode.Input, OpenAccess.ReadWrite, OpenShare.Default)
        'Open TpathnameAm For Input As 11
        FileOpen(17, TpathnameM, OpenMode.Input, OpenAccess.ReadWrite, OpenShare.Default)
        'Open TpathnameM For InputFileOpen(11, TpathnameAm, OpenMode.Input, OpenAccess.ReadWrite, OpenShare.Default) As 17
        FileOpen(14, TpathnamePm, OpenMode.Input, OpenAccess.ReadWrite, OpenShare.Default)
        'Open TpathnamePm For Input As 14
        FileOpen(15, TpathnameE, OpenMode.Input, OpenAccess.ReadWrite, OpenShare.Default)
        'Open TpathnameE For Input As 15
        FileOpen(16, TpathnameNi, OpenMode.Input, OpenAccess.ReadWrite, OpenShare.Default)
        'Open TpathnameNi For Input As 16

        FileOpen(1, pathnameAm, OpenMode.Append, OpenAccess.ReadWrite, OpenShare.Default)
        'Open pathnameAm For Append As 1
        FileOpen(7, pathnameM, OpenMode.Append, OpenAccess.ReadWrite, OpenShare.Default)

        FileOpen(4, pathnamePm, OpenMode.Append, OpenAccess.ReadWrite, OpenShare.Default)
        'Open pathnamePm For Append As 4
        FileOpen(5, pathnameE, OpenMode.Append, OpenAccess.ReadWrite, OpenShare.Default)
        'Open pathnameE For Append As 5
        FileOpen(6, pathnameNi, OpenMode.Append, OpenAccess.ReadWrite, OpenShare.Default)
        'Open pathnameNi For Append As 6


        'copy link records in temperory files to net files
        Do Until EOF(11)
            InputString(11, NextLine)
            If Trim(NextLine) <> "" Then Print(1, NextLine)
        Loop
        Do Until EOF(17)
            InputString(17, NextLine)
            If Trim(NextLine) <> "" Then Print(7, NextLine)
        Loop
        Do Until EOF(14)
            InputString(14, NextLine)
            If Trim(NextLine) <> "" Then Print(4, NextLine)
        Loop
        Do Until EOF(15)
            InputString(15, NextLine)
            If Trim(NextLine) <> "" Then Print(5, NextLine)
        Loop
        Do Until EOF(16)
            InputString(16, NextLine)
            If Trim(NextLine) <> "" Then Print(6, NextLine)
        Loop


        FileClose(11, 17, 14, 15, 16, 1, 7, 3, 4, 5, 6)

        'Make copies of the final netfiles to temporary files

        FileCopy(pathnameAm, TpathnameAm)
        FileCopy(pathnameM, TpathnameM)
        FileCopy(pathnamePm, TpathnamePm)
        FileCopy(pathnameE, TpathnameE)
        FileCopy(pathnameNi, TpathnameNi)
        FileCopy(pathnameA, TpathnameA)

        '    'delete the netfiles
        '    Kill pathnameAm
        '    Kill pathnameM
        '    Kill pathnamePm
        '    Kill pathnameE
        '    Kill pathnameNi
        '
        'create empty final netfiles
        FileOpen(1, pathnameAm, OpenMode.Output, OpenAccess.ReadWrite, OpenShare.Default)
        'Open pathnameAm For Output As 1
        FileOpen(7, pathnameM, OpenMode.Output, OpenAccess.ReadWrite, OpenShare.Default)
        'Open pathnameM For Output As 7
        FileOpen(4, pathnamePm, OpenMode.Output, OpenAccess.ReadWrite, OpenShare.Default)
        'Open pathnamePm For Output As 4
        FileOpen(5, pathnameE, OpenMode.Output, OpenAccess.ReadWrite, OpenShare.Default)
        'Open pathnameE For Output As 5
        FileOpen(6, pathnameNi, OpenMode.Output, OpenAccess.ReadWrite, OpenShare.Default)
        'Open pathnameNi For Output As 6
        FileOpen(3, pathnameA, OpenMode.Output, OpenAccess.ReadWrite, OpenShare.Default)
        'Open pathnameA For Output As 3

        'open temperary files
        FileOpen(11, TpathnameAm, OpenMode.Input, OpenAccess.ReadWrite, OpenShare.Default)
        'Open TpathnameAm For Input As 11
        FileOpen(17, TpathnameM, OpenMode.Input, OpenAccess.ReadWrite, OpenShare.Default)
        'Open TpathnameM For Input As 17
        FileOpen(14, TpathnamePm, OpenMode.Input, OpenAccess.ReadWrite, OpenShare.Default)
        'Open TpathnamePm For Input As 14
        FileOpen(15, TpathnameE, OpenMode.Input, OpenAccess.ReadWrite, OpenShare.Default)
        'Open TpathnameE For Input As 15
        FileOpen(16, TpathnameNi, OpenMode.Input, OpenAccess.ReadWrite, OpenShare.Default)
        'Open TpathnameNi For Input As 16
        FileOpen(13, TpathnameA, OpenMode.Input, OpenAccess.ReadWrite, OpenShare.Default)
        'Open TpathnameA For Input As 13

        'copy content in the temporary files to netfiles. skip those blank lines
        Do Until EOF(11)
            InputString(11, NextLine)
            If Trim(NextLine) <> "" Then Print(1, NextLine)
        Loop
        Do Until EOF(17)
            InputString(17, NextLine)
            If Trim(NextLine) <> "" Then Print(7, NextLine)
        Loop
        Do Until EOF(14)
            InputString(14, NextLine)
            If Trim(NextLine) <> "" Then Print(4, NextLine)
        Loop
        Do Until EOF(15)
            InputString(15, NextLine)
            If Trim(NextLine) <> "" Then Print(5, NextLine)
        Loop
        Do Until EOF(16)
            InputString(16, NextLine)
            If Trim(NextLine) <> "" Then Print(6, NextLine)
        Loop

        Do Until EOF(13)
            InputString(13, NextLine)
            If Trim(NextLine) <> "" Then Print(3, NextLine)
        Loop
        'close all opened files
        FileClose(11, 17, 14, 15, 16, 13, 1, 7, 3, 4, 5, 6)
    End Sub




    Public Sub ArchivingNetwork()
        '9/9/2003 new way to archive
        '[091206] jaf:
        'On Error GoTo eh
        MsgBox("ArchivingNetwork Called but is not activated.", , "GlobalMod.ArchivingNetwork")
        Exit Sub

        If GlobalMod.m_fNetChanged Then
            WriteLogLine("GlobalMod.ArchivingNetwork: Begin archiving TransRef features")
        Else
            WriteLogLine("GlobalMod.ArchivingNetwork: Exiting without running; no archivable changes were made")
            Exit Sub
        End If

        m_Map.ClearSelection()
        Dim PMx As IMxDocument
        PMx = m_App.Document
        PMx.ActiveView.Refresh()
        Dim pStatusBar As IStatusBar
        Dim pProgbar As IStepProgressor
        pStatusBar = m_App.StatusBar
        pProgbar = pStatusBar.ProgressBar

        Dim pFeatLayer As IFeatureLayer
        pFeatLayer = get_MasterNetwork(1, m_Map) 'should be junction

        Dim pFeatLayerPR As IFeatureLayer
        pFeatLayerPR = get_FeatureLayer("sde.PSRC.ProjectRoutes")

        Dim pQF As IQueryFilter, pQFpr As IQueryFilter
        Dim pFeatSelection As IFeatureSelection, pFeatSelectPR As IFeatureSelection
        pFeatSelectPR = pFeatLayerPR
        pFeatSelectPR.SelectFeatures(Nothing, esriSelectionResultEnum.esriSelectionResultNew, False)
        pFeatSelection = pFeatLayer 'QI
        pFeatSelection.SelectFeatures(Nothing, esriSelectionResultEnum.esriSelectionResultNew, False)

        Dim pSelectSet As ISelectionSet, pSelectSetPR As ISelectionSet
        Dim pFeatCursor As IFeatureCursor, pFeatCursorPR As IFeatureCursor


        Dim pFeature As IFeature, pFeaturePR As IFeature

        Dim pDate As Date, fDate As Date
        pDate = Date.Now

        Dim pTablemode As ITable
        Dim pTablePoint As ITable
        Dim pTableALIne As ITable
        Dim pTableAPoint As ITable
        Dim pTableLine As ITable

        Dim pDatasetE As IEnumDataset
        Dim pDataset As IDataset
        Dim pWorkspace As IWorkspace
        pWorkspace = pFeatLayer.FeatureClass.FeatureDataset.Workspace
        pDatasetE = pWorkspace.Datasets(esriDatasetType.esriDTTable)
        pDataset = pDatasetE.Next

        Do While Not pDataset Is Nothing
            If (pDataset.Name = "sde.PSRC.modeAttributes") Then
                pTablemode = pDataset
            ElseIf (pDataset.Name = "sde.PSRC.tblLineProjectsArchive") Then
                pTableALIne = pDataset
            ElseIf (pDataset.Name = "sde.PSRC.tblPointProjectsArchive") Then
                pTableAPoint = pDataset
            ElseIf (pDataset.Name = "sde.PSRC.tblPointProjects") Then
                pTablePoint = pDataset
            ElseIf (pDataset.Name = "sde.PSRC.tblLineProjects") Then
                pTableLine = pDataset
            End If
            pDataset = pDatasetE.Next
        Loop

        Dim pTableCursor As ICursor, pTableCursorMA As ICursor
        Dim pRow As IRow, pNewRow As IRow, pMARow As IRow
        Dim pWorkspaceEdit As IWorkspaceEdit
        pWorkspaceEdit = pWorkspace

        Dim index As Long, prIndex As Long, prIndexA As Long, maIndex As Long, indexp As Long
        Dim indexE As Long
        Dim tempIndex As Long, tempIndex2 As Long
        Dim querystring As String, modequery As String, Aquerystring As String
        Dim check As Boolean

        Dim pField As IField
        Dim pFields As IFields
        Dim FieldCount As Integer
        Dim index1 As Long, indexT As Long

        WriteLogLine("GlobalMod.ArchivingNetwork: Done with setup")

        index = pTableLine.FindField("inServiceDate") 'way written in tblLineProjects
        'indexp = pFeatLayer.FeatureClass.FindField("PSRCEdgeID")
        indexE = pTableLine.FindField("PSRCEdgeID")

        maIndex = pTablemode.FindField("PSRCEdgeID")
        prIndex = pTableLine.FindField("projRteID")
        prIndexA = pTableALIne.FindField("projRteID")

        Dim pQFt As IQueryFilter
        pQFt = New QueryFilter
        '[jaf20041129] service dates are now small integers storing four digit YEAR
        'pQFt.WhereClause = "inServiceDate <= " + CStr(pDate) + "#"
        '[jaf20041129][NeedsAttention!] is the Service Date the proper filter for Archiving???
        '[jaf20041129][NeedsAttention!] changed from inServiceDate to dateLastModified for now...
        'pQFt.WhereClause = "inServiceDate <= " & CStr(pDate)
        pQFt.WhereClause = "dateLastModified = " & CStr(pDate)
        pTableCursor = pTableLine.Search(pQFt, False)
        pRow = pTableCursor.NextRow

        pProgbar.Position = 0
        pStatusBar.ShowProgressBar("Archiving Lines...", 0, pTableLine.RowCount(pQFt), 1, True)

        WriteLogLine("GlobalMod.ArchivingNetwork: Starting row enumeration for pTableLine")
        Do While Not pRow Is Nothing
            WriteLogLine("GlobalMod.ArchivingNetwork: archiving pTableLine with PSRCEdgeID = " & pRow.Value(indexE))
            modequery = "PSRCEdgeID = " & CStr(pRow.Value(indexE))
            Aquerystring = "PSRCEdgeID = " & CStr(pRow.Value(indexE))
            querystring = "projRteID ='" & CStr(pRow.Value(prIndex)) + "'"
            'check if it is already in archive
            pQF = New QueryFilter
            pQF.WhereClause = Aquerystring
            If (pTableALIne.RowCount(pQF) = 0) Then 'its has not been archived yet
                'get the row in modeAttributes for this edge
                pQF = New QueryFilter
                pQF.WhereClause = modequery
                If (pTablemode.RowCount(pQF) > 0) Then
                    pTableCursorMA = pTablemode.Search(pQF, False)
                    pMARow = pTableCursorMA.NextRow
                    'create new row in archive
                    'MsgBox "Creating Archive row"
                    pNewRow = pTableALIne.CreateRow
                    pQFpr = New QueryFilter
                    pQFpr.WhereClause = querystring
                    pSelectSetPR = pFeatSelectPR.SelectionSet
                    pSelectSetPR.Search(pQFpr, False, pFeatCursorPR)
                    'have it in project routes and now get info. need out of it
                    pFeaturePR = pFeatCursorPR.NextFeature

                    pFields = pMARow.Fields
                    For FieldCount = 0 To pFields.FieldCount - 1
                        pField = pFields.Field(FieldCount)
                        'If pField.Editable Then
                        If Not pField.Type = esriFieldType.esriFieldTypeOID And Not pField.Type = esriFieldType.esriFieldTypeGeometry Then
                            indexT = pNewRow.Fields.FindField(pField.Name)
                            If (indexT <> -1) Then
                                'MsgBox pNewRow.Fields.Field(indexT).name
                                pNewRow.Value(indexT) = pMARow.Value(FieldCount)
                            End If
                        End If
                        'End If
                    Next FieldCount
                    index1 = pRow.Fields.FindField("projDBS")
                    indexT = pNewRow.Fields.FindField("projDBS")
                    pNewRow.Value(indexT) = pRow.Value(index1)
                    index1 = pRow.Fields.FindField("projID_Ver")
                    indexT = pNewRow.Fields.FindField("projID_Ver")
                    pNewRow.Value(indexT) = pRow.Value(index1)
                    index1 = pRow.Fields.FindField("projRteID")
                    indexT = pNewRow.Fields.FindField("projRteID")
                    pNewRow.Value(indexT) = pRow.Value(index1)

                    pNewRow.Store()
                    'now assign mode new attributes from tblLine
                    pFields = pRow.Fields
                    For FieldCount = 0 To pFields.FieldCount - 1
                        pField = pFields.Field(FieldCount)
                        'If pField.Editable Then
                        If Not pField.Type = esriFieldType.esriFieldTypeOID And Not pField.Type = esriFieldType.esriFieldTypeGeometry Then
                            indexT = pMARow.Fields.FindField(pField.Name)
                            If (indexT <> -1) Then
                                pMARow.Value(indexT) = pRow.Value(FieldCount)
                            End If

                        End If
                        'End If
                    Next FieldCount

                    pMARow.Store()
                End If
            End If
            pRow = pTableCursor.NextRow
            pStatusBar.StepProgressBar()
        Loop
        pStatusBar.HideProgressBar()

        Dim indexJ As Long
        index = pTablePoint.FindField("inServiceDate") 'way written in tblPointProjects
        'indexp = pFeatLayer.FeatureClass.FindField("PSRCEdgeID")
        indexJ = pTablePoint.FindField("PSRCJunctID")
        prIndex = pTablePoint.FindField("projRteID")
        prIndexA = pTableAPoint.FindField("projRteID")
        pQFt.WhereClause = "inServiceDate <= #" + CStr(pDate) + "#"
        pTableCursor = pTablePoint.Search(pQFt, False)
        pRow = Nothing
        pRow = pTableCursor.NextRow

        pProgbar.Position = 0
        pStatusBar.ShowProgressBar("Archiving Points...", 0, pTablePoint.RowCount(pQFt), 1, True)

        WriteLogLine("GlobalMod.ArchivingNetwork: Starting row enumeration for pTablePoint")
        Do While Not pRow Is Nothing
            WriteLogLine("GlobalMod.ArchivingNetwork: archiving pTablePoint with PSRCJunctID = " & pRow.Value(indexJ))
            Aquerystring = "PSRCJunctID = " + CStr(pRow.Value(indexJ))
            querystring = "projRteID ='" + CStr(pRow.Value(prIndex)) + "'"
            'check if it is already in archive
            pQF = New QueryFilter
            pQF.WhereClause = Aquerystring

            If (pTableAPoint.RowCount(pQF) = 0) Then 'its has not been archived yet
                'only attributes are in transrefjunction
                pSelectSet = pFeatSelection.SelectionSet ' to transrefjunction
                pSelectSet.Search(pQF, False, pFeatCursor)
                MsgBox(pSelectSet.Count)
                pFeature = pFeatCursor.NextFeature
                'create new row in archive
                pNewRow = pTableAPoint.CreateRow
                'copy attributes from transrefjunction to archive
                pFields = pFeature.Fields
                For FieldCount = 0 To pFields.FieldCount - 1
                    pField = pFields.Field(FieldCount)
                    'If pField.Editable Then
                    If Not pField.Type = esriFieldType.esriFieldTypeOID And Not pField.Type = esriFieldType.esriFieldTypeGeometry Then
                        indexT = pNewRow.Fields.FindField(pField.Name)
                        If (indexT <> -1) Then
                            MsgBox(pField.Name)
                            pNewRow.Value(indexT) = pFeature.Value(FieldCount)
                        End If
                    End If
                    'End If
                Next FieldCount
                index1 = pRow.Fields.FindField("projDBS")
                indexT = pNewRow.Fields.FindField("projDBS")
                pNewRow.Value(indexT) = pRow.Value(index1)
                index1 = pRow.Fields.FindField("projID_Ver")
                indexT = pNewRow.Fields.FindField("projID_Ver")
                pNewRow.Value(indexT) = pRow.Value(index1)
                index1 = pRow.Fields.FindField("projRteID")
                indexT = pNewRow.Fields.FindField("projRteID")
                pNewRow.Value(indexT) = pRow.Value(index1)

                pNewRow.Store()
                'now assign transrefjunct new attributes from tblPoint
                'have to be in edit mode
                pWorkspaceEdit.StartEditing(False)
                pWorkspaceEdit.StartEditOperation()

                pFields = pRow.Fields
                For FieldCount = 0 To pFields.FieldCount - 1
                    pField = pFields.Field(FieldCount)
                    'If pField.Editable Then
                    If Not pField.Type = esriFieldType.esriFieldTypeOID And Not pField.Type = esriFieldType.esriFieldTypeGeometry Then
                        indexT = pFeature.Fields.FindField(pField.Name)
                        If (indexT <> -1) Then
                            MsgBox(pField.Name)
                            pFeature.Value(indexT) = pRow.Value(FieldCount)
                        End If
                    End If
                    'End If
                Next FieldCount

                pFeature.Store()
                pWorkspaceEdit.StopEditOperation()
                pWorkspaceEdit.StopEditing(True)
            End If
            pRow = pTableCursor.NextRow
            pStatusBar.StepProgressBar()
        Loop

        pStatusBar.HideProgressBar()
        'initializes globals 'do in terminate frmUncodedPRojects
        'unInit
        m_Map.ClearSelection()
        PMx = m_App.Document
        PMx.ActiveView.Refresh()

        pFeatLayer = Nothing
        pTableLine = Nothing
        pTablePoint = Nothing
        pTableALIne = Nothing
        pTableAPoint = Nothing
        pRow = Nothing
        pNewRow = Nothing
        pMARow = Nothing
        pFeature = Nothing
        Exit Sub
eh:
        MsgBox("Error #" & Err.Number & ", " & Err.Description, vbInformation, "ArchivingNetwork")
        GlobalMod.CloseLogFile("GlobalMod.ArchivingNetwork: eh closed log file due to " & Err.Description)
        Err.Clear()

        If (pStatusBar.Visible = True) Then
            'MsgBox "closing statusbar"
            pStatusBar.HideProgressBar()
        End If
        If pWorkspaceEdit.IsBeingEdited Then
            pWorkspaceEdit.AbortEditOperation()
            pWorkspaceEdit.StopEditing(False)
        End If
    End Sub
    Public Sub create_modeAtt()
        'when a new edge is created, must have a linking rwo in modeAttribute table
        On Error GoTo eh


        Dim pTable As ITable
        Dim pWorkspace As IWorkspace
        Dim pRow As IRow

        pWorkspace = get_Workspace()
        pTable = get_TableClass("modeAttributes")

        pRow = pTable.CreateRow
        Dim index As Long
        index = pTable.Fields.FindField("PSRCEdgeID")
        If (index = -1) Then
            MsgBox("modeAttributes is missing the PSRCEdgeID field. Please exit software and add this field!")
            Exit Sub
        End If
        pRow.Value(index) = bigEid
        GlobalMod.WriteLogLine("GlobalMod.create_modeAtt: creating row for EdgeID=" & CStr(pRow.Value(index)))
        pRow.Store()

        Exit Sub
eh:
        MsgBox("Error #" & Err.Number & Err.Description, , "GlobalMod.create_modeAtt")
        GlobalMod.CloseLogFile("GlobalMod.create_modeAtt: eh closed log file with " & Err.Description)
        Err.Clear()
    End Sub
    Public Sub create_ProjectRoute(ByVal pGeometryBag As IGeometryCollection)
        'jaf--creates new project route feature from ref features in pGeometryBag
        'NOTE: there can be no route event source feature classes in map document to run this
        On Error GoTo eh

        'union all selected edges into one route
        Dim pMseg As IMSegmentation
        Dim pMA As IMAware
        Dim pGeom As IGeometry
        Dim pPolyline As IPolyline
        Dim pLine As ILine
        Dim pTopoOp As ITopologicalOperator
        'Dim pNewFeature As IFeature
        Dim pNewFeature2 As IFeature
        Dim pFeatLayer As IFeatureLayer
        pFeatLayer = get_FeatureLayer("sde.PSRC.ProjectRoutes")

        'now  up for editing!
        Dim pWorkspaceEdit As IWorkspaceEdit
        pWorkspaceEdit = get_Workspace()
        pWorkspaceEdit.StartEditing(False)
        pWorkspaceEdit.StartEditOperation()

        Dim tempIndex As Long
        ' pNewFeature = pFeatLayer.FeatureClass.CreateFeature
        pNewFeature2 = pFeatLayer.FeatureClass.CreateFeature

        pPolyline = New Polyline
        pTopoOp = pPolyline
        pTopoOp.ConstructUnion(pGeometryBag)
        pGeom = pPolyline
        pMA = pPolyline
        pMA.MAware = True
        pMseg = pPolyline
        pMseg.SetMsAsDistance(False)

        'jaf--put the actual shape into the new feature
        pNewFeature2.Shape = pGeom

        'jaf-- the feature attribs of the new feature
        If (databaseOpen = "TIP") Then
            tempIndex = pNewFeature2.Fields.FindField("projDBS")
            pNewFeature2.Value(tempIndex) = "TIP"
        Else 'its mtp project
            tempIndex = pNewFeature2.Fields.FindField("projDBS")
            pNewFeature2.Value(tempIndex) = "MTP"
        End If
        tempIndex = pNewFeature2.Fields.FindField("projID_Ver")
        pNewFeature2.Value(tempIndex) = m_prjAttr.PrjId + "-" + CStr(m_prjAttr.version)
        tempIndex = pNewFeature2.Fields.FindField("projID")
        pNewFeature2.Value(tempIndex) = m_prjAttr.PrjId
        GlobalMod.WriteLogLine("GlobalMod.create_ProjectRoute: just wrote to new route m_prjAttr.PrjId=" & m_prjAttr.PrjId)
        tempIndex = pNewFeature2.Fields.FindField("version")
        pNewFeature2.Value(tempIndex) = CStr(m_prjAttr.version)
        tempIndex = pNewFeature2.Fields.FindField("projRteID")
        bigRid = bigRid + 1
        m_prjAttr.PrjRteID = bigRid
        pNewFeature2.Value(tempIndex) = bigRid

        'jaf-- the length field to the shape's shape_length field: ANDY, DO WE NEED THIS?
        GlobalMod.WriteLogLine("GlobalMod.create_ProjectRoute: about to cc Shape_Length to Length")
        tempIndex = pNewFeature2.Fields.FindField("Shape_Length")
        Dim dblLen As Double
        dblLen = pNewFeature2.Value(tempIndex)
        tempIndex = pNewFeature2.Fields.FindField("Length")
        pNewFeature2.Value(tempIndex) = dblLen

        'jaf--this appears to commit the feature edits
        GlobalMod.WriteLogLine("GlobalMod.create_ProjectRoute: about to commit new route attribs")
        pNewFeature2.Store()

        'jaf--stop the editor
        pWorkspaceEdit.StopEditOperation()
        pWorkspaceEdit.StopEditing(True)
        'MsgBox "created a new feature"
        Exit Sub
eh:
        MsgBox("Error #" & Err.Number & ", " & Err.Description, vbInformation, "GlobalMod.create_ProjectRoute")
        GlobalMod.CloseLogFile("GlobalMod.create_ProjectRoute: eh closed log file due to " & Err.Description)
        Err.Clear()

        pWorkspaceEdit.AbortEditOperation()
        pFeatLayer = Nothing
        pNewFeature2 = Nothing

    End Sub
    Public Sub launchTableToEdit(ByVal tablename As String, ByVal Clstype As Long, ByVal id() As Long, ByVal count As Long)
        'new way of getting to attributes to edit that skips the domains
        'opens a table in the ArcMap window for editing
        'Clstype:
        '     0 = modeAttributes
        '     1 = a transref FeatureClass attrib table
        '     2 = one of the project outcomes tables (added JAF 091304)

        On Error GoTo eh
        GlobalMod.WriteLogLine("GlobalMod.launchTableToEdit: launching table " & tablename & ", id(0)=" & id(0))

        Dim pMxDoc As IMxDocument
        'Dim pUnknown As IUnknown
        Dim pLayer As ILayer
        Dim pStandaloneTable As IStandaloneTable
        Dim pTableWindow2 As ITableWindow2
        Dim pExistingTableWindow As ITableWindow
        Dim SetProperties As Boolean

        Dim pFeatLayer As IFeatureLayer
        'Get the selected item from the current contents view
        pMxDoc = m_App.Document
        pTableWindow2 = New TableWindow

        Dim pTable As ITable
        Dim pWS As IWorkspace
        pWS = get_Workspace()

        Dim pTableCursor As ICursor
        Dim pSelectionset As ISelectionSet
        Dim pQF As IQueryFilter
        pQF = New QueryFilter
        Dim pFeatSelect As IFeatureSelection

        'JAF 091304: commented out IF block in favor of CASE
        Select Case Clstype
            Case 0   'modeAttributes
                '   If Clstype = 0 Then 'opening modeattributes
                pTable = get_TableClass(tablename)
                pQF.WhereClause = "PSRCEdgeID >= " + CStr(id(0))
                'MsgBox pTable.rowcount(pQF)
                pSelectionset = pTable.Select(pQF, esriSelectionType.esriSelectionTypeHybrid, esriSelectionOption.esriSelectionOptionNormal, pWS)
                'pSelectionset.AddList count, id(0)
                ' A standalone table
                pStandaloneTable = New StandaloneTable
                pStandaloneTable.Table = pTable
                pExistingTableWindow = pTableWindow2.FindViaStandaloneTable(pStandaloneTable)
                ' Check if a table already exists; if not, create one

                If pExistingTableWindow Is Nothing Then
                    pTableWindow2.StandaloneTable = pStandaloneTable
                    SetProperties = True
                End If

            Case 1 'transRef feature class attrib table
                '   ElseIf Clstype = 1 Then 'opening transref table
                ' MsgBox "here open up junct"
                pFeatLayer = get_FeatureLayer(tablename)
                pFeatSelect = pFeatLayer
                pSelectionset = pFeatSelect.SelectionSet
                pSelectionset.AddList(count, id(0))
                ' Exit sub if item is not a feature layer or standalone table
                pExistingTableWindow = pTableWindow2.FindViaLayer(pFeatLayer)
                ' Check if a table already exists; if not create one
                If pExistingTableWindow Is Nothing Then
                    pTableWindow2.Layer = pFeatLayer
                    SetProperties = True
                End If

            Case 2 'project standalone table
                pTable = get_TableClass(tablename)
                pQF.WhereClause = "ProjRteID = " + CStr(id(0))
                'MsgBox pTable.rowcount(pQF)
                pSelectionset = pTable.Select(pQF, esriSelectionType.esriSelectionTypeHybrid, esriSelectionOption.esriSelectionOptionNormal, pWS)
                'pSelectionset.AddList count, id(0)
                ' A standalone table
                pStandaloneTable = New StandaloneTable
                pStandaloneTable.Table = pTable
                pExistingTableWindow = pTableWindow2.FindViaStandaloneTable(pStandaloneTable)
                ' Check if a table already exists; if not, create one
                If pExistingTableWindow Is Nothing Then
                    pTableWindow2.StandaloneTable = pStandaloneTable
                    SetProperties = True
                End If

        End Select
        '   End If
        ' pUnknown = pTable

        If SetProperties Then
            pTableWindow2.TableSelectionAction = esriTableSelectionActions.esriSelectFeatures
            pTableWindow2.ShowSelected = True
            pTableWindow2.UpdateSelection(pSelectionset)
            pTableWindow2.ShowAliasNamesInColumnHeadings = True
            pTableWindow2.Application = m_App
        Else
            pTableWindow2 = pExistingTableWindow
        End If

        ' Ensure Table Is Visible
        If Not pTableWindow2.IsVisible Then pTableWindow2.Show(True)
        pWS = Nothing
        pFeatLayer = Nothing
        pTable = Nothing

        Exit Sub
eh:
        MsgBox("Error #" & Err.Number & ", " & Err.Description, vbInformation, "GlobalMod.launchTableToEdit")
        GlobalMod.CloseLogFile("GlobalMod.launchTableToEdit: eh closed log file due to " & Err.Description)
        Err.Clear()

        pWS = Nothing
        pFeatLayer = Nothing
        pTable = Nothing
    End Sub
    Public Sub launchTblProjects(ByVal tbltype As Integer, ByVal lngPrjRteID As Long)
        'jaf--opens outcome attribute table and selects rec from latest project route creation
        'jaf--believe this is what lets user edit new route outcomes attribs

        'tblType argument:
        '  global constant projTypeLine = open tblLineProjects
        '  global constant projTypePoint = open tblPointProjects

        'lngProjRteID:  ProjectRoutes key for the project route row to edit in outcomes table


        'MsgBox "launching table"
        'new way of getting to attributes to edit that skips the domains
        'Dim pathname As String
        ' pathname = "c:\MODULEOUTPUTlaunch.txt"
        ' Open pathname For Output As #2

        Dim mydate As Date
        Dim header As String
        header = "GlobalMod.launchTblProjects:"
        mydate = Date.Now
        GlobalMod.WriteLogLine(header & " begin")
        'Print #2, "Maintenance launchTblProjects"
        On Error GoTo eh
        Dim pMxDoc As IMxDocument
        'Dim pUnknown As IUnknown
        Dim pLayer As ILayer
        Dim pStandaloneTable As IStandaloneTable
        Dim pTableWindow2 As ITableWindow2
        Dim pExistingTableWindow As ITableWindow
        Dim SetProperties As Boolean

        'Get the selected item from the current contents view
        pMxDoc = m_App.Document
        pTableWindow2 = New TableWindow

        Dim pTable As ITable
        Dim pWS As IWorkspace
        pWS = get_Workspace()

        If tbltype = projTypeLine Then
            pTable = get_TableClass("sde.PSRC.tblLineProjects")
            GlobalMod.WriteLogLine(header & " line table opening")
        Else
            pTable = get_TableClass("tblPointProjects")
            GlobalMod.WriteLogLine(header & " point table opening")
        End If

        'pUnknown = pTable

        Dim pTableCursor As ICursor
        Dim pSelectionset As ISelectionSet
        Dim pQF As IQueryFilter
        pQF = New QueryFilter

        pQF.WhereClause = "projRteID = " & lngPrjRteID
        'pQF.WhereClause = "projRteID = " + CStr(bigRid) '+ "'" 'this the current projrteid that was assigned

        pSelectionset = pTable.Select(pQF, esriSelectionType.esriSelectionTypeHybrid, esriSelectionOption.esriSelectionOptionNormal, pWS)
        ' MsgBox "launch" + CStr(pSelectionset.count)
        ' A standalone table
        pStandaloneTable = New StandaloneTable
        pStandaloneTable.Table = pTable
        pExistingTableWindow = pTableWindow2.FindViaStandaloneTable(pStandaloneTable)
        ' Check if a table already exists; if not, create one
        If pExistingTableWindow Is Nothing Then
            pTableWindow2.StandaloneTable = pStandaloneTable
            SetProperties = True
        End If
        GlobalMod.WriteLogLine(header & " creating standalone table ")

        If SetProperties Then
            pTableWindow2.TableSelectionAction = esriTableSelectionActions.esriSelectFeatures
            pTableWindow2.ShowSelected = True
            pTableWindow2.UpdateSelection(pSelectionset)
            pTableWindow2.ShowAliasNamesInColumnHeadings = True
            pTableWindow2.Application = m_App
        Else
            pTableWindow2 = pExistingTableWindow
        End If
        GlobalMod.WriteLogLine(header & " launching standalone table ")
        ' Ensure Table Is Visible
        If Not pTableWindow2.IsVisible Then pTableWindow2.Show(True)
        If (g_frmIDfacilities.chkboxEvents.Checked = True) Then

            GlobalMod.WriteLogLine(header & " have events!")
            g_frmEvent.LabelprjR.Text = GlobalMod.bigRid
            g_frmEvent.Show()
        Else
            'SEC- commented out below to make code run
            'g_frmIDfacilities.resetProjectList()
            GlobalMod.WriteLogLine(header & " reset idFacilities ")
        End If
        pWS = Nothing
        'Close #2
        pTable = Nothing
        Exit Sub
eh:
        MsgBox("Error #" & Err.Number & ", " & Err.Description, vbInformation, header)
        pWS = Nothing
        GlobalMod.CloseLogFile(header & " eh closed log file due to " & Err.Description)
        Err.Clear()

        pTable = Nothing
    End Sub



    Public Sub setDefinitionQueryGLOBAL(ByVal strWhere As String, ByVal strLayerName As String)
        'sets the Definition Query of layer strLayerName to strWhere
        Dim header As String
        header = "GlobalMod.setDefinitionQuery"
        On Error GoTo eh

        GlobalMod.WriteLogLine(header & ": creating FeatureLayerDefinition")
        Dim fl As IFeatureLayer
        fl = GlobalMod.get_FeatureLayer(strLayerName)
        Dim flDef As IFeatureLayerDefinition
        flDef = fl
        flDef.DefinitionExpression = strWhere
        GlobalMod.WriteLogLine("fl.fc=" & fl.Name & ", DefExpression=" & flDef.DefinitionExpression)

        'as usual, have to do fifty thousand extraneous things...
        GlobalMod.WriteLogLine(header & ": creating QueryFilter")
        Dim qF As IQueryFilter
        qF = New QueryFilter
        qF.WhereClause = strWhere
        Dim fFS As IFeatureSelection
        fFS = fl
        fFS.SelectFeatures(qF, esriSelectionResultEnum.esriSelectionResultNew, False)

        MsgBox("Selected clause: " & strWhere, , header)
        '   'finally, apply the Definition Query (I think?!?)
        '   GlobalMod.WriteLogLine header & ": setting Definition Query"
        '   Dim SelFeatLayer As IFeatureLayer
        '    SelFeatLayer = flDef.CreateSelectionLayer("Projects_Filtered", True, "", strWhere)

        Exit Sub
eh:
        'error handler
        MsgBox("Error #" & Err.Number & ", " & Err.Description, vbInformation, header)
        GlobalMod.CloseLogFile(header & ": eh closed log file due to " & Err.Description)
        Err.Clear()
    End Sub

    Public Sub Update_ProjectRoutes()
        'called to update project route feature attributes (basically reset version numbers)
        'updates route feature attributes, NOT feature geometry

        On Error GoTo eh
        Dim pMasterLayer As IFeatureLayer
        pMasterLayer = get_FeatureLayer("sde.PSRC.ProjectRoutes")

        'now  up for editing!
        Dim pWorkspaceEdit As IWorkspaceEdit
        pWorkspaceEdit = get_Workspace()
        pWorkspaceEdit.StartEditing(False)
        pWorkspaceEdit.StartEditOperation()

        Dim pFeatSelect As IFeatureSelection
        pFeatSelect = pMasterLayer
        pFeatSelect.SelectFeatures(Nothing, esriSelectionResultEnum.esriSelectionResultNew, False)

        Dim pSelectSet As ISelectionSet
        pSelectSet = pFeatSelect.SelectionSet
        Dim pFC As IFeatureCursor
        Dim pQF As IQueryFilter
        pQF = New QueryFilter

        pQF.WhereClause = "projID = '" + CStr(m_prjAttr.PrjId) + "'"
        'MsgBox pMasterLayer.FeatureClass.featurecount(pQF)
        pSelectSet.Search(pQF, False, pFC)

        Dim pFeature As IFeature
        pFeature = pFC.NextFeature
        Dim index As Long
        'update the concatenated id-ver
        index = pFeature.Fields.FindField("projID_Ver")
        pFeature.Value(index) = m_prjAttr.PrjId + "-" + CStr(m_prjAttr.version)
        'update the version
        index = pFeature.Fields.FindField("version")
        pFeature.Value(index) = m_prjAttr.version

        pFeature.Store()

        pWorkspaceEdit.StopEditOperation()
        pWorkspaceEdit.StopEditing(True)

        m_Map.ClearSelection()
        Dim PMx As IMxDocument
        PMx = m_App.Document
        PMx.ActiveView.Refresh()
        pMasterLayer = Nothing
        pFeature = Nothing
        Exit Sub
eh:
        MsgBox("Error #" & Err.Number & ", " & Err.Description, vbInformation, "GlobalMod.Update_ProjectRoutes")
        GlobalMod.CloseLogFile("GlobalMod.Update_ProjectRoutes:  eh closed log file due to " & Err.Description)
        Err.Clear()

        If pWorkspaceEdit.IsBeingEdited Then
            pWorkspaceEdit.AbortEditOperation()
            pWorkspaceEdit.StopEditing(False)
        End If

        pMasterLayer = Nothing
        pFeature = Nothing
    End Sub
    Public Sub UpdateIJ(ByVal edgeID As Long, ByVal splitjunct As Long)
        ' ij nodes
        On Error GoTo eh
        Dim pJuncLayer As IFeatureLayer
        Dim pedgeLayer As IFeatureLayer
        pJuncLayer = get_MasterNetwork(1, m_Map) 'should be junction
        pedgeLayer = get_MasterNetwork(0, m_Map) 'should be edge

        Dim pFeature As IFeature, pJFeat As IFeature

        pFeature = pedgeLayer.FeatureClass.GetFeature(edgeID)
        Dim lIndex As Long, slIndex As Long
        Dim lenname As IField
        lIndex = pedgeLayer.FeatureClass.FindField("length")
        lenname = pedgeLayer.FeatureClass.LengthField
        slIndex = pedgeLayer.FeatureClass.FindField(lenname.Name)
        pFeature.Value(lIndex) = pFeature.Value(slIndex)
        Dim pFC As IFeatureCursor
        Dim pFilter As ISpatialFilter
        Dim index As Long, indexE As Long

        pFilter = New SpatialFilter 'use a spatialfilter to get junctions this edge is connected to
        With pFilter
            .Geometry = pFeature.ShapeCopy
            .GeometryField = pJuncLayer.FeatureClass.ShapeFieldName
            .SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects
        End With

        pFC = pJuncLayer.Search(pFilter, False)
        pJFeat = pFC.NextFeature
        Dim nodecnt As Long
        nodecnt = 0
        Do Until pJFeat Is Nothing
            index = pJFeat.Fields.FindField("PSRCJunctID")
            If (IsDBNull(pJFeat.Value(index)) = False) Then
                'MsgBox "PSrcJunct" + CStr(pJFeat.value(index))
                If (nodecnt = 0) Then
                    indexE = pFeature.Fields.FindField("Inode")
                    pFeature.Value(indexE) = pJFeat.Value(index)
                    nodecnt = 1
                    'MsgBox " from" + CStr(pfeature.value(indexE))
                ElseIf (nodecnt = 1) Then
                    indexE = pFeature.Fields.FindField("JNode")
                    pFeature.Value(indexE) = pJFeat.Value(index)
                    nodecnt = 2
                    'MsgBox " to" + CStr(pfeature.value(indexE))
                End If
            Else 'its Null
                'MsgBox "PSRCJunctionID is null"
                If (nodecnt = 0) Then
                    indexE = pFeature.Fields.FindField("Inode")
                    pFeature.Value(indexE) = splitjunct
                    nodecnt = 1
                    'MsgBox " from" + CStr(pfeature.value(indexE))
                ElseIf (nodecnt = 1) Then
                    indexE = pFeature.Fields.FindField("JNode")
                    pFeature.Value(indexE) = splitjunct
                    nodecnt = 2
                    'MsgBox " to" + CStr(pfeature.value(indexE))
                End If
            End If

            pJFeat = pFC.NextFeature
        Loop

        pFeature.Store()
        pJFeat = Nothing
        pJuncLayer = Nothing
        pedgeLayer = Nothing
        Exit Sub
eh:
        MsgBox("Error #" & Err.Number & ", " & Err.Description, vbInformation, "GlobalMod.UpdateIJ")
        GlobalMod.CloseLogFile("GlobalMod.UpdateIJ: closed log file due to " & Err.Description)
        Err.Clear()

        pJFeat = Nothing
        pJuncLayer = Nothing
        pedgeLayer = Nothing
    End Sub
    Public Sub writeEvents(ByVal tbltype As Long, ByVal mstart As Double, ByVal mstop As Double)
        On Error GoTo eh
        Dim pWorkspace As IWorkspace
        pWorkspace = get_Workspace()
        Dim pTable As ITable
        'MsgBox "creating event"
        If (tbltype = 1) Then 'its an line event
            pTable = get_TableClass("evtLineProjectOutcomes")
        Else 'its a junction route
            pTable = get_TableClass("evtPointProjectOutcomes")
        End If

        'now  up for editing!
        Dim pWorkspaceEdit As IWorkspaceEdit
        pWorkspaceEdit = pWorkspace
        pWorkspaceEdit.StartEditing(False)
        pWorkspaceEdit.StartEditOperation()

        Dim pTableCursor As ICursor
        Dim pRow As IRow
        Dim pQF As IQueryFilter
        Dim tempIndex As Long
        Dim tempString As String
        'insert new row
        pRow = pTable.CreateRow
        'must  database and prjrteid here
        tempIndex = pRow.Fields.FindField("projDBS")

        pRow.Value(tempIndex) = databaseOpen
        tempIndex = pRow.Fields.FindField("projRteID")
        pRow.Value(tempIndex) = m_prjAttr.PrjRteID

        'need to store measures now
        If (tbltype = 1) Then 'its a line
            'MsgBox "line"
            tempIndex = pRow.Fields.FindField("MStart")
            pRow.Value(tempIndex) = mstart
            tempIndex = pRow.Fields.FindField("MStop")
            pRow.Value(tempIndex) = mstop
        Else
            'MsgBox "point"
            tempIndex = pRow.Fields.FindField("M")
            pRow.Value(tempIndex) = mstart
        End If
        tempIndex = pRow.Fields.FindField("projID_Ver")
        pRow.Value(tempIndex) = m_prjAttr.PrjId + "-" + CStr(m_prjAttr.version)
        pRow.Store()

        pWorkspaceEdit.StopEditOperation()
        pWorkspaceEdit.StopEditing(True)

        m_Map.ClearSelection()
        Dim pMxDoc As IMxDocument
        pMxDoc = m_App.Document
        pMxDoc.ActiveView.Refresh()
        pWorkspace = Nothing
        pTable = Nothing
        pRow = Nothing
        Exit Sub
eh:
        MsgBox("Error #" & Err.Number & ", " & Err.Description, vbInformation, "GlobalMod.writeEvents")
        GlobalMod.CloseLogFile("GlobalMod.writeEvents: eh closed log file due to " & Err.Description)
        Err.Clear()
    End Sub
    Public Sub writeProjOutcomesData(ByVal optiontype As Integer, ByVal transrefType As Integer)
        'Public Sub writeAttrData(optiontype As Integer, transrefType As Integer)
        'jaf--writes outcomes attributes data for project to tblLineProjects or tblPointProjects
        'arguments:
        '  optiontype:    1 = update existing proj rec, nonzero = new row needs insertion
        '  transreftype:  1 = line project, 0 = point project

        Dim pFeatLayer As IFeatureLayer
        Dim i As Long
        Dim pTable As ITable
        Dim header As String
        header = "GlobalMod.WriteProjOutcomesData"
        On Error GoTo eh

        GlobalMod.WriteLogLine(header & ": Called with optiontype=" & optiontype & ", transrefType=" & transrefType)

        For i = 0 To m_Map.LayerCount - 1
            'jaf--skip non-feature layers
            If (TypeOf m_Map.Layer(i) Is IFeatureLayer) Then
                If (m_Map.Layer(i).Name = "TransRefEdges") Then
                    pFeatLayer = m_Map.Layer(i)
                End If
            End If
        Next i
        Dim pWorkspace As IWorkspace
        pWorkspace = pFeatLayer.FeatureClass.FeatureDataset.Workspace

        If (transrefType = 1) Then 'its an edge route
            pTable = get_TableClass("sde.PSRC.tblLineProjects")
            GlobalMod.WriteLogLine(header & ": Will write to tblLineProjects")
        Else 'its a junction route
            pTable = get_TableClass("sde.PSRC.tblPointProjects")
            GlobalMod.WriteLogLine(header & ": OOOOOPS, will write to tblPointProjects!")
        End If

        'now  up for editing!
        Dim pWorkspaceEdit As IWorkspaceEdit
        pWorkspaceEdit = pWorkspace
        pWorkspaceEdit.StartEditing(False)
        pWorkspaceEdit.StartEditOperation()

        Dim pTableCursor As ICursor
        Dim pRow As IRow
        Dim pQF As IQueryFilter
        Dim tempIndex As Long
        Dim tempString As String

        Dim id As Long

        ' jaf--see note lower down containing word "dodge" why these are commented out
        '   Dim biggestRteID As Long
        '   biggestRteID = bigRid

        If (optiontype = 1) Then
            'update an existing record

            GlobalMod.WriteLogLine(header & ": Option=1, Commencing Existing Row Modification")

            pQF = New QueryFilter

            'load m_prjAttr with the project route ID
            GlobalMod.get_UpdatePrjRteID()

            'jaf--commented out if block, results are same either way
            'If (transrefType = 1) Then
            pQF.WhereClause = "projRteID = " + m_prjAttr.PrjRteID '+ "'"
            'Else
            '   pQF.WhereClause = "projRteID = " + m_prjAttr.PrjRteID '+ "'"
            'End If

            'do we have a row for this projRteID?
            If (pTable.RowCount(pQF) > 0) Then
                'we have an existing row
                GlobalMod.WriteLogLine(header & ": found outcomes record for prjrteID=" & m_prjAttr.PrjRteID)
            Else
                'error, expected row but found none
                MsgBox("ERROR: Found no rows in outcomes table for project routeID=" & m_prjAttr.PrjRteID, , "GlobalMod.writeProjOutcomesData")
                Exit Sub
            End If

            ' outcomes attributes values for all rows (should only be one)
            pTableCursor = pTable.Search(pQF, False)
            pRow = pTableCursor.NextRow
            Do While Not pRow Is Nothing
                'write the ID-Version concatenation
                tempIndex = pRow.Fields.FindField("projID_Ver")
                pRow.Value(tempIndex) = m_prjAttr.PrjId + "-" + CStr(m_prjAttr.version)

                'write the version number by itself
                tempIndex = pRow.Fields.FindField("projVer")
                pRow.Value(tempIndex) = m_prjAttr.version

                'write the inServiceDate (derived from the external DB completion year)
                tempIndex = pRow.Fields.FindField("inServiceDate")
                'pRow.value(tempIndex) = m_prjAttr.inservice
                'pRow.value(tempIndex) = CDate("12-31-" & CStr(m_prjAttr.inservice))
                '[jaf20041129] service dates are now small integers storing four digit YEAR
                pRow.Value(tempIndex) = m_prjAttr.inservice
                GlobalMod.WriteLogLine("GlobalMod.writeProjOutcomesData: just wrote inservice date of " & pRow.Value(tempIndex))

                'store it
                pRow.Store()
                m_prjOID = pRow.OID
                pRow = pTableCursor.NextRow
            Loop

            ' first, call to launchTblProjects MOVED 2-22-04 to frmIDFacilities.OKCmd_Click
            ' second, all this bigRID stuff was a dodge to hand launchTblProjects the route to select
            ' launchTblProjects has been modified to require an argument of the routeID
            '     bigRid = m_prjAttr.PrjRteID
            '      'now open the table for direct user editing
            '      GlobalMod.launchTblProjects transrefType
            '      bigRid = biggestRteID

        Else
            'need to insert NEW outcomes row

            'For i = 0 To edgesInroute - 1 (used to do this for ALL edges for section table method)

            GlobalMod.WriteLogLine(header & ": Option<>1, Commencing New Row Insertion")

            'insert new row
            pRow = pTable.CreateRow

            'write the source DB
            tempIndex = pRow.Fields.FindField("projDBS")
            GlobalMod.WriteLogLine("GlobalMod.writeProjOutcomesData: about to write projDBS=" & databaseOpen)
            pRow.Value(tempIndex) = databaseOpen
            'write the route ID
            tempIndex = pRow.Fields.FindField("projRteID")
            GlobalMod.WriteLogLine("GlobalMod.writeProjOutcomesData: about to write projRteID=" & m_prjAttr.PrjRteID)
            pRow.Value(tempIndex) = m_prjAttr.PrjRteID ' in create_projectRoute
            'write the version number by itself
            tempIndex = pRow.Fields.FindField("version")
            GlobalMod.WriteLogLine("GlobalMod.writeProjOutcomesData: about to write version=" & m_prjAttr.version)
            pRow.Value(tempIndex) = m_prjAttr.version
            'write the inServiceDate (derived from the external DB completion date)
            tempIndex = pRow.Fields.FindField("inServiceDate")
            '[jaf20041129] service dates are now small integers storing four digit YEAR
            'pRow.value(tempIndex) = CDate("12-31-" & CStr(m_prjAttr.inservice))
            GlobalMod.WriteLogLine("GlobalMod.writeProjOutcomesData: about to write inservice date of " & pRow.Value(tempIndex))
            pRow.Value(tempIndex) = m_prjAttr.inservice
            GlobalMod.WriteLogLine("GlobalMod.writeProjOutcomesData: just wrote inservice date of " & pRow.Value(tempIndex))

            'write the project ID-version concatenation
            tempIndex = pRow.Fields.FindField("projID_Ver")
            pRow.Value(tempIndex) = m_prjAttr.PrjId & "-" & CStr(m_prjAttr.version)
            'write the base project ID
            tempIndex = pRow.Fields.FindField("projID")
            pRow.Value(tempIndex) = m_prjAttr.PrjId
            'write id (object ID?) from arrays of ref features used to make the route
            If (transrefType = 1) Then
                'line project:  not sure what to do yet as of 2-3-04
                'id = edgeArray(0)
                'tempIndex = pRow.Fields.FindField("PSRCEdgeID")
                'pRow.value(tempIndex) = id
            Else
                'point project:  store the underlying junction id
                id = junctionArray(0)
                tempIndex = pRow.Fields.FindField("PSRCJunctID")
                pRow.Value(tempIndex) = id
            End If

            'store it
            pRow.Store()
            ' global objectid
            m_prjOID = pRow.OID

            'Next i
        End If

        'stop edits
        pWorkspaceEdit.StopEditOperation()
        pWorkspaceEdit.StopEditing(True)

        m_Map.ClearSelection()
        Dim pMxDoc As IMxDocument
        pMxDoc = m_App.Document
        pMxDoc.ActiveView.Refresh()

        GoTo cleanup
eh:
        'error handler
        MsgBox("Error #" & Err.Number & ", " & Err.Description, vbInformation, "GlobalMod.writeProjOutcomesData")
        GlobalMod.CloseLogFile("GlobalMod.writeProjOutcomesData: eh closed log file due to " & Err.Description)
        Err.Clear()

        If pWorkspaceEdit.IsBeingEdited Then
            pWorkspaceEdit.AbortEditOperation()
            pWorkspaceEdit.StopEditing(False)
        End If
        'fall through to cleanup to clear memory...

cleanup:

        pWorkspace = Nothing
        pFeatLayer = Nothing
        pTable = Nothing
        pRow = Nothing
    End Sub
    Public Sub ZoomToLayer(ByVal strLayerName As String)
        'jaf--zooms to extent of layer named strLayerName
        'if strLayerName is not a valid layer name, nothing happens
        On Error GoTo eh
        Dim pMxDoc As IMxDocument
        Dim pMap As IMap
        Dim pActiveView As IActiveView
        Dim i As Integer

        pMxDoc = GlobalMod.m_App.Document
        pMap = pMxDoc.FocusMap
        pActiveView = pMap

        For i = 0 To pMap.LayerCount - 1
            If (pMap.Layer(i).Name = strLayerName) Then
                GlobalMod.WriteLogLine("GlobalMod.ZoomToLayer: Zooming to " & strLayerName)
                pActiveView.Extent = pMap.Layer(i).AreaOfInterest
                pActiveView.Refresh()
                Exit For
            End If
        Next i
        Exit Sub
eh:
        'error handler
        MsgBox("Error #" & Err.Number & ", " & Err.Description, vbInformation, "GlobalMod.zoomToLayer")
        GlobalMod.CloseLogFile("GlobalMod.zoomToLayer: eh closed log file due to " & Err.Description)
        Err.Clear()
    End Sub
    Public Function getAllNodes(ByVal dct As Dictionary(Of Object, Object))
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
    Public Function FindProjectRoute(ByVal strID As String) As Long
        'jaf--selects the ProjectRoute feature (if any) with ProjID=strID
        'returns ProjRteID if something found, 0 if nothing found
        On Error GoTo eh

        '    Dim pApp As IApplication   'this is for VBA
        '     pApp = Application   'this is for VBA
        Dim pMxDoc As IMxDocument
        Dim pMap As IMap
        Dim pActiveView As IActiveView
        Dim pFeatureLayer As IFeatureLayer
        Dim pFeatureSelection As IFeatureSelection
        Dim pQueryFilter As IQueryFilter
        Dim pFC As IFeatureCursor

        '    If GlobalMod.fDebug Then MsgBox "frmChkProjects.FindProjectRoute: about to try doc ref", , "Debug"
        ' pMxDoc = Application.Document     'this is for VBA
        'dll's get the app handle from the hook handed to the main class constructor
        pMxDoc = GlobalMod.m_App.Document
        '    If GlobalMod.fDebug Then MsgBox "frmChkProjects.FindProjectRoute: got doc ref", , "Debug"

        ' pMxDoc = pApp.Document
        pMap = pMxDoc.FocusMap
        pActiveView = pMap
        'If GlobalMod.fDebug Then MsgBox "frmChkProjects.FindProjectRoute: got map ref", , "Debug"

        pFeatureLayer = GlobalMod.get_FeatureLayer("sde.PSRC.ProjectRoutes")
        pFeatureSelection = pFeatureLayer 'QI
        '    If GlobalMod.fDebug Then MsgBox "frmChkProjects.FindProjectRoute: got featurelayer ref", , "Debug"

        'Create the query filter
        pQueryFilter = New QueryFilter
        pQueryFilter.WhereClause = "projID='" & strID & "'"
        '    If GlobalMod.fDebug Then MsgBox "frmChkProjects.FindProjectRoute:  up qry with WHERE=" & pQueryFilter.WhereClause, , "Debug"

        'Invalidate only the selection cache
        'Flag the original selection
        pActiveView.PartialRefresh(esriViewDrawPhase.esriViewGeoSelection, Nothing, Nothing)
        'Perform the selection
        pFeatureSelection.SelectFeatures(pQueryFilter, esriSelectionResultEnum.esriSelectionResultNew, False)
        '    If GlobalMod.fDebug Then MsgBox "frmChkProjects.FindProjectRoute: fired the qry", , "Debug"

        'Flag the new selection
        pActiveView.PartialRefresh(esriViewDrawPhase.esriViewGeoSelection, Nothing, Nothing)

        If pFeatureSelection.SelectionSet.Count < 1 Then
            'If pSelectSet.count < 1 Then
            FindProjectRoute = 0
        Else
            'we found one or more!
            '        If GlobalMod.fDebug Then MsgBox "We got one!", , "Debug"
            pFeatureSelection.SelectionSet.Search(Nothing, False, pFC)

            Dim pFeature As IFeature
            pFeature = pFC.NextFeature
            '        If GlobalMod.fDebug Then MsgBox "frmChkProjects.FindProjectRoute: got the feature", , "Debug"

            Dim index As Long
            index = pFeature.Fields.FindField("projRteID")
            If index < 0 Then
                'error
                MsgBox("Error=Unable to find field projRteID in ProjectRoutes", , "frmChkProjects.FindProjectRoute")
            End If
            '        If GlobalMod.fDebug Then MsgBox "Found feature", , "Debug"
            '        MsgBox "Feature OID=" & pfeature.OID, , "Results"
            '        MsgBox "Feature proj-ver=" & CStr(pfeature.value(index)), , "Results"
            FindProjectRoute = pFeature.Value(index)
        End If
        Exit Function

        'pfeature.Store
eh:
        MsgBox("Error=" & Err.Description, , "frmChkProjects.FindProjectRoute")
        'End
    End Function



    Public Function get_LargestRteID()
        'routine to get the last project route id
        On Error GoTo eh

        GlobalMod.WriteLogLine("GlobalMod.get_LargestRteID:  Begin")
        Dim pFeatLayer As IFeatureLayer
        pFeatLayer = get_FeatureLayer("sde.PSRC.ProjectRoutes")
        Dim pFeatSelect As IFeatureSelection
        Dim pSelectSet As ISelectionSet
        Dim pFC As IFeatureCursor
        Dim pFeature As IFeature

        pFeatSelect = pFeatLayer
        pFeatSelect.SelectFeatures(Nothing, esriSelectionResultEnum.esriSelectionResultNew, False)
        pSelectSet = pFeatSelect.SelectionSet
        pSelectSet.Search(Nothing, False, pFC)
        pFeature = pFC.NextFeature

        Dim idtemp As Long
        Dim index As Long
        index = pFeatLayer.FeatureClass.FindField("projRteID")
        If (index = -1) Then
            MsgBox("ProjectRoutes is missing the projRteID field. Please exit software and add this field!")
            Exit Function
        End If
        idtemp = 0

        '   If fDebug Then MsgBox "commencing pFeature loop", , "get_LargestRteID"
        Do While Not pFeature Is Nothing

            '    If fDebug Then MsgBox "projRteID=" & pfeature.value(Index), , "get_LargestRteID"
            If (IsDBNull(pFeature.Value(index)) = False) Then
                If (pFeature.Value(index) > idtemp) Then
                    If (pFeature.Value(index) <> "Number Nul") Then
                        idtemp = pFeature.Value(index)
                    Else
                        MsgBox("'Number Nul' as value for ProjRteID " + CStr(pFeature.OID))

                    End If
                End If
            End If
            pFeature = pFC.NextFeature
        Loop
        '   If fDebug Then MsgBox "largest" + CStr(idtemp), , "get_LargestRteID"
        get_LargestRteID = idtemp
        Exit Function
eh:
        MsgBox("Error#" & Err.Number & Err.Description, , "GlobalMod.get_LargestRteID")
        GlobalMod.CloseLogFile("GlobalMod.get_LargestRteID: eh closed log file due to " & Err.Description)
        Err.Clear()
    End Function



    Public Function LoadProjectIntoSelection(ByVal idProject As String)
    End Function
    Public Function strDSNmtp() As String
        'string to  up ADO connection to MTP DB
        'test database in Access
        'strDSNmtp = "FileDSN=X:\TRANS\DataSystemsIntegration\DataServerConnection\projectsMTP_TEST.dsn"
        'test database in server
        'strDSNmtp = "FileDSN=X:\TRANS\DataSystemsIntegration\DataServerConnection\SQLserverDSNforMTP_db_GIStest.dsn"
        'real database in server
        strDSNmtp = "FileDSN=X:\TRANS\DataSystemsIntegration\DataServerConnection\SQLserverDSNforMTP_db_4.dsn"
    End Function
    Public Function strDSNtip() As String
        'returns TIP DSN conn string
        'test database in server
        'strDSNtip = "FileDSN=X:\TRANS\DataSystemsIntegration\DataServerConnection\SQLserverDSNforRTIPdataSQL_GIStest.dsn"
        'real database in server
        strDSNtip = "FileDSN=X:\TRANS\DataSystemsIntegration\DataServerConnection\SQLserverDSNforRTIPdataSQL.dsn"
    End Function
    Public Function strWrapWild(ByVal strIn As String) As String
        'wraps input text in the wildcard chars (%)
        strWrapWild = "%" & strIn & "%"
        'MsgBox "strWrapWild Result = " & strWrapWild, , "Debug"
    End Function

    Public Function getWorkspaceOfWorkVersion()

        Dim pLayer As ILayer, pFLy As IFeatureLayer
        Dim pDS As IDataset
        Dim pWS As IWorkspace
        Dim pELy As IEnumLayer
        If m_Map.LayerCount > 0 Then
            pELy = m_Map.Layers
            pLayer = pELy.Next

            If pLayer.Valid Then
                Do Until pLayer Is Nothing
                    If pLayer.Valid Then
                        If TypeOf pLayer Is IFeatureLayer And pLayer.Name = m_layers(0) Then
                            pFLy = pLayer
                            pDS = pFLy
                            pWS = pDS.Workspace
                            getWorkspaceOfWorkVersion = pWS
                            GoTo ReleaseObjs
                        End If
                    End If
                    pLayer = pELy.Next
                Loop
            End If
        End If
        ' pLayer = pMap.Layer(1)


ReleaseObjs:
        pDS = Nothing
        pFLy = Nothing
        pLayer = Nothing


    End Function

    Public Function splitEmme2OutputRow(ByVal sLine As String, ByRef sSplit() As String)
        'split the input string into array.
        'the input string should be delimited by fixed width like the following example
        '12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
        '  inode  jnode      @sov   @hov2   @hov3  @vpool  @light  @mediu  @heavy  @voltr    @lid  @ltype    result
        ReDim sSplit(0 To 11)

        sSplit(0) = Trim(Mid(sLine, 0, 7))
        sSplit(1) = Trim(Mid(sLine, 8, 7))
        sSplit(2) = Trim(Mid(sLine, 15, 10))
        sSplit(3) = Trim(Mid(sLine, 25, 8))
        sSplit(4) = Trim(Mid(sLine, 33, 8))
        sSplit(5) = Trim(Mid(sLine, 41, 8))
        sSplit(6) = Trim(Mid(sLine, 49, 8))
        sSplit(7) = Trim(Mid(sLine, 57, 8))
        sSplit(8) = Trim(Mid(sLine, 65, 8))
        sSplit(9) = Trim(Mid(sLine, 73, 8))
        sSplit(10) = Trim(Mid(sLine, 81, 8))
        sSplit(11) = Trim(Mid(sLine, 89, 8))

    End Function


    Public Function getStandaloneTable(ByVal sTableName As String) As ITable



        Dim pStTblCol As IStandaloneTableCollection


        pStTblCol = m_Map

        Dim i As Integer
        For i = 0 To pStTblCol.StandaloneTableCount - 1
            If pStTblCol.StandaloneTable(i).Name = sTableName Then
                If pStTblCol.StandaloneTable(i).Valid = True Then
                    getStandaloneTable = pStTblCol.StandaloneTable(i)
                End If
            End If
        Next i

        pStTblCol = Nothing

    End Function

    Public Sub addVertex(ByVal pPline As IPolyline4, ByVal pPoint As IPoint)

        Dim pSegCol As ISegmentCollection, pSegCol2 As ISegmentCollection, pSegCol3 As ISegmentCollection
        Dim pHT As IHitTest
        Dim pHtPt As IPoint, dDist As Double, lPart As Long, lSeg As Long, bright As Boolean
        Dim pClone As IClone
        Dim l As Long
        Dim pPl As IPolyline4
        Dim bSplit As Boolean

        pClone = pPline
        pSegCol = pClone.Clone
        pHT = pSegCol
        pSegCol3 = New Polyline

        pPline.SplitAtPoint(pPoint, True, False, bSplit, lPart, lSeg)
        pSegCol2 = New Polyline
        Do While bSplit
            pSegCol = pPline
            For l = 0 To lSeg
                pSegCol2.AddSegment(pSegCol.Segment(l))
            Next l
            pSegCol.RemoveSegments(0, lSeg + 1, False)

            pPline = pSegCol
            If pSegCol.SegmentCount = 0 Then
                bSplit = False
            Else
                Dim pPt2 As IPoint
                Dim dAlong As Double, dFrom As Double
                '            pPline.QueryPointAndDistance esriNoExtension, pPoint, False, pPt2, dAlong, dFrom, bright
                '
                '            If dAlong > 0.5 Or pPline.length - dAlong > 0.5 Then
                pPline.SplitAtPoint(pPoint, True, False, bSplit, lPart, lSeg)
                '            Else
                '                bSplit = False
                '            End If
            End If
        Loop

        If pSegCol.SegmentCount > 0 Then pSegCol2.AddSegmentCollection(pSegCol)
        pPline = pSegCol2
        '    Do While pHT.HitTest(pPoint, 0.5, esriGeometryPartBoundary, pHtPt, dDist, lPart, lSeg, bRight)
        '         pSegCol2 = New Polyline
        '        For l = 0 To lSeg
        '            pSegCol2.AddSegment pSegCol.Segment(0)
        '            pSegCol.RemoveSegments 0, 1, False
        '        Next l
        '         pPl = pSegCol2
        '        pPl.SplitAtPoint pPoint, True, False, bSplit, lPart, lSeg
        '        pSegCol3.AddSegmentCollection pSegCol2
        '    Loop
        '
        '    If pSegCol.SegmentCount > 0 Then
        '        pSegCol3.AddSegmentCollection pSegCol
        '    End If
        '
        '     pPline = pSegCol3

    End Sub

    Public Sub updateTransitLines(ByVal pPoint As IPoint)
        Dim pFClsTr As IFeatureClass
        ' pFClsTr = get_FeatureLayer(m_layers(13)).FeatureClass
        pFClsTr = getInterimTransitLines()

        Dim pSegCol As ISegmentCollection
        pSegCol = New Polygon
        pSegCol.SetCircle(pPoint, 0.5)

        Dim pSFilt As ISpatialFilter
        pSFilt = New SpatialFilter
        With pSFilt
            .Geometry = pSegCol
            .GeometryField = pFClsTr.ShapeFieldName
            .SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects
        End With

        Dim pFCS As IFeatureCursor
        Dim pFt As IFeature
        Dim pPline As IPolyline4
        ' pFCS = pFClsTr.Search(pSFilt, False)
        pFCS = pFClsTr.Update(pSFilt, False)
        pFt = pFCS.NextFeature
        Do Until pFt Is Nothing
            pPline = pFt.Shape
            addVertex(pPline, pPoint)
            pFt.Shape = pPline
            'pFt.Store
            pFCS.UpdateFeature(pFt)
            pFt = pFCS.NextFeature
        Loop

        pFCS = Nothing
        pSFilt = Nothing
        pFClsTr = Nothing

    End Sub


    Public Function getInterimTransitLines() As IFeatureClass
        Dim pWSF As IWorkspaceFactory2
        Dim pFWS As IFeatureWorkspace
        pWSF = New ShapefileWorkspaceFactory
        pFWS = pWSF.OpenFromFile(g_sOutPath, 0)
        getInterimTransitLines = pFWS.OpenFeatureClass(g_sOutName & "TransitLines")

        pFWS = Nothing
        pWSF = Nothing
    End Function

    Public Function writeTransitNode(ByVal lTOD As Long, ByVal sLine As String)
        Dim t As Integer
        If lTOD = 0 Then
            For t = 1 To 5
                PrintLine(t, sLine)
            Next t
        Else
            PrintLine(lTOD, sLine)
        End If
    End Function

    Public Function SortProjectsByYear()
        'sort the project by in service year in ascending order
        Dim psort As ITableSort
        Dim pTbl As ITable
        Dim pFilt As IQueryFilter2
        Dim pCs As ICursor
        Dim pRow As IRow
        Dim i As Integer
        Dim prj As ClassPrjSelect ' PrjSelected

        pTbl = get_TableClass(m_layers(15))
        pFilt = New QueryFilter
        pFilt.WhereClause = "Scenario_ID=" & m_ScenarioId
        psort = New TableSort

        With psort
            .Table = pTbl
            .Fields = g_InSvcDate ' "InServiceDate"
            .Ascending(g_InSvcDate) = True
            .QueryFilter = pFilt
            .Sort(Nothing)
        End With
        pCs = psort.Rows

        pRow = pCs.NextRow

        For i = prjselectedCol.Count To 1 Step -1
            prjselectedCol.Remove(i)
        Next i

        Do Until pRow Is Nothing
            prj = New ClassPrjSelect
            prj.PrjId = pRow.Value(pRow.Fields.FindField("ProjRteID")) '"ProjID_Ver"))
            ' prjselectedCol.Item(i) = prj
            prjselectedCol.Add(prj)
            pRow = pCs.NextRow
        Loop

        '     SortProjectsByYear = pSort.Rows
    End Function

    Public Function UpdateProjectEdgeAttributes(ByVal ScenarioID As Long, ByVal pFtPrj As IFeature, ByVal pAttPrj As IRow, ByVal pFtEdge As IFeature, ByVal pJctFCls As IFeatureClass, ByRef dirPrj As Integer)
        'if the edge is not in the project edge attributes table, copy the modes attributes
        'find the direction relationship between the project and the underlying edge.
        'update the attributes by those in the tblLineProjects.
        'the junction featureclass (pJctFcls) should be the TransRefJunction.  In case the project routes overlies edges/junctions _
        ' that are not included in the ScenarioEdge/ScenarioJunctions, the junction location needs to be get from the TransRefJunction layer.

        '072010-SEC-This procedure does deal with the problem of overlapping projects!!!!!!!!!!!!!Check again to make sure logic is Correct


        Dim sPrjID As String, lEdgeID As Long
        Dim pTblAtt As ITable
        Dim pTblMode As ITable
        Dim pFilt As IQueryFilter
        Dim pCs As ICursor, pRow As IRow, pRowMode As IRow
        Dim boolPrevProj As Boolean
        boolPrevProj = False

        Dim clsExistingPrjAtts As clsModeAttributes

        lEdgeID = pFtEdge.Value(pFtEdge.Fields.FindField(g_PSRCEdgeID))

        '[100207] hyu: use the projRteID instead of projID_Ver
        sPrjID = pFtPrj.Value(pFtPrj.Fields.FindField("projRteID")) '"ProjID_Ver"))

        pTblAtt = get_TableClass(m_layers(24))
        pFilt = New QueryFilter
        'pFilt.WhereClause = "Scenario_ID=" & ScenarioID & " AND " & g_PSRCEdgeID & "=" & lEdgeID

        pFilt.WhereClause = g_PSRCEdgeID & "=" & lEdgeID
        pCs = pTblAtt.Search(pFilt, False)
        pRow = pCs.NextRow

        'if the edgeID is not found in the project edge attributes table, create a new row and copy all the attributes from the ModeAttributes table
        If pRow Is Nothing Then
            pRow = pTblAtt.CreateRow
            pFilt.WhereClause = g_PSRCEdgeID & "=" & lEdgeID
            pTblMode = get_TableClass(m_layers(2))
            pCs = pTblMode.Search(pFilt, False)
            pRowMode = pCs.NextRow
            CopyAttributes(pRowMode, pRow, "")
            'else, need to get a pointer to current edge attributes to compare with current project
        Else
            clsExistingPrjAtts = New clsModeAttributes(pRow)
            boolPrevProj = True
        End If

        '== find the direction relationship ==
        Dim dir As Integer
        dir = projectAndEdgeDirectionRelationship(pFtPrj, pFtEdge, dirPrj)
        '== get project attributes

        '== update attributes ==
        Dim i As Long
        Dim oneWay As Integer
        Dim fldName As String, fldName2 As String
        oneWay = pFtEdge.Value(pFtEdge.Fields.FindField(g_OneWay))
        '
        For i = 0 To pAttPrj.Fields.FieldCount - 1
            If pAttPrj.Fields.Field(i).Type <> esriFieldType.esriFieldTypeOID Then
                fldName = pAttPrj.Fields.Field(i).Name
                fldName2 = ""
                If pRow.Fields.FindField(fldName) > -1 Then
                    '                If UCase(Left(fldName, 2)) = "IJ" Or UCase(Left(fldName, 2)) = "JI" Then
                    If TypeOf pAttPrj.Value(pAttPrj.Fields.FindField(fldName)) Is Integer Then
                        If pAttPrj.Value(pAttPrj.Fields.FindField(fldName)) <> -1 Then

                            If dir = 1 Then 'with the direction of the project route's IJ
                                'If oneWay = 0 Then 'OnewayIJ
                                'If Left(fldName, 2) = "IJ" Then
                                'fldName2 = fldName
                                'Else  'IJ doen't apply here
                                'End If
                                'ElseIf oneWay = 1 Then 'OnewayJI
                                'If Left(fldName, 2) = "JI" Then
                                'fldName2 = fldName
                                'Else  'IJ doen't apply here
                                ' End If
                                ' Else  'Twoway
                                fldName2 = fldName
                                'End If

                                If fldName2 <> "" And pRow.Fields.FindField(fldName2) <> -1 And boolPrevProj = True Then
                                    If pRow.Fields.Field(pRow.Fields.FindField(fldName2)).Type = esriFieldType.esriFieldTypeDouble _
                                        Or pRow.Fields.Field(pRow.Fields.FindField(fldName2)).Type = esriFieldType.esriFieldTypeInteger Or pRow.Fields.Field(pRow.Fields.FindField(fldName2)).Type = esriFieldType.esriFieldTypeSingle Then
                                        If IsDBNull(clsExistingPrjAtts.FindField(fldName2)) = False Then
                                            If pAttPrj.Value(i) > clsExistingPrjAtts.FindField(fldName2) Then
                                                pRow.Value(pRow.Fields.FindField(fldName2)) = pAttPrj.Value(i)
                                            End If
                                        End If
                                    End If
                                    'pRow.value(pRow.Fields.FindField(fldName2)) = clsExistingPrjAtts.FindField(fldName2)

                                    'first entry
                                ElseIf fldName2 <> "" And pRow.Fields.FindField(fldName2) <> -1 Then
                                    pRow.Value(pRow.Fields.FindField(fldName2)) = pAttPrj.Value(i)
                                End If






                            Else    'dir=-1, against the direciton of the project route's IJ

                                If oneWay = 0 Then 'OnewayIJ
                                    fldName2 = fldName
                                    '   If Left(fldName, 2) = "JI" Then
                                    'fldName2 = "IJ" & Right(fldName, Len(fldName) - 2)
                                    '  Else 'IJ doesn't apply here
                                    ' End If
                                ElseIf oneWay = 1 Then 'OnewayJI
                                    '   If Left(fldName, 2) = "IJ" Then
                                    'fldName2 = "JI" & Right(fldName, Len(fldName) - 2)
                                    '  Else 'JI doesn't apply here
                                    ' End If
                                    'Twoway
                                Else
                                    If Left(fldName, 2) = "IJ" Then
                                        fldName2 = "JI" & Right(fldName, Len(fldName) - 2)
                                    Else
                                        fldName2 = "IJ" & Right(fldName, Len(fldName) - 2)

                                    End If
                                End If
                                '[081208] sec: I added a check to make sure the field exists in pRow.
                                'If fldName2 <> "" And pRow.Fields.FindField(fldName2) <> -1 Then pRow.value(pRow.Fields.FindField(fldName2)) = pRow.value(pRow.Fields.FindField(fldName2))
                                'If fldName2 <> "" And pRow.Fields.FindField(fldName2) <> -1 And boolPrevProj = True And (pRow.Fields.field(fldName2).Type = esriFieldTypeDouble Or esriFieldType.esriFieldTypeInteger Or esriFieldTypeSingle) Then
                                If fldName2 <> "" And pRow.Fields.FindField(fldName2) <> -1 And boolPrevProj = True Then
                                    If pRow.Fields.Field(pRow.Fields.FindField(fldName2)).Type = esriFieldType.esriFieldTypeDouble _
                                        Or pRow.Fields.Field(pRow.Fields.FindField(fldName2)).Type = esriFieldType.esriFieldTypeInteger Or pRow.Fields.Field(pRow.Fields.FindField(fldName2)).Type = esriFieldType.esriFieldTypeSingle Then
                                        If pAttPrj.Value(i) > clsExistingPrjAtts.FindField(fldName2) Then
                                            pRow.Value(pRow.Fields.FindField(fldName2)) = pAttPrj.Value(i)
                                        End If
                                    End If
                                ElseIf fldName2 <> "" And pRow.Fields.FindField(fldName2) <> -1 Then
                                    pRow.Value(pRow.Fields.FindField(fldName2)) = pAttPrj.Value(i)
                                End If


                            End If
                        End If 'pAttPrj.value(pAttPrj.Fields.FindField(fldName)) <> -1

                        '                Else 'not (UCase(Left(fldName, 2)) = "IJ" Or UCase(Left(fldName, 2)) = "JI")
                        '                    pRow.value(pRow.Fields.FindField(fldName)) = pAttPrj.value(i)
                        '                End If
                    End If
                End If
            End If
        Next i
        pRow.Store()
    End Function

    Public Function getProjectAtt(ByVal pFtPrj As IFeature) As IRow
        Dim pERelCls As IEnumRelationshipClass
        Dim pRelCls As IRelationshipClass2
        Dim pFCls As IFeatureClass

        pFCls = pFtPrj.Class
        pERelCls = pFCls.RelationshipClasses(esriRelRole.esriRelRoleOrigin)
        pRelCls = pERelCls.Next
        Do Until pRelCls.DestinationClass.AliasName = m_layers(3)
            pRelCls = pERelCls.Next
        Loop

        Dim piSet As ISet



        piSet = pRelCls.GetObjectsRelatedToObject(pFtPrj)
        getProjectAtt = piSet.Next
        pFCls = Nothing
        pRelCls = Nothing
        pERelCls = Nothing
    End Function

    Public Function updateProjectJunctionAttributes(ByVal pPrjFeat As IFeature) As Boolean
        '[100207] hyu
        'This procedure is called when ONLY point events is indicated in the selected project route.
        'The P_RStalls value from the evtPointProjectOutcomes is updated to the Junction in the ScenarioJunct layer

        Dim pRelCls As IRelationshipClass
        Dim pEnumRelCls As IEnumRelationshipClass
        Dim pFCls As IFeatureClass
        pFCls = pPrjFeat.Class
        pEnumRelCls = pFCls.RelationshipClasses(esriRelRole.esriRelRoleOrigin)
        pRelCls = pEnumRelCls.Next
        Do Until pRelCls Is Nothing
            If pRelCls.DestinationClass.AliasName = m_layers(6) Then 'evtPointProjectOutcomes
                Exit Do
            End If
            pRelCls = pEnumRelCls.Next
        Loop

        If pRelCls Is Nothing Then Exit Function

        Dim piSet As ISet
        piSet = pRelCls.GetObjectsRelatedToObject(pPrjFeat)

        Dim pRow As IRow
        Dim iPRStall As Long, lJunctID As Long
        Dim pCs As ICursor, pFilt As IQueryFilter2, pJctRow As IRow

        pFilt = New QueryFilter
        pRow = piSet.Next
        Do Until pRow Is Nothing
            If IsDBNull(pRow.Value(pRow.Fields.FindField("P_RStalls"))) Then
                WriteLogLine("P_RStalls is Null in " & m_layers(6))
            ElseIf IsDBNull(pRow.Fields.FindField("PSRCJunctID")) Then
                WriteLogLine("P_RStalls with PSRCJunctID is Null in " & m_layers(6))
            ElseIf (pRow.Fields.FindField("P_RStalls")) = 0 Then
                WriteLogLine("P_RStalls with PSRCJunctID = 0 in " & m_layers(6))
            Else
                'Now update the PRSTALL value of the junction in the scenarioJunct layer
                iPRStall = pRow.Value(pRow.Fields.FindField("P_RStalls"))
                lJunctID = pRow.Value(pRow.Fields.FindField("PSRCJunctID"))

                pFilt.WhereClause = g_PSRCJctID & "=" & lJunctID
                pCs = m_junctShp.Search(pFilt, False)
                pJctRow = pCs.NextRow
                pJctRow.Value(pJctRow.Fields.FindField("P_RStalls")) = iPRStall
                pJctRow.Value(pJctRow.Fields.FindField("JunctionType")) = 7
                pJctRow.Store()
            End If
            pRow = piSet.Next
        Loop

        pJctRow = Nothing
        pCs = Nothing
        piSet = Nothing
        pRelCls = Nothing
        pEnumRelCls = Nothing
    End Function

    Public Function projectAndEdgeDirectionRelationship(ByVal pFtPrj As IFeature, ByVal pFtEdge As IFeature, ByRef dirPrj As Integer) As Integer
        '== find the direction relationship ==
        Dim iNode As Long, JNode As Long
        Dim pPl As IPolyline4
        Dim pFt As IFeature, pFCS As IFeatureCursor, pFilt As IQueryFilter
        Dim pPt As IPoint, pPtFr As IPoint, pPtTo As IPoint
        Dim pRelOp As IRelationalOperator
        Dim pJctFCls As IFeatureClass

        pJctFCls = g_FWS.OpenFeatureClass(m_layers(1))

        ' pJctFCls = m_junctShp
        ' pJctFCls = pFeatLayerJ.FeatureClass
        pFilt = New QueryFilter

        'figure out the relationship between the project route's IJ direction and digitizing direction
        If dirPrj = 0 Then
            iNode = pFtPrj.Value(pFtPrj.Fields.FindField(g_INode))
            JNode = pFtPrj.Value(pFtPrj.Fields.FindField(g_JNode))
            If Not (iNode = 0 Or JNode = 0) Then
                pPl = pFtPrj.Shape

                pFilt.WhereClause = g_PSRCJctID & "=" & iNode
                ' pFCS = m_junctShp.Search(pFilt, False)
                pFCS = pJctFCls.Search(pFilt, False)
                pFt = pFCS.NextFeature

                pPt = pFt.Shape
                pPtFr = pPl.FromPoint
                pPtTo = pPl.ToPoint

                pRelOp = pPt
                If pRelOp.Equals(pPtFr) Then        'INode coincident the FROM point
                    dirPrj = 1
                ElseIf pRelOp.Equals(pPtTo) Then    'INode coincident the TO point
                    dirPrj = -1
                ElseIf (pPt.X - pPtFr.X) ^ 2 + (pPt.Y - pPtFr.Y) ^ 2 < (pPt.X - pPtTo.X) ^ 2 + (pPt.Y - pPtTo.Y) ^ 2 Then
                    'INODE is closer to FROM point, treat it as coincident
                    dirPrj = 1
                Else    'INODE is closer to TO point, treat it as coincident
                    dirPrj = -1
                End If
            Else
                'one or both of INode and JNode is 0, treat the IJ direction as the same of digitizing direction
                dirPrj = 1
            End If
        End If

        'relationship between the edge's IJ direction and digitizing direction
        Dim DirEdge As Integer
        '    If pFtEdge.value(pFtEdge.Fields.FindField("Flag")) = 0 Then

        iNode = pFtEdge.Value(pFtEdge.Fields.FindField(g_INode))
        JNode = pFtEdge.Value(pFtEdge.Fields.FindField(g_JNode))
        pPl = pFtEdge.Shape
        pFilt.WhereClause = g_PSRCJctID & "=" & iNode
        pFCS = m_junctShp.Search(pFilt, False)
        pFt = pFCS.NextFeature
        pPt = pFt.Shape
        pPtFr = pPl.FromPoint
        pPtTo = pPl.ToPoint

        pRelOp = pPt
        If pRelOp.Equals(pPtFr) Then        'INode coincident the FROM point
            DirEdge = 1
        ElseIf pRelOp.Equals(pPtTo) Then    'INode coincident the TO point
            DirEdge = -1
        ElseIf (pPt.X - pPtFr.X) ^ 2 + (pPt.Y - pPtFr.Y) ^ 2 < (pPt.X - pPtTo.X) ^ 2 + (pPt.Y - pPtTo.Y) ^ 2 Then
            'INODE is closer to FROM point, treat it as coincident
            DirEdge = 1
        Else    'INODE is closer to TO point, treat it as coincident
            DirEdge = -1
        End If
        '    Else
        '        DirEdge = pFtEdge.value(pFtEdge.Fields.FindField("Flag"))
        '    End If

        'final relationship between the project route's IJ direction and the edge's IJ direction
        pPl = pFtPrj.Shape
        Dim pHT As IHitTest
        Dim pPtCol As IPointCollection4
        Dim pPt2 As IPoint, lPart As Long, lSegFr As Long, lSegTo As Long, bright As Boolean, dDist As Double
        Dim dir As Integer

        pHT = pPl
        pPtCol = pPl
        'the following sees if the edges Inode hits a vertex on the project line that is after the vertex hit by it'sInode.

        pHT.HitTest(pPtFr, 100000, esriGeometryHitPartType.esriGeometryPartVertex, pPt2, dDist, lPart, lSegFr, bright)
        pHT.HitTest(pPtTo, 100000, esriGeometryHitPartType.esriGeometryPartVertex, pPt2, dDist, lPart, lSegTo, bright)

        If lSegFr < lSegTo Then
            dir = dirPrj * DirEdge
        Else
            dir = -(dirPrj * DirEdge)
        End If

        projectAndEdgeDirectionRelationship = dir
    End Function
   
    Public Function GetDwellTime(ByVal TransitPointFC As IFeatureClass, ByVal LineID As Long, ByVal PSRCJunctID As Long) As String

        Dim pFCursor As IFeatureCursor
        Dim pFilter As IQueryFilter
        Dim pFeature As IFeature
        pFilter = New QueryFilter
        pFilter.WhereClause = "LineID = " & LineID & " AND PSRCJunctID = " & PSRCJunctID
        pFCursor = TransitPointFC.Search(pFilter, False)
        pFeature = pFCursor.NextFeature
        If Not pFeature Is Nothing Then
            GetDwellTime = pFeature.Value(pFeature.Fields.FindField("txtDwt"))




        Else : GetDwellTime = "dwt=#.00"
        End If



    End Function
    Public Function GetDwellTime2(ByVal LineID As Long, ByVal intOrder As Integer) As String
        'Dim Pfltransitpoints As IFeatureLayer
        Pfltransitpoints = get_FeatureLayer("sde.SDE.TransitPoints")
        Dim pFCursor As IFeatureCursor
        Dim pFilter As IQueryFilter
        Dim pFeature As IFeature
        Dim indexTxtDWT As Integer
        Dim str As String

        indexTxtDWT = Pfltransitpoints.FeatureClass.FindField("txtDWT")
        pFilter = New QueryFilter
        pFilter.WhereClause = "LineID = " & LineID & " AND PointOrder = " & intOrder
        pFCursor = Pfltransitpoints.FeatureClass.Search(pFilter, False)
        pFeature = pFCursor.NextFeature
        If Not pFeature Is Nothing Then

            str = pFeature.Value(indexTxtDWT)




        Else
            str = "dwt=#.00"

        End If

        GetDwellTime2 = str
    End Function
    Public Function GetTransitPointsByLineID(ByVal dct As Dictionary(Of Object, Object), ByVal dct2 As Dictionary(Of Object, Object), ByVal TransitPointFC As IFeatureClass, ByVal LineID As Long)
        '[080309] SEC: stores all the transit points in a dictionary with a counter as the key. Could not use pointorder as
        'key because there are sequential transit nodes that are on the same node/junction. This was causing a problem and
        'are not included int the dictionary.

        Dim pFCursor As IFeatureCursor
        Dim pFilter As IQueryFilter
        Dim pFeature As IFeature
        'Dim pDictionary As Dictionary
        ' pDictionary = New Dictionary(of Long, Long)
        Dim indexOrder As Integer
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
                dct.Add(CType(Counter, Integer), pFeature.Value(CType(indexJunctID, Integer)))
                dct2.Add(CType(Counter, Integer), pFeature.Value(indexDwell))
                prevDwell = pFeature.Value(indexDwell)
                lngPrevJunctID = pFeature.Value(indexJunctID)
                Counter = Counter + 1
                pFeature = pFCursor.NextFeature
            End If

        Loop



        ' GetTransitPointsByLineID = pDictionary


    End Function
    Public Sub create_TransitFile3(ByVal pathnameN As String, ByVal filenameN As String)

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

            thepath = pathnameN + "\" + filenameN + "TranAM.txt"
            FileOpen(1, thepath, OpenMode.Output, OpenAccess.Default, OpenShare.Shared)
            'Open thepath For Output As 1

            thepath = pathnameN + "\" + filenameN + "TranMD.txt"
            FileOpen(2, thepath, OpenMode.Output, OpenAccess.Default, OpenShare.Shared)
            'Open thepath For Output As 2
            thepath = pathnameN + "\" + filenameN + "TranPM.txt"
            FileOpen(3, thepath, OpenMode.Output, OpenAccess.Default, OpenShare.Shared)
            'Open thepath For Output As 3
            thepath = pathnameN + "\" + filenameN + "TranE.txt"
            FileOpen(4, thepath, OpenMode.Output, OpenAccess.Default, OpenShare.Shared)
            'Open thepath For Output As 4
            thepath = pathnameN + "\" + filenameN + "TranN.txt"
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




            'sec 072609
            'Pfltransitpoints = get_FeatureLayer("sde.SDE.TransitPoints")

            pWS = get_Workspace()
            '    Set pTLineLayer = get_FeatureLayer(m_layers(13))  'transitlines
            Dim pTLineFCls As IFeatureClass
            pTLineFCls = getInterimTransitLines()
            tblTSeg = get_TableClass(m_layers(14))  'tbltransitsegments
            pModeAtts = get_TableClass(m_layers(2))
            Dim pMARow As IRow
            Dim pEdgeID As Long
            Dim pEdgeIDFilter As IQueryFilter
            Dim clsModeAtts As New clsModeAttributes(pMARow)
            Dim pMACursor As ICursor



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
            Dim dctTransitLine As Dictionary(Of Object, Object)
            dctTransitLine = New Dictionary(Of Object, Object)
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
            Dim pTransitMode As String
            Dim intOperator As Integer
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
                index = pFeatTRoute.Fields.FindField("UL2")
                If Not IsDBNull(pFeatTRoute.Value(index)) Then
                    tempString = tempString + " 0" + " " + CStr(pFeatTRoute.Value(index))
                Else
                    tempString = tempString + " 0" + " 1"
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
            Dim dctNodes As Dictionary(Of Object, Object)
            dctNodes = New Dictionary(Of Object, Object)
            getAllNodes(dctNodes)
            ' pStatusBar.ShowProgressBar("Creating Transit...", 0, tblTSeg.rowcount(pQF), 1, True)
            Dim intTPCounter As Integer
            Dim dctTransitPoints As Dictionary(Of Object, Object)
            Dim dctDwellTimes As Dictionary(Of Object, Object)
            Dim dctStopDistance As Dictionary(Of Object, Object)
            'Dim x As Long
            x = 0
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
                    dctTransitPoints = New Dictionary(Of Object, Object)
                    dctDwellTimes = New Dictionary(Of Object, Object)
                    dctStopDistance = New Dictionary(Of Object, Object)
                    GetTransitPointsByLineID(dctTransitPoints, dctDwellTimes, Pfltransitpoints.FeatureClass, lLineID)
                    intTPCounter = 1
                    x = pRow.Value(fldLineId)
                    GetStopDistances(lLineID, dctStopDistance, tblTSeg, dctTransitPoints)
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
                'see if the current node is a transit point
                If dctTransitPoints.Item(intTPCounter) = pRow.Value(fldINode) Then
                    lngStopDistance = dctStopDistance.Item(intTPCounter) - dctStopDistance.Item(intTPCounter + 1)
                    strDwell = dctDwellTimes.Item(intTPCounter)
                    intPrevTN = intTPCounter
                    intTPCounter = intTPCounter + 1
                End If

                If dctTransitPoints.Item(intPrevTN + 1) = pRow.Value(fldJNode) Then
                    sDwtStop = " " & strDwell
                Else
                    sDwtStop = " dwt=#.00"
                End If

                'sDwtStop = " " & GetDwellTime2(lLineID, intTPCounter)
                'if so, grab the DWT from the DWT dictionary




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

                    Else    'dctNodes.ContainsKey(CStr(curNode))
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

                            If Not pFeature Is Nothing Then
                                'Do TTF stuff here
                                If intOperator = 4 Then
                                    'do nothing
                                Else
                                    'pTransitMode = pFeature.Fields.FindField("NewFacilityType")
                                    lngStopDistance = Math.Abs(lngStopDistance)
                                    Dim strSegmentMode As String
                                    strSegmentMode = pFeature.Value(pFeature.Fields.FindField("Modes"))
                                    If pTransitMode = "r" Then
                                        stimeFuncID = 4
                                    ElseIf pTransitMode = "f" Then
                                        stimeFuncID = 5
                                    ElseIf strSegmentMode = "br" Or strSegmentMode = "bwk" Or strSegmentMode = "b" Or strSegmentMode = "wkb" Then
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
                                                pEdgeIDFilter = New QueryFilter
                                                pEdgeIDFilter.WhereClause = "PSRCEdgeID = " & pEdgeID
                                                pMACursor = pModeAtts.Search(pEdgeIDFilter, True)

                                                pMARow = pMACursor.NextRow
                                                System.Runtime.InteropServices.Marshal.ReleaseComObject(pMACursor)
                                                '111309- Previous code assumed that an edge flagged HOV in ScenarioEdge is HOV for both directions at all times of day
                                                'the following makes sure the at the edge is HOV for TOD and direction
                                                If IsHOV(pMARow, Left(CStr(lLineID), 1), 1) Then
                                                    ' clsModeAtts.modeAttributeRow = pMARow
                                                    'MsgBox (clsModeAtts.IJLANESHOVAM)
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
                                                pEdgeIDFilter = New QueryFilter
                                                pEdgeIDFilter.WhereClause = "PSRCEdgeID = " & pEdgeID
                                                pMACursor = pModeAtts.Search(pEdgeIDFilter, True)
                                                pMARow = pMACursor.NextRow
                                                System.Runtime.InteropServices.Marshal.ReleaseComObject(pMACursor)
                                                If IsHOV(pMARow, Left(CStr(lLineID), 1), 2) Then
                                                    curWNode = IIf(IsDBNull(pFeature.Value(m_edgeShp.FindField("HOV_J"))), 0, pFeature.Value(m_edgeShp.FindField("HOV_J")) + m_Offset)
                                                    nextWNode = IIf(IsDBNull(pFeature.Value(m_edgeShp.FindField("HOV_I"))), 0, pFeature.Value(m_edgeShp.FindField("HOV_I")) + m_Offset)
                                                End If
                                                System.Runtime.InteropServices.Marshal.ReleaseComObject(pMACursor)
                                        End Select

                                        If fVerboseLog Then WriteLogLine("potential weave nodes (HOV_I/J) " & preWNode & ", " & curWNode)
                                    End If
                                End If 'pfeature.value(m_edgeShp.FindField("INode")) = curNode



                                If bPreWeave And preWNode = curWNode And curWNode <> 0 Then
                                    writeTransitNode(lTOD, " " + CStr(nextWNode) + tempString)
                                    bPreWeave = True
                                Else
                                    writeTransitNode(lTOD, " " + CStr(dctNodes.Item(CStr(curNode))) + tempString)
                                    If curWNode > 0 And nextWNode > 0 Then
                                        writeTransitNode(lTOD, " " + CStr(curWNode) + tempString)
                                        writeTransitNode(lTOD, " " + CStr(nextWNode) + tempString)
                                        bPreWeave = True
                                    Else
                                        bPreWeave = False
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
                            writeTransitNode(lTOD, " " + CStr(dctNodes.Item(CStr(curNode))) + tempString)
                            bPreWeave = False

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
                    writeTransitNode(lTOD, lastNodeString)
                Else
                    If pRow.Value(fldLineId) <> lLineID Then
                        'a new transit line starts. so the unwritten last node of the previous line should be written now
                        writeTransitNode(lTOD, lastNodeString)
                    End If
                End If
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

    Public Sub create_TransitFile4(ByVal pathnameN As String, ByVal filenameN As String)

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

        thepath = pathnameN + "\" + filenameN + "TranAM.txt"
        FileOpen(1, thepath, OpenMode.Output, OpenAccess.Default, OpenShare.Shared)
        'Open thepath For Output As 1

        thepath = pathnameN + "\" + filenameN + "TranMD.txt"
        FileOpen(2, thepath, OpenMode.Output, OpenAccess.Default, OpenShare.Shared)
        'Open thepath For Output As 2
        thepath = pathnameN + "\" + filenameN + "TranPM.txt"
        FileOpen(3, thepath, OpenMode.Output, OpenAccess.Default, OpenShare.Shared)
        'Open thepath For Output As 3
        thepath = pathnameN + "\" + filenameN + "TranE.txt"
        FileOpen(4, thepath, OpenMode.Output, OpenAccess.Default, OpenShare.Shared)
        'Open thepath For Output As 4
        thepath = pathnameN + "\" + filenameN + "TranN.txt"
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




        'sec 072609
        'Pfltransitpoints = get_FeatureLayer("sde.SDE.TransitPoints")

        pWS = get_Workspace()
        '    Set pTLineLayer = get_FeatureLayer(m_layers(13))  'transitlines
        Dim pTLineFCls As IFeatureClass
        pTLineFCls = getInterimTransitLines()
        tblTSeg = get_TableClass(m_layers(14))  'tbltransitsegments
        pModeAtts = get_TableClass(m_layers(2))
        Dim pMARow As IRow
        Dim pEdgeID As Long
        Dim pEdgeIDFilter As IQueryFilter
        Dim clsModeAtts As New clsModeAttributes(pMARow)
        Dim pMACursor As ICursor



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
        Dim dctTransitLine As Dictionary(Of Object, Object)
        dctTransitLine = New Dictionary(Of Object, Object)
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
        Dim pTransitMode As String
        Dim intOperator As Integer
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
            index = pFeatTRoute.Fields.FindField("UL2")
            If Not IsDBNull(pFeatTRoute.Value(index)) Then
                tempString = tempString + " 0" + " " + CStr(pFeatTRoute.Value(index))
            Else
                tempString = tempString + " 0" + " 1"
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
        Dim dctNodes As Dictionary(Of Object, Object)
        dctNodes = New Dictionary(Of Object, Object)
        getAllNodes(dctNodes)
        ' pStatusBar.ShowProgressBar("Creating Transit...", 0, tblTSeg.rowcount(pQF), 1, True)
        Dim intTPCounter As Integer
        Dim dctTransitPoints As Dictionary(Of Object, Object)
        Dim dctDwellTimes As Dictionary(Of Object, Object)
        Dim dctStopDistance As Dictionary(Of Object, Object)
        'Dim x As Long
        x = 0
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
                dctTransitPoints = New Dictionary(Of Object, Object)
                dctDwellTimes = New Dictionary(Of Object, Object)
                dctStopDistance = New Dictionary(Of Object, Object)
                GetTransitPointsByLineID(dctTransitPoints, dctDwellTimes, Pfltransitpoints.FeatureClass, lLineID)
                intTPCounter = 1
                x = pRow.Value(fldLineId)
                GetStopDistances(lLineID, dctStopDistance, tblTSeg, dctTransitPoints)
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
            'see if the current node is a transit point
            If dctTransitPoints.Item(intTPCounter) = pRow.Value(fldINode) Then
                lngStopDistance = dctStopDistance.Item(intTPCounter) - dctStopDistance.Item(intTPCounter + 1)
                strDwell = dctDwellTimes.Item(intTPCounter)
                intPrevTN = intTPCounter
                intTPCounter = intTPCounter + 1
            End If

            If dctTransitPoints.Item(intPrevTN + 1) = pRow.Value(fldJNode) Then
                sDwtStop = " " & strDwell
            Else
                sDwtStop = " dwt=#.00"
            End If

            'sDwtStop = " " & GetDwellTime2(lLineID, intTPCounter)
            'if so, grab the DWT from the DWT dictionary




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

                Else    'dctNodes.ContainsKey(CStr(curNode))
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

                        If Not pFeature Is Nothing Then
                            'Do TTF stuff here
                            If intOperator = 4 Then
                                'do nothing
                            Else
                                'pTransitMode = pFeature.Fields.FindField("NewFacilityType")
                                lngStopDistance = Math.Abs(lngStopDistance)
                                Dim strSegmentMode As String
                                strSegmentMode = pFeature.Value(pFeature.Fields.FindField("Modes"))
                                If pTransitMode = "r" Then
                                    stimeFuncID = 4
                                ElseIf pTransitMode = "f" Then
                                    stimeFuncID = 5
                                ElseIf strSegmentMode = "br" Or strSegmentMode = "bwk" Or strSegmentMode = "b" Or strSegmentMode = "wkb" Then
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
                            If pFeature.Value(m_edgeShp.FindField("INode")) = curNode Then 'JI
                                If pFeature.Value(m_edgeShp.FindField("TR_I")) > 0 Then
                                    'weave nodes from "TR_I" and "TR_J"
                                    curWNode = pFeature.Value(m_edgeShp.FindField("TR_I")) + m_Offset
                                    nextWNode = pFeature.Value(m_edgeShp.FindField("TR_J")) + m_Offset
                                    If fVerboseLog Then WriteLogLine("potential weave nodes (TR_J/I) " & preWNode & ", " & curWNode)
                                Else
                                    'weave nodes may from "HOV_I/J"
                                    Select Case pFeature.Value(m_edgeShp.FindField("FacilityType"))
                                        Case 11, 12, 3, 13, 9
                                            curWNode = 0
                                            nextWNode = 0
                                        Case Else
                                            pEdgeID = pFeature.Value(m_edgeShp.FindField("PSRCEdgeID"))
                                            pEdgeIDFilter = New QueryFilter
                                            pEdgeIDFilter.WhereClause = "PSRCEdgeID = " & pEdgeID
                                            pMACursor = pModeAtts.Search(pEdgeIDFilter, True)
                                            pMARow = pMACursor.NextRow
                                            '111309- Previous code assumed that an edge flagged HOV in ScenarioEdge is HOV for both directions at all times of day
                                            'the following makes sure the at the edge is HOV for TOD and direction
                                            If IsHOV(pMARow, Left(CStr(lLineID), 1), 1) Then
                                                'Set clsModeAtts.modeAttributeRow = pMARow
                                                'MsgBox (clsModeAtts.IJLANESHOVAM)
                                                curWNode = IIf(IsDBNull(pFeature.Value(m_edgeShp.FindField("HOV_I"))), 0, pFeature.Value(m_edgeShp.FindField("HOV_I")) + m_Offset)
                                                nextWNode = IIf(IsDBNull(pFeature.Value(m_edgeShp.FindField("HOV_J"))), 0, pFeature.Value(m_edgeShp.FindField("HOV_J")) + m_Offset)
                                            End If
                                    End Select

                                    If fVerboseLog Then WriteLogLine("potential weave nodes (HOV_J/I) " & curWNode & ", " & nextWNode)
                                End If

                            Else 'Transit segment goes in the JI direction
                                If pFeature.Value(m_edgeShp.FindField("TR_J")) > 0 Then
                                    'weave nodes from "TR_I" and "TR_J"
                                    curWNode = pFeature.Value(m_edgeShp.FindField("TR_J")) + m_Offset
                                    nextWNode = pFeature.Value(m_edgeShp.FindField("TR_I")) + m_Offset

                                    If fVerboseLog Then WriteLogLine("potential weave nodes (TR_I/J) " & curWNode & ", " & nextWNode)
                                Else
                                    'weave nodes may from "HOV_I/J"
                                    Select Case pFeature.Value(m_edgeShp.FindField("FacilityType"))
                                        Case 11, 12, 3, 13, 9
                                            curWNode = 0
                                            nextWNode = 0
                                        Case Else
                                            pEdgeID = pFeature.Value(m_edgeShp.FindField("PSRCEdgeID"))
                                            pEdgeIDFilter = New QueryFilter
                                            pEdgeIDFilter.WhereClause = "PSRCEdgeID = " & pEdgeID
                                            pMACursor = pModeAtts.Search(pEdgeIDFilter, True)
                                            pMARow = pMACursor.NextRow
                                            If IsHOV(pMARow, Left(CStr(lLineID), 1), 2) Then
                                                curWNode = IIf(IsDBNull(pFeature.Value(m_edgeShp.FindField("HOV_J"))), 0, pFeature.Value(m_edgeShp.FindField("HOV_J")) + m_Offset)
                                                nextWNode = IIf(IsDBNull(pFeature.Value(m_edgeShp.FindField("HOV_I"))), 0, pFeature.Value(m_edgeShp.FindField("HOV_I")) + m_Offset)
                                            End If
                                    End Select

                                    If fVerboseLog Then WriteLogLine("potential weave nodes (HOV_I/J) " & preWNode & ", " & curWNode)
                                End If
                            End If 'pfeature.value(m_edgeShp.FindField("INode")) = curNode

                            If bPreWeave And preWNode = curWNode And curWNode <> 0 Then
                                writeTransitNode(lTOD, " " + CStr(nextWNode) + tempString)
                                bPreWeave = True
                            Else
                                writeTransitNode(lTOD, " " + CStr(dctNodes.Item(CStr(curNode))) + tempString)
                                If curWNode > 0 And nextWNode > 0 Then
                                    writeTransitNode(lTOD, " " + CStr(curWNode) + tempString)
                                    writeTransitNode(lTOD, " " + CStr(nextWNode) + tempString)
                                    bPreWeave = True
                                Else
                                    bPreWeave = False
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
                        writeTransitNode(lTOD, " " + CStr(dctNodes.Item(CStr(curNode))) + tempString)
                        bPreWeave = False

                        preNode = curNode
                        curNode = nextNode
                        nextNode = 0
                        curWNode = 0
                        preWNode = 0
                        nextWNode = 0
                    End If  'lUseGP = 0
                End If  'Not dctNodes.ContainsKey(CStr(curNode))
            End If  'curNode <> preNode
            pRow = pTC.NextRow

            If pRow Is Nothing Then


                'now is at the end, write the last node
                writeTransitNode(lTOD, lastNodeString)
            Else
                If pRow.Value(fldLineId) <> lLineID Then
                    'a new transit line starts. so the unwritten last node of the previous line should be written now
                    writeTransitNode(lTOD, lastNodeString)
                End If
            End If
        Loop

        'pStatusBar.HideProgressBar()
        FileClose(1)
        FileClose(2)
        FileClose(3)
        FileClose(4)
        FileClose(5)


        pFeatLayerE = Nothing
        '  Set pTLineLayer = Nothing
        pTLineFCls = Nothing
        pFeature = Nothing
        WriteLogLine("FINISHED create_TransitFile at " & Now())
        dctTransitLine.Clear()
        dctNodes.Clear()
        dctTransitLine = Nothing
        dctNodes = Nothing
        Exit Sub
eh:
        CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.create_TransitFile")
        MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.create_TransitFile")
        'Print #9, Err.Description, vbExclamation, "TransitFile"
        FileClose(1)
        FileClose(2)
        FileClose(3)
        FileClose(4)
        FileClose(5)
        'Close #9
        pFeatLayerE = Nothing
        '  Set pTLineLayer = Nothing
        pTLineFCls = Nothing
        pFeature = Nothing
        dctTransitLine = Nothing
        dctNodes = Nothing

    End Sub
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


        Try



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

            thepath = pathnameN + "\" + filenameN + "TranAM.txt"
            FileOpen(1, thepath, OpenMode.Output, OpenAccess.Default, OpenShare.Shared)
            'Open thepath For Output As 1

            thepath = pathnameN + "\" + filenameN + "TranMD.txt"
            FileOpen(2, thepath, OpenMode.Output, OpenAccess.Default, OpenShare.Shared)
            'Open thepath For Output As 2
            thepath = pathnameN + "\" + filenameN + "TranPM.txt"
            FileOpen(3, thepath, OpenMode.Output, OpenAccess.Default, OpenShare.Shared)
            'Open thepath For Output As 3
            thepath = pathnameN + "\" + filenameN + "TranE.txt"
            FileOpen(4, thepath, OpenMode.Output, OpenAccess.Default, OpenShare.Shared)
            'Open thepath For Output As 4
            thepath = pathnameN + "\" + filenameN + "TranN.txt"
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


            'sec 072609
            'Pfltransitpoints = get_FeatureLayer("sde.SDE.TransitPoints")

            pWS = get_Workspace()
            '     pTLineLayer = get_FeatureLayer(m_layers(13))  'transitlines
            Dim pTLineFCls As IFeatureClass
            pTLineFCls = getInterimTransitLines()
            tblTSeg = get_TableClass(m_layers(14))  'tbltransitsegments
            pModeAtts = get_TableClass(m_layers(2))
            Dim pMARow As IRow
            Dim pEdgeID As Long
            Dim pEdgeIDFilter As IQueryFilter
            Dim clsModeAtts As clsModeAttributes
            Dim pMACursor As ICursor



            Dim pQFtlines As IQueryFilter, pQF As IQueryFilter, pQF2 As IQueryFilter
            Dim pFeature As IFeature, pFeatTRoute As IFeature, pFeatE As IFeature
            Dim pFeatCursor As IFeatureCursor
            Dim pFCtroute As IFeatureCursor
            Dim pTC As ICursor
            Dim pRow As IRow
            Dim i As Long

            pQFtlines = New QueryFilter
            '    pQFtlines.WhereClause = "InServiceDate <= " + CStr(inserviceyear)
            '     pFCtroute = pTLineLayer.Search(pQFtlines, False)
            pFCtroute = pTLineFCls.Search(Nothing, False)
            pFeatTRoute = pFCtroute.NextFeature

            Dim index As Long
            Dim idIndex As Long, sIndex As Long
            '    Dim iNode As Long, jNode As Long

            'sort transit segment table by LineID and SegOrder, then loop through the sorted rows
            Dim psort As ITableSort
            Dim dctTransitLine As Dictionary(Of Object, Object)
            dctTransitLine = New Dictionary(Of Object, Object)
            psort = New TableSort
            pQF = New QueryFilter

            If strLayerPrefix = "SDE" Then
                pQF.WhereClause = "SUBSTRING(CAST(LineID AS CHARACTER), 2, 2) = '" & Right(CStr(inserviceyear), 2) & "'"
            Else
                pQF.WhereClause = "Mid([LineID], 2, 2) = '" & Right(CStr(inserviceyear), 2) & "'"
            End If

            With psort
                .Fields = "LineID, SegOrder"
                .Ascending("LineID") = True
                .Ascending("SegOrder") = True
                .QueryFilter = Nothing ' pQF
                .Table = tblTSeg
            End With
            psort.Sort(Nothing)
            pTC = psort.Rows

            'start statusbar
            Dim count As Long
            Dim sLineID As String
            '    count = pTLineLayer.FeatureClass.featurecount(pQFtlines)
            count = pTLineFCls.FeatureCount(Nothing)
            'pStatusBar.ShowProgressBar("Preparing for Transit...", 0, count, 1, True)
            'jaf--let's keep the user updated
            g_frmNetLayer.lblStatus.Text = "GlobalMod.create_Netfile: Starting Transit Build..."
            'pan frmNetLayer.Refresh

            'loop through searched transit Lines table, and get all transit info into a dictionary
            '    idIndex = pTLineLayer.FeatureClass.FindField("LineID")
            idIndex = pTLineFCls.FindField("LineID")
            sIndex = tblTSeg.FindField("SegOrder")
            Dim x As Long
            'Dim intTPCounter As Integer
            Do Until pFeatTRoute Is Nothing
                '[040307] hyu: per Jeff's email on [05/16/2006]
                'For each insertion of a new transit line in the buildfile, the application should
                'extract the first three digits of TransitLine.LineID to report as the time-of-day and year code
                'in a new comment line placed just above the buildfile line that inserts the transit line.
                'All remaining TransitLine.LineID digits to the right of the first three
                'should be written as the line ID in the route insertion line.




                sLineID = CStr(pFeatTRoute.Value(idIndex))



                If Len(sLineID) > 3 Then
                    tempString = "c '" + Left(sLineID, 2) + "'" + vbNewLine

                    'tempString = tempString + "a '" + Mid(sLineID, 3, Len(sLineID) - 3) + "'"
                    tempString = tempString + "a '" + sLineID.Remove(0, 2) + "'"
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
                'If Not IsDBNull(pFeatTRoute.value(index)) Then
                'tempString = tempString + " " + CStr(pFeatTRoute.value(index))
                'Else
                'tempString = tempString + " 1"
                'End If
                'index = pFeatTRoute.Fields.FindField("Oneway")
                'If Not IsDBNull(pFeatTRoute.value(index)) Then
                'tempString = tempString + " " + CStr(pFeatTRoute.value(index))
                'Else
                'tempString = tempString + " 1"
                'End If
                index = pFeatTRoute.Fields.FindField("UL2")
                If Not IsDBNull(pFeatTRoute.Value(index)) Then
                    tempString = tempString + " 0" + " " + CStr(pFeatTRoute.Value(index))
                Else
                    tempString = tempString + " 0" + " 1"
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
            Dim boolWeave As Boolean

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
            preLine = pRow.Value(pRow.Fields.FindField("LineID"))





            'get all nodes in the intermediate layer
            Dim dctNodes As Dictionary(Of Object, Object)
            dctNodes = New Dictionary(Of Object, Object)
            getAllNodes(dctNodes)
            'pStatusBar.ShowProgressBar("Creating Transit...", 0, tblTSeg.RowCount(pQF), 1, True)
            Dim intTPCounter As Integer
            Dim dctTransitPoints As Dictionary(Of Object, Object)
            Dim dctDwellTimes As Dictionary(Of Object, Object)
            'Dim x As Long
            x = 0
            Do Until pRow Is Nothing
                'pStatusBar.StepProgressBar()

                lLineID = pRow.Value(fldLineId)
                'sec 073009
                'get all the TransitPoints that belong to the route, store them in a dictionary
                'store their DWTs in another dictionary
                If x <> pRow.Value(fldLineId) Then
                    dctTransitPoints = New Dictionary(Of Object, Object)
                    dctDwellTimes = New Dictionary(Of Object, Object)
                    GetTransitPointsByLineID(dctTransitPoints, dctDwellTimes, Pfltransitpoints.FeatureClass, lLineID)
                    intTPCounter = 1
                    x = pRow.Value(fldLineId)
                End If
                Do Until dctTransitLine.ContainsKey(CStr(lLineID))
                    'if transit info doesn't exist, then skip all the records of this line.
                    pRow = pTC.NextRow
                    'pStatusBar.StepProgressBar()

                    If pRow Is Nothing Then
                        FileClose(1, 2, 3, 4, 5)
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
                'see if the current node is a transit point
                If dctTransitPoints.Item(intTPCounter) = pRow.Value(fldINode) Then
                    strDwell = dctDwellTimes.Item(intTPCounter)
                    intPrevTN = intTPCounter
                    intTPCounter = intTPCounter + 1
                End If

                If dctTransitPoints.Item(intPrevTN + 1) = pRow.Value(fldJNode) Then
                    sDwtStop = " " & strDwell
                Else
                    sDwtStop = " dwt=.00"
                End If

                'sDwtStop = " " & GetDwellTime2(lLineID, intTPCounter)
                'if so, grab the DWT from the DWT dictionary




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

                    Else    'dctNodes.ContainsKey(CStr(curNode))
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
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(pFeatCursor)

                            If Not pFeature Is Nothing Then

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
                                                pEdgeIDFilter = New QueryFilter
                                                pEdgeIDFilter.WhereClause = "PSRCEdgeID = " & pEdgeID
                                                pMACursor = pModeAtts.Search(pEdgeIDFilter, True)
                                                pMARow = pMACursor.NextRow
                                                '111309- Previous code assumed that an edge flagged HOV in ScenarioEdge is HOV for both directions at all times of day
                                                'the following makes sure the at the edge is HOV for TOD and direction
                                                If IsHOV(pMARow, Left(CStr(lLineID), 1), 1) Then
                                                    ' clsModeAtts.modeAttributeRow = pMARow
                                                    'MsgBox (clsModeAtts.IJLANESHOVAM)
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
                                                pEdgeIDFilter = New QueryFilter
                                                pEdgeIDFilter.WhereClause = "PSRCEdgeID = " & pEdgeID
                                                pMACursor = pModeAtts.Search(pEdgeIDFilter, True)
                                                pMARow = pMACursor.NextRow
                                                If IsHOV(pMARow, Left(CStr(lLineID), 1), 2) Then
                                                    curWNode = IIf(IsDBNull(pFeature.Value(m_edgeShp.FindField("HOV_J"))), 0, pFeature.Value(m_edgeShp.FindField("HOV_J")) + m_Offset)
                                                    nextWNode = IIf(IsDBNull(pFeature.Value(m_edgeShp.FindField("HOV_I"))), 0, pFeature.Value(m_edgeShp.FindField("HOV_I")) + m_Offset)
                                                End If
                                        End Select

                                        If fVerboseLog Then WriteLogLine("potential weave nodes (HOV_I/J) " & preWNode & ", " & curWNode)
                                    End If
                                End If 'pfeature.value(m_edgeShp.FindField("INode")) = curNode



                                If bPreWeave And preWNode = curWNode And curWNode <> 0 Then
                                    writeTransitNode(lTOD, " " + CStr(nextWNode) + tempString)
                                    bPreWeave = True
                                Else
                                    writeTransitNode(lTOD, " " + CStr(dctNodes.Item(CStr(curNode))) + tempString)
                                    If curWNode > 0 And nextWNode > 0 Then
                                        writeTransitNode(lTOD, " " + CStr(curWNode) + tempString)
                                        writeTransitNode(lTOD, " " + CStr(nextWNode) + tempString)
                                        bPreWeave = True
                                    Else
                                        bPreWeave = False
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
                            writeTransitNode(lTOD, " " + CStr(dctNodes.Item(CStr(curNode))) + tempString)
                            bPreWeave = False

                            preNode = curNode
                            curNode = nextNode
                            nextNode = 0
                            curWNode = 0
                            preWNode = 0
                            nextWNode = 0
                        End If  'lUseGP = 0
                    End If  'Not dctNodes.ContainsKey(CStr(curNode))
                End If  'curNode <> preNode
                pRow = pTC.NextRow

                If pRow Is Nothing Then


                    'now is at the end, write the last node
                    writeTransitNode(lTOD, lastNodeString)
                Else
                    If pRow.Value(fldLineId) <> lLineID Then
                        'a new transit line starts. so the unwritten last node of the previous line should be written now
                        writeTransitNode(lTOD, lastNodeString)
                    End If
                End If
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
    Public Function IsHOV(ByVal MARow As IRow, ByVal TOD As Integer, ByVal IJ As Integer) As Boolean
        'determines if the TransitSegment has HOV for the time and direction of the current Transit Line.
        Dim clsModeAtts As New clsModeAttributes(MARow)
        'clsModeAtts.modeAttributeRow = MARow
        'am
        If TOD = 1 Then
            If IJ = 1 Then
                If clsModeAtts.IJLANESHOVAM > 0 Then
                    IsHOV = True
                Else
                    IsHOV = False
                End If
            Else
                If clsModeAtts.JILANESHOVAM > 0 Then
                    IsHOV = True
                Else
                    IsHOV = False
                End If
            End If
            'md
        Else
            If IJ = 1 Then
                If clsModeAtts.IJLANESHOVMD > 0 Then
                    IsHOV = True
                Else
                    IsHOV = False
                End If
            Else
                If clsModeAtts.JILANESHOVMD > 0 Then
                    IsHOV = True
                Else
                    IsHOV = False
                End If
            End If
        End If



    End Function
    Public Function IsHOV2(ByVal EdgeID As Long, ByVal dctHOV As Dictionary(Of Long, Long)) As Boolean
        'determines if the TransitSegment has HOV for the time and direction of the current Transit Line.

        'clsModeAtts.modeAttributeRow = MARow
        'am


        If dctHOV.ContainsKey(EdgeID) Then
            IsHOV2 = True


        Else

            IsHOV2 = False
        End If




    End Function

    Public Function getPGDws(ByVal pMap As IMap) As IFeatureWorkspace
        'Dim pDoc As IMxDocument, 
        'Dim pMap As IMap
        'pDoc = ThisDocument
        'pMap = m_App.Document.


        Dim pFLy As IFeatureLayer
        Dim pDS As IDataset
        Dim pEL As IEnumLayer
        pEL = pMap.Layers
        Dim pLy As ILayer
        pLy = pEL.Next
        Do Until pLy Is Nothing
            If pLy.Valid Then
                If TypeOf pLy Is IFeatureLayer2 Then
                    pFLy = pLy
                    pDS = pFLy
                    getPGDws = pDS.Workspace
                    GoTo ExitFunction
                End If
            End If
            pLy = pEL.Next
        Loop

ExitFunction:
        pDS = Nothing
        pFLy = Nothing
        'pMap = Nothing
        'pDoc = Nothing
        pEL = Nothing
        pLy = Nothing

    End Function
    Public Function GetStopDistances(ByVal LineID As Long, ByVal dct As Dictionary(Of Object, Object), ByVal TransitSegments As ITable, ByVal dctTP As Dictionary(Of Object, Object))
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
        Dim x As Integer
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



            If intTPCounter = 1 Then
                dct.Add(CType(intTPCounter, Integer), 0)
                intTPCounter = intTPCounter + 1
                If JNode = dctTP.Item(CType(intTPCounter, Integer)) Then
                    lngDist = pRow.Value(indexDist)
                    dct.Add(CType(intTPCounter, Integer), CType(lngDist, Integer))
                    intTPCounter = intTPCounter + 1
                    'test = True
                Else
                    lngDist = pRow.Value(indexDist)
                End If

            ElseIf JNode = dctTP.Item(CType(intTPCounter, Integer)) Then



                lngDist = lngDist + pRow.Value(indexDist)
                dct.Add(CType(intTPCounter, Integer), CType(lngDist, Integer))
                intTPCounter = intTPCounter + 1


            Else
                lngDist = lngDist + pRow.Value(indexDist)


            End If







            pRow = pCursor.NextRow

        Loop





    End Function

End Module
