Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Carto
Imports Scripting
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.esriSystem
Imports Microsoft.SqlServer
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.CatalogUI
Imports ESRI.ArcGIS.Catalog
Imports ESRI.ArcGIS.DataSourcesFile
Imports ESRI.ArcGIS.DataSourcesGDB






Public Class frmBuildFiles
    Public m_application As IApplication
    Private iName As String
    Private pathName As String
    Public m_dctReservedNodes As Dictionary(Of Object, Object)
    Public m_Doc As IMxDocument
    Public m_Map As IMap
    Public m_ActiveView As IActiveView
    Public m_turnLayer As IFeatureLayer

    Public m_projectLayer As IFeatureLayer
    Public m_transitLayer As IFeatureLayer



    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOpenFile.Click
        On Error GoTo eh

        Dim pGxDialog As IGxDialog
        Dim pGxObjectFilter As IGxObjectFilter
        Dim pEnumGxObject As IEnumGxObject
        Dim pGxObject As IGxObject
        Dim pgxfile As IGxFile
        Dim bSelected As Boolean
        pGxObjectFilter = New GxFilterShapefiles
        pGxDialog = New GxDialog
        pGxDialog.ObjectFilter = pGxObjectFilter
        pGxDialog.Title = "Select the intermediate edge shapefile"
        bSelected = pGxDialog.DoModalSave(0)

        If bSelected Then
            pathName = pGxDialog.FinalLocation.FullName
            iName = pGxDialog.Name
        Else
            pGxDialog = Nothing
            Exit Sub
        End If

        pGxDialog = Nothing

        txtFile.Text = pathName & "\" & iName
        Dim len1 As Integer
        len1 = Len(iName)
        iName.Remove(len1 - 4, 4)
        'iName = trimstart((iName, (len1 - 4))

        g_sOutPath = pathName
        g_sOutName = iName
        Exit Sub

eh:



        SaveFileDialog1.ShowDialog()
        Me.txtFile.Text = SaveFileDialog1.FileName
    End Sub

    Private Sub btnContinue_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnContinue.Click
        InitiateMainVariables()
        'SetMainVariables()
        g_FWS = getPGDws(m_Map)
        ' End If
        init(m_application, txtSchema.Text)
        OpenLogFile("C:", "LogBuildFile2.txt")

        'File Names:
        Dim roadNetworkName As String = "roadway.in"
        Dim transitName As String = "transit.in"
        Dim turnsName As String = "turns.in"
        Dim parkAndRideName As String = "pnr_lot_capacities.in"

        Dim strTblModeTolls As String
        strTblModeTolls = m_layers(9)
        g_TblMode = g_FWS.OpenTable(strTblModeTolls)

        gTblModeAtt = g_FWS.OpenTable(m_layers(2))
        
        m_Offset = CType(Me.txtOffset.Text, Long)
        GlobalMod.inserviceyear = txtYear.Text

        If (Me.txtFile.Text = "") Then
            Me.txtFile.Text = SaveFileDialog1.FileName
            'MsgBox "Please enter a filename"
            Exit Sub
        End If

        Dim strHolder As String
        Dim Leng As Long, tempname As String

        strHolder = Me.txtFile.Text
        Leng = Len(iName)

        If (pathName = "") Then Exit Sub

        GlobalMod.pWorkspaceI = g_FWS

        Dim pFeatlayerIE As IFeatureLayer
        Dim pFeatSelect As IFeatureSelection
        pFeatlayerIE = New FeatureLayer

        'now open the new intermediate edge shp
        m_edgeShp = g_FWS.OpenFeatureClass(m_layers(22)) 'ScenarioEdge
        m_TransRefEdges = g_FWS.OpenFeatureClass(m_layers(0))
        Dim pSubType As ISubtypes
        pSubType = m_TransRefEdges

        '  get  the enumeration of all of the subtypes for this feature class 
        Dim pEnumSubTypes As IEnumSubtype
        Dim lSubT As Long
        Dim sSubT As String
        pEnumSubTypes = pSubType.Subtypes

        ' loop through all of the subtypes and bring up a message 
        ' box with each subtype's code and name 
        sSubT = pEnumSubTypes.Next(lSubT)
        While Len(sSubT) > 0
            ' do something. Here we are sending the subtype code and name 
            ' to a message box 
            g_FacilityTypeLookup.Add(lSubT, pSubType.SubtypeName(lSubT))
            sSubT = pEnumSubTypes.Next(lSubT)
        End While

      


        pFeatlayerIE.FeatureClass = m_edgeShp
        pFeatSelect = pFeatlayerIE
        pFeatSelect.SelectFeatures(Nothing, esriSelectionResultEnum.esriSelectionResultNew, False)
        GlobalMod.m_EdgeSSet = pFeatSelect.SelectionSet

        'export test
        'Dim outFC As IFeatureClass
        'outFC = FeatureClassToShapefile(pathName, pFeatlayerIE, "testout")
        'Dim field As IField = New FieldClass()
        'Dim fieldEdit As IFieldEdit = CType(field, IFieldEdit)
        'fieldEdit.Type_2 = esriFieldType.esriFieldTypeInteger
        'fieldEdit.Name_2 = "NewINode"
        'fieldEdit.Length_2 = 20
        'AddFieldToFeatureClass(outFC, field)







        Dim filenameJ As String
        filenameJ = m_layers(21)
        'filenameJ = "sde.PSRC.ScenarioJunctv2"
        m_junctShp = get_FeatureLayer3(filenameJ).FeatureClass
        m_turnLayer = get_FeatureLayer3(m_layers(12))
        m_projectLayer = get_FeatureLayer3(m_layers(8))
        m_transitLayer = get_FeatureLayer3(m_layers(13))
        Pfltransitpoints = get_FeatureLayer3(m_layers(23))
        'Pfltransitpoints = get_FeatureLayer3("sde.SDE.TransitPoints")
        'user chose to do EVERYTHING
        Dim strTransitQuery As String = "InServiceDate=" & Me.txtYear.Text & " AND OutServiceDate >=" & Me.txtYear.Text
        FeatureClassToShapefile(pathName, m_transitLayer, iName, strTransitQuery)
        'ExportTransit(pathName, iName)

        If chkExportShapefiles.Checked = True Then g_ExportShapefiles = True
        If chkNetwork.Checked = True Then
            'create a temporary transit shapefile

            WriteLogLine("Finish exporting transit lines to an interim shapefile " & Now())

            'Me.lblStatus.Text = "frmBuildFiles.cmdCont_Click: Writing nodes & links buildfile"
            'Me.Refresh()

            '[021206] hyu: sometime the iName dosn't have file extension
            'tempname = Left(iName, (leng - 4))
            tempname = Replace(iName, ".shp", "", , , vbTextCompare)


            create_NetFile3(pathName, roadNetworkName, pWorkspaceI, m_application)
        End If

        If chkParkRides.Checked = True Then
            'Me.lblStatus.Text = "frmBuildFiles.cmdCont_Click: Writing Park&Ride  buildfile from " & pFeatlayerIE.Name
            'Me.Refresh()
            createParkRideFile(pathName, parkAndRideName)
        End If

        If chkTurns.Checked = True Then


            'Me.lblStatus.Text = "frmBuildFiles.cmdCont_Click: Writing turn buildfile from " & pFeatlayerIE.Name
            'Me.Refresh()
            '[021206] hyu: sometime the iName dosn't have extension
            'tempname = Left(iName, (leng - 4))
            tempname = Replace(iName, ".shp", "", , , vbTextCompare)

            create_TurnFile(pathName, turnsName, pFeatlayerIE, CType(Me.txtYear.Text, Integer), m_turnLayer)

        End If

        If chkTolls.Checked = True Then
            'Me.lblStatus.Text = "frmBuildFiles.cmdCont_Click: Writing Toll buildfile from " & pFeatlayerIE.Name
            'Me.Refresh()
            '[021206] hyu: sometime the iName dosn't have extension
            '      tempname = Left(iName, (leng - 4))
            tempname = Replace(iName, ".shp", "", , , vbTextCompare)
            createTollsFile(pathName, iName, m_dctReservedNodes, m_application, m_projectLayer)
        End If

        If chkTransit.Checked = True Then
            StefanRetrace.RetraceTransitSegments(m_application, txtYear.Text)
            'Me.lblStatus.Text = "frmBuildFiles.cmdCont_Click: Writing transit routes buildfile"
            'Me.Refresh()
            '[021206] hyu: sometime the iName dosn't have extension
            '        tempname = Left(iName, (leng - 4))
            tempname = Replace(iName, ".shp", "", , , vbTextCompare)
            If rdoOldTransit.Checked Then
                StefanRetrace.create_TransitFile5(pathName, transitName)
            Else
                StefanRetrace.create_TransitFile_NewSchema2(pathName, transitName)
            End If
            'StefanRetrace.create_TransitFile5(pathName, transitName)
            'StefanRetrace.create_TransitFile_NewSchema2(pathName, transitName)
        End If

        'Me.lblStatus.Text = "BUILDFILE WRITE COMPLETED " & Now()
        'Me.Refresh()

        Close()
        'pWSF = Nothing
        'pFWS = Nothing
    End Sub
    Private Sub ExportTransit(ByVal pathName As String, ByVal sName As String)
        Dim pWSF As IWorkspaceFactory2
        Dim pFWS As IFeatureWorkspace
        pWSF = New ShapefileWorkspaceFactory
        pFWS = pWSF.OpenFromFile(pathName, 0)


        Dim pFCls As IFeatureClass
        pFCls = get_FeatureLayer3(m_layers(13)).FeatureClass

        Dim pClone As IClone
        Dim pFlds As IFields2
        pClone = pFCls.Fields
        pFlds = pClone.Clone

        Dim pFCls2 As IFeatureClass
        If Dir(pathName & "\" & sName & "TransitLines.*") <> "" Then
            Dim pDS As IDataset
            pFCls2 = pFWS.OpenFeatureClass(sName & "TransitLines")
            pDS = pFCls2
            pDS.Delete()
        End If

        pFCls2 = pFWS.CreateFeatureClass(sName & "TransitLines", pFlds, Nothing, Nothing, esriFeatureType.esriFTSimple, pFCls.ShapeFieldName, "")

        Dim pFCS As IFeatureCursor
        Dim pFilt As IQueryFilter2
        Dim pFt As IFeature
        Dim pIns As IFeatureCursor, pFBuf As IFeatureBuffer
        Dim fld As Long, i As Long
        Dim pStr As String

        pFilt = New QueryFilter
        '090808 SEC changed this so that only transit routes for selected model year are picked
        pFilt.WhereClause = "InServiceDate=" & Me.txtYear.Text & " AND OutServiceDate >=" & Me.txtYear.Text
        '    pFilt.WhereClause = "InServiceDate=" & Me.txtYear
        'MsgBox pFCls.featurecount(pFilt)
        pIns = pFCls2.Insert(True)
        pFCS = pFCls.Search(pFilt, True)
        pFt = pFCS.NextFeature

        Do Until pFt Is Nothing
            pFBuf = pFCls2.CreateFeatureBuffer
            With pFBuf
                .Shape = pFt.ShapeCopy
                For i = 0 To pFCls.Fields.FieldCount - 1
                    'If pFt.Fields.field(i).Type <> esriFieldTypeGeometry And pFt.Fields.field(i).Type <> esriFieldTypeOID Then
                    pStr = pFt.Fields.Field(i).Name
                    If pStr.Length > 10 Then

                        pStr = pStr.Remove(10, pStr.Length - 10)
                    End If
                    fld = pFBuf.Fields.FindField(pStr)
                    If fld > 0 And Not IsDBNull(pFt.Value(i)) Then .Value(fld) = pFt.Value(i)
                    'End If
                Next i
            End With
            pIns.InsertFeature(pFBuf)
            pFt = pFCS.NextFeature
        Loop

        pFCls2 = Nothing
        pIns = Nothing
        pFBuf = Nothing
        pFWS = Nothing
        pFCls = Nothing
        pFCS = Nothing
        pWSF = Nothing
        pFlds = Nothing
        pClone = Nothing

    End Sub
    Public Sub InitiateMainVariables()
        m_Doc = m_application.Document
        m_ActiveView = m_Doc.ActiveView
        m_Map = m_ActiveView.FocusMap


    End Sub
    Private Function get_FeatureLayer3(ByVal featClsname As String) As IFeatureLayer
        'returns the feature class stated in featClsname
        'on error GoTo eh




        Dim pEnumLy As IEnumLayer


        m_Map = m_ActiveView.FocusMap
        pEnumLy = m_Map.Layers

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

End Class
