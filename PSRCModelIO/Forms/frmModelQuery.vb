
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
Public Class frmModelQuery
    Public m_application As IApplication
    'Public m_App As IApplication
    'Public m_App As IApplication
    Public m_Doc As IMxDocument
    Public m_Map As IMap
    Public m_ActiveView As IActiveView
    Public g_FWS As IFeatureWorkspace
    Public m_Editor As ESRI.ArcGIS.Editor.IEditor
    Public m_title As String
    Public m_descript As String
    Public m_path As String
    Public g_frmChkProjects As New frmChkProjects
    Public g_OldTAZ As Boolean




    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDesc.TextChanged

    End Sub

    Private Sub Label3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label3.Click

    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        ' SetMainVariables()
        If rdoOldTAZ.Checked Then
            g_OldTAZ = True
        Else
            g_OldTAZ = False

        End If



        If loaded = False Then
            '    Set g_FWS = setSDEWorkspace
            InitiateMainVariables()
            g_FWS = getPGDws(m_Map)
        End If

        'pan Option to build transit is removed for now
        ' If MsgBox("Do you want to build transit? (thinning will still retain features used by transit unless you choose not to)", vbYesNo, "Layer Source") = vbYes Then fBuildTransit = True Else fBuildTransit = False
        If MessageBox.Show("Do you want to keep edges & juncts used by tolls, turns, and transit?", "Log?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            fXchk = True
        Else : fXchk = False
        End If
        'If MsgBox("Do you want to keep edges & juncts used by tolls, turns, and transit?", vbYesNo, "Layer Source") = vbYes Then fXchk = True Else fXchk = False

        'open a log file
        If MessageBox.Show("Do you want full details in the log (mayb be a huge file!)?", "Log?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            fVerboseLog = True
        Else : fVerboseLog = False
        End If
        OpenLogFile("C:", "LogModelPrep.txt")

        'note that m_app is set in GlobalMod.setMainVariables
        init(m_App, Me.txtSchema.Text.ToString)







        g_frmBuildFiles.txtOffset.Text = Me.txtOffset.Text

        Dim year As String
        Dim yearnum As Long
        Dim mydate As Date

        If (txtYear.Text = "") Then
            MsgBox("Please enter a scenario year")
        Else
            yearnum = CLng(txtYear.Text)
            g_ModelYear = yearnum
            year = "12/30/" + txtYear.Text
            mydate = year

        End If

        Debug.Print("")
        Debug.Print("Start here: " & Now())

        Me.Hide()
        'If (ckcAdv.value = True) Then
        'GlobalMod.advCheck = True
        'Else
        'GlobalMod.advCheck = False
        'End If

        If (ckcPseudo.Checked = True) Then
            GlobalMod.pseudothin = True
        Else
            GlobalMod.pseudothin = False
        End If

        'Dissolve by Facility Type?
        If chkFT.Checked = True Then
            GlobalMod.FTthin = True
        Else
            GlobalMod.FTthin = False

        End If

        If (ckcActive.Checked = True) Then
            GlobalMod.activethin = True
        Else
            GlobalMod.activethin = False
        End If

        Dim fcstring As String
        Dim i As Integer, temp As Integer
        If (ckcFC.Checked = True) Then
            GlobalMod.FCthin = True

            For i = 0 To cmbFC.Items.Count - 1

                If cmbFC.SelectedItem(i) = True Then
                    temp = i
                    'coded out the following code to make it run
                    'lookupFC(temp)
                    If Len(fcstring) > 1 Then
                        fcstring = fcstring + " Or FunctionalClass = " + CStr(temp)
                    Else
                        fcstring = "FunctionalClass = " + CStr(temp)
                    End If
                End If
            Next i
            m_FC = fcstring

        Else
            FCthin = False
        End If
        m_Offset = CType(txtOffset.Text, Long)

        'MsgBox mydate & " " & yearnum

        select_InServiceNet(mydate, txtTitle.Text, txtDesc.Text, yearnum, g_OldTAZ)
        If m_EdgeSSet.Count = 0 Then
            MsgBox("There's No edges are selected", vbInformation)
            WriteLogLine("No edges are selected")
            Close()
        Else
            'GlobalMod.testexport
            '[090506] pan:  uncomment frmchkprojects and set pathname
            g_frmChkProjects.PassedFilePath = m_path
            g_frmChkProjects.PassedIMap = m_Map
            g_frmChkProjects.PassedIApp = m_application
            g_frmChkProjects.passed_fXchk = fXchk
            g_frmChkProjects.passedScenarioDescription = txtDesc.Text
            g_frmChkProjects.passedScenarioDescription = txtTitle.Text

            Me.Dispose()
            GC.Collect()
            g_frmChkProjects.ShowDialog()


        End If
    End Sub

    Private Sub frmModelQuery_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
    Public Sub InitiateMainVariables()
        m_Doc = m_application.Document
        m_ActiveView = m_Doc.ActiveView
        m_Map = m_ActiveView.FocusMap


    End Sub
    Public Sub select_InServiceNet(ByVal mydate As Date, ByVal title As String, ByVal descript As String, ByVal theyear As Long, ByVal OldTAZ As Boolean)
        'called by frmModelQUery
        'select the current transref network
        'on error GoTo eh
        WriteLogLine("called select_InServiceNet")
        WriteLogLine("--------------------------")

        Dim pFeatLayerE As IFeatureLayer, pFeatLayerJ As IFeatureLayer

        pFeatLayerE = get_MasterNetwork(0, m_Map)
        pFeatLayerJ = get_MasterNetwork(1, m_Map)

        Dim pQF As IQueryFilter
        pQF = New QueryFilter
        Dim currentDate As Date
        Dim inserviceStr As String
        Dim inFieldsE As String, inFieldsJ As String
        Dim pFields As IFields
        pFields = pFeatLayerE.FeatureClass.Fields
        Dim i As Long
        Dim strActiveE As String 'ActiveLink field
        Dim strFC As String 'ActiveLink field
        Dim intThin As Integer 'code for thinning selection in ActiveLink

        ' future in service date to be used in openTIPorMTP
        ' thinning code to 1 for existing emme2 net

        intThin = 1 ' exists in the old emme2 network links
        inserviceyear = theyear
        inserviceDate = mydate

        'check if user selected to thin by activelink 1/31/05
        'Currently use ActiveLink value 1 to thin


        pFields = pFeatLayerE.FeatureClass.Fields
        For i = 0 To pFields.FieldCount - 1
            If (LCase(pFields.Field(i).Name) = "inservicedate") Then
                inFieldsE = pFields.Field(i).Name
                'Exit For
            End If
            If (LCase(pFields.Field(i).Name) = "activelink") Then
                strActiveE = pFields.Field(i).Name
                'Exit For
            End If

            If (LCase(pFields.Field(i).Name) = "functionalclass") Then
                strFC = pFields.Field(i).Name
                'Exit For
            End If
        Next i


        Dim pFeatSelectE As IFeatureSelection, pFeatSelectJ As IFeatureSelection

        If (strActiveE = "") Then
            MsgBox("Did not find the ActiveLink field in TransRefEdges")
            Exit Sub
        End If

        If (inFieldsE = "") Then
            MsgBox("Did not find the InServiceDate field in TransRefEdges")
            Exit Sub
        End If
        '*************************************************************
        'jaf--not sure current date is valid, but OK for now
        '*************************************************************
        'get current date
        'currentDate = Date
        Dim querystring As String
        'pan For testing changed querystring to greater than 0ne
        If GlobalMod.activethin = True Then
            If OldTAZ = True Then
                'querystring = strActiveE + " > 0 And "
                querystring = strActiveE + " > 0 And " & strActiveE + " <> 998 And " & strActiveE + " <> 995 And "
            Else
                querystring = strActiveE + " > 0 And " & strActiveE + " <> 999 And "
            End If

        End If
            If GlobalMod.FCthin = True Then
                querystring = querystring + m_FC + " And "
            End If
            'pQF.WhereClause = inFieldsE + " <= #" + CStr(currentDate) + "#"
            querystring = querystring + inFieldsE + " <= " + CStr(inserviceyear) & " AND " & g_OutSvcDate & ">" & CStr(inserviceyear)
            WriteLogLine("select_InServiceNet filter=" & querystring)
            pQF.WhereClause = querystring
            pFeatSelectE = pFeatLayerE 'QI
            WriteLogLine("pQF.WhereClause=" & pQF.WhereClause)
            pFeatSelectE.SelectFeatures(pQF, esriSelectionResultEnum.esriSelectionResultNew, False)
            WriteLogLine("edge selection complete")

            ' global selection sets
            m_EdgeSSet = pFeatSelectE.SelectionSet
            WriteLogLine("select_InServiceNet m_EdgeSSet.count=" & m_EdgeSSet.Count)

            'not getting junction selectin  here anymore- do on the fly in create_Sceanrio 11/16/03
            '  m_JunctSSet = pFeatSelectJ.SelectionSet

            Dim length As Long
            ' globals for used in creat_ScenarioROw
            Dim tempString As String
            length = Len(title)
            If (length < 100) Then
                tempString = Microsoft.VisualBasic.Left(title, 99)
                m_title = tempString
            Else
                m_title = title
            End If
            length = Len(descript)
            If (length < 255) Then
                tempString = Microsoft.VisualBasic.Left(descript, 254)
                m_descript = tempString
            Else
                m_descript = descript
            End If

            WriteLogLine("finished select_inServiceNet at " & Now())
            pFeatLayerJ = Nothing
            pFeatLayerE = Nothing

            Exit Sub
eh:

            CloseLogFile("PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.select_inServiceNet")
            MsgBox("PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.select_inServiceNet")

            'MsgBox Err.Description, vbInformation, "select_inServiceNet"
            pFeatLayerJ = Nothing
            pFeatLayerE = Nothing

    End Sub

    Private Sub btnOpenFBD1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOpenFBD1.Click
        FolderBrowserDialog1.ShowDialog()
        txtOutPutDir.Text = FolderBrowserDialog1.SelectedPath.ToString
        m_path = txtOutPutDir.Text
    End Sub

    Private Sub FolderBrowserDialog1_HelpRequest(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FolderBrowserDialog1.HelpRequest

    End Sub

    Private Sub FolderBrowserDialog1_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles FolderBrowserDialog1.Disposed

    End Sub

    Private Sub chkDisolveOnly_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkDisolveOnly.CheckedChanged

    End Sub
End Class