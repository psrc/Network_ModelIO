

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
Public Class frmScenarioProjects
    Private _passedFilePath As String
    Private _passedIMap As IMap
    Private _passedIApp As IApplication
    Private _passed_fXchk As Boolean
    Private _ScenarioNameCollection As Collection
    Private _scenarioName As String
    Public Property [PassedFilePath]() As String
        Get
            Return _passedFilePath
        End Get
        Set(ByVal value As String)
            _passedFilePath = value

        End Set
    End Property
    Public Property [scenarioName]() As String
        Get
            Return _scenarioName
        End Get
        Set(ByVal value As String)
            _scenarioName = value

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
    Public Property [PassedScenarioNameCollection]() As Collection
        Get
            Return _ScenarioNameCollection
        End Get
        Set(ByVal value As Collection)
            _ScenarioNameCollection = value

        End Set
    End Property
    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboScenarios.SelectedIndexChanged

    End Sub

    Private Sub frmScenarioProjects_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim item As Object
        'remove any items that might in the combo box
        'Try

        If cboScenarios.Items.Count > 0 Then

            cboScenarios.Items.Clear()

        End If
        'load the combo box 
        For Each item In PassedScenarioNameCollection
            cboScenarios.Items.Add(item.ToString)
            cboScenarios.Refresh()
        Next item


        'Catch ex As Exception


        'MessageBox.Show(ex.ToString)

        'End Try



    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        _scenarioName = cboScenarios.SelectedItem.ToString
        Me.DialogResult = Windows.Forms.DialogResult.OK
        Me.Hide()

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.DialogResult = Windows.Forms.DialogResult.Cancel


    End Sub
    
End Class