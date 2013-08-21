

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
Public Class frmCreateScenarioFromSelProjects
    Private _scenarioName As String
    Private Sub frmCreateScenarioFromSelProjects_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

    End Sub
    Private _passedFilePath As String
    Private _passedIMap As IMap
    Private _passedIApp As IApplication
    Private _passed_fXchk As Boolean
    Private _ScenarioNameCollection As Collection

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
    Public Property [PassedScenarioNameCollection]() As Collection
        Get
            Return _ScenarioNameCollection
        End Get
        Set(ByVal value As Collection)
            _ScenarioNameCollection = value

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

    Private Sub btnRun_Click(sender As System.Object, e As System.EventArgs) Handles btnRun.Click
        _scenarioName = txtScenarioTitle.Text
        Me.DialogResult = Windows.Forms.DialogResult.OK
        Me.Hide()
    End Sub
End Class