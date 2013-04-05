Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Carto

Module hyu_globalVar
    Public Const g_PSRCEdgeID = "PSRCEDGEID"
    Public Const g_PSRCJctID = "PSRCJunctID"
    Public Const g_INode = "INode"
    Public Const g_JNode = "JNode"
    Public Const g_Modes = "MODES"
    Public Const g_LinkType = "LinkType"
    Public Const g_FacType = "FacilityType"
    Public Const g_TrLineNo = "TRANSLINENO"
    Public Const g_TrlineID = "LINEID"
    Public Const g_OneWay = "OneWay"
    Public Const g_UseEmmeN = "UseEmmeN"
    Public Const g_InSvcDate = "InServiceDate"
    Public Const g_OutSvcDate = "OutServiceDate"
    Public g_frmPath As New frmPath
    Public g_frmMapped As New frmMapped
    Public g_fmrNetLayer As New frmNetLayer
    Public g_frmUnMap_model As New frmUnMap_model
    Public g_frmChkProjects As New frmChkProjects
    Public g_frmNetLayer As New frmNetLayer
    Public g_frmOutNetSelect As New frmOutNetSelect
    Public g_frmModelQuery As New frmModelQuery
    Public g_frmRender As New frmRender
    Public g_frmLocateEdg_Jun As New frmLocateEdg_Jun
    Public g_frmBuildFiles As New frmBuildFiles2
    Public g_frmIDfacilities As New frmIDFacilities
    Public g_frmScenarioProjects As New frmScenarioProjects

    Public g_frmEvent As New frmEvent

    Public g_sOutPath As String
    Public g_sOutName As String
    Public g_TblMode As ITable
    Public Pfltransitpoints As IFeatureLayer
    Public g_ModelYear As Integer


    Public gTblModeAtt As ITable
    Public g_xMove As Double
    Public g_yMove As Double




End Module
