Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports System.Runtime.InteropServices

<ComClass(ArcGISToolbar1.ClassId, ArcGISToolbar1.InterfaceId, ArcGISToolbar1.EventsId), _
 ProgId("PSRCModelIO.ArcGISToolbar1")> _
Public NotInheritable Class ArcGISToolbar1
    Inherits BaseToolbar

#Region "COM Registration Function(s)"
    <ComRegisterFunction(), ComVisibleAttribute(False)> _
    Public Shared Sub RegisterFunction(ByVal registerType As Type)
        ' Required for ArcGIS Component Category Registrar support
        ArcGISCategoryRegistration(registerType)

        'Add any COM registration code after the ArcGISCategoryRegistration() call

    End Sub

    <ComUnregisterFunction(), ComVisibleAttribute(False)> _
    Public Shared Sub UnregisterFunction(ByVal registerType As Type)
        ' Required for ArcGIS Component Category Registrar support
        ArcGISCategoryUnregistration(registerType)

        'Add any COM unregistration code after the ArcGISCategoryUnregistration() call

    End Sub

#Region "ArcGIS Component Category Registrar generated code"
    ''' <summary>
    ''' Required method for ArcGIS Component Category registration -
    ''' Do not modify the contents of this method with the code editor.
    ''' </summary>
    Private Shared Sub ArcGISCategoryRegistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxCommandBars.Register(regKey)

    End Sub
    ''' <summary>
    ''' Required method for ArcGIS Component Category unregistration -
    ''' Do not modify the contents of this method with the code editor.
    ''' </summary>
    Private Shared Sub ArcGISCategoryUnregistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxCommandBars.Unregister(regKey)

    End Sub

#End Region
#End Region

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "dae547fe-9d28-485e-b49d-f7a0d18a883a"
    Public Const InterfaceId As String = "952ad3b0-cf4a-4bd0-a103-3e7ffa9c6c8e"
    Public Const EventsId As String = "cb634fbf-5879-42c1-8db1-52a290d18ca1"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()

        '
        'TODO: Define your toolbar here by adding items
        '
        'AddItem("esriArcMapUI.ZoomOutTool")
        'BeginGroup() 'Separator
        addItem("{cf0af72b-ddb2-4e17-a59b-fc643a3dccc1}", 1) 'undo command
        AddItem("{9ed346bc-ce0e-4541-a584-e482128e00d7}", 1) 'undo command

        'AddItem(New Guid("FBF8C3FB-0480-11D2-8D21-080009EE4E51"), 2) 'redo command
    End Sub

    Public ReadOnly Property Caption() As String
        Get
            'TODO: Replace bar caption
            Return "My VB.Net Toolbar"
        End Get
    End Property

    Public ReadOnly Property Name() As String
        Get
            'TODO: Replace bar ID
            Return "ArcGISToolbar1"
        End Get
    End Property

    Private Sub addItem(ByVal p1 As String, ByVal p2 As Integer)
        Throw New NotImplementedException
    End Sub

End Class
