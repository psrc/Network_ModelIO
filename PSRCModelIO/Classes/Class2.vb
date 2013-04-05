Imports System.Runtime.InteropServices

<ComClass(Class2.ClassId, Class2.InterfaceId, Class2.EventsId), _
 ProgId("PSRCModelIO.Class1")> _
Public Class Class2

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "91a31037-fbe2-48bf-b4fe-f9a370c102f2"
    Public Const InterfaceId As String = "5c8b6394-912c-4803-8742-f0de2678557d"
    Public Const EventsId As String = "2a735afd-d7c6-4f98-9912-5c3d48857454"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub

End Class


