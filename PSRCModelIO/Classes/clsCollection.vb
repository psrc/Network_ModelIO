Imports Scripting
Public Class clsCollection
    Private pDictionary As New Dictionary(Of Object, Object)
    Public Sub New()
        pDictionary = New Dictionary(Of Object, Object)
    End Sub
    Public Property myDictionary() As Dictionary(Of Object, Object)
        Get
            myDictionary = pDictionary
        End Get
        Set(ByVal value As Dictionary(Of Object, Object))
            value = pDictionary
        End Set
    End Property
    Public ReadOnly Property ItemExists(ByVal item As Object) As Boolean
        Get

            If pDictionary.ContainsKey(item) Then
                Return True
            Else
                Return False
            End If

        End Get
        
    End Property
End Class
