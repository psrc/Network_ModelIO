Imports Scripting
Public Class clsCollection
    Private pDictionary As New Dictionary
    Public Sub New()
        pDictionary = New Dictionary
    End Sub
    Public Property myDictionary() As Dictionary
        Get
            myDictionary = pDictionary
        End Get
        Set(ByVal value As Dictionary)
            value = pDictionary
        End Set
    End Property
    Public ReadOnly Property ItemExists(ByVal item As Object) As Boolean
        Get

            If pDictionary.Exists(item) Then
                Return True
            Else
                Return False
            End If

        End Get
        
    End Property
End Class
