Public Class PrjSelected
    Public PrjId As String
    Public database As String     'database containing this project ("TIP" or "MTP")
    Public listselected As Long   'project is unmapped (1), updated (2) or unmappable (3)
    Public index As Long
End Class
