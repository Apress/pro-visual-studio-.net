Public Class [!output SAFE_ITEM_NAME] 
        Inherits Exception

        Public Sub New(ByVal message As String)
            MyBase.New(message)
        End Sub

        Public Sub New(ByVal message As String, ByVal cause As Exception)
            MyBase.New(message, cause)
        End Sub

        Public Sub New()
            MyBase.New()
        End Sub

End Class
