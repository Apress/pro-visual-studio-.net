Public Class Customer
    Protected _customerId As Integer
    Private _firstName As String
    Private _lastName As String



	''' Gets or sets _customerId
	Public Property CustomerId() As Integer
		Get
			Return Me._customerId
		End Get
		Set(ByVal Value As Integer)
			Me._customerId = Value
		End Set
	End Property


	''' Gets or sets _firstName
	Public Property FirstName() As String
        Get
            ' Comment created by Jan Machacek
            '
            Return Me._firstName
        End Get
        Set(ByVal Value As String)
            Me._firstName = Value
        End Set
    End Property

End Class
