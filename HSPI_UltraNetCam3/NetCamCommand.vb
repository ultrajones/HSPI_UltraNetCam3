Public Class NetCamCommand

  Private m_Command As String = ""
  Private m_QueryString As String = ""
  Private m_Delay As Double = 0
  Private m_EventId As Integer = 0
  Private m_SnapshotId As Integer = 0

  Public Sub New(ByVal Command As String, _
                 ByVal Delay As Double, _
                 Optional ByVal EventId As Integer = 0, _
                 Optional ByVal SnapshotId As Integer = 0, _
                 Optional ByVal QueryString As String = "")

    Me.m_Command = Command
    Me.m_Delay = Delay
    Me.m_EventId = EventId
    Me.m_SnapshotId = SnapshotId
    Me.m_QueryString = QueryString

  End Sub

  Public ReadOnly Property Command() As String
    Get
      Return Me.m_Command
    End Get
  End Property

  Public ReadOnly Property Delay() As Double
    Get
      Return Me.m_Delay
    End Get
  End Property

  Public ReadOnly Property EventId() As Integer
    Get
      Return Me.m_EventId
    End Get
  End Property

  Public ReadOnly Property SnapshotId() As Integer
    Get
      Return Me.m_SnapshotId
    End Get
  End Property

  Public ReadOnly Property QueryString() As String
    Get
      Return Me.m_QueryString
    End Get
  End Property

End Class
