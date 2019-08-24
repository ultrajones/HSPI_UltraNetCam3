<Serializable()>
Public Class SnapshotInfo

  Private m_SnapshotFileName As String = ""
  Private m_ThumbnailFileName As String = ""

  Private m_NetCamId As String = ""
  Private m_EventId As String = ""
  Private m_CreationDate As String = ""
  Private m_SnapshotAge As String = ""

  Public Sub New(ByVal NetCamId As String, _
                 ByVal EventId As String, _
                 ByVal ThumbnailFileName As String, _
                 ByVal SnapshotFileName As String, _
                 ByVal CreationDate As DateTime)

    Me.m_NetCamId = NetCamId
    Me.m_EventId = EventId

    Me.m_SnapshotFileName = SnapshotFileName
    Me.m_ThumbnailFileName = ThumbnailFileName

    Me.m_CreationDate = CreationDate.ToString("yyyy-MM-dd HH:mm:ss")
    Me.m_SnapshotAge = GetSnapshotAge(CreationDate)

  End Sub

  Public ReadOnly Property NetCamId() As String
    Get
      Return Me.m_NetCamId
    End Get
  End Property

  Public ReadOnly Property EventId() As String
    Get
      Return Me.m_EventId
    End Get
  End Property

  Public ReadOnly Property SnapshotFileName() As String
    Get
      Return Me.m_SnapshotFileName
    End Get
  End Property

  Public ReadOnly Property ThumbnailFileName() As String
    Get
      Return Me.m_ThumbnailFileName
    End Get
  End Property

  Public ReadOnly Property LightboxId() As String
    Get
      Return String.Format("lightbox[{0}]", Me.m_EventId)
    End Get
  End Property

  Public ReadOnly Property CreationDate() As String
    Get
      Return Me.m_CreationDate
    End Get
  End Property

  Public ReadOnly Property SnapshotAge() As String
    Get
      Return Me.m_SnapshotAge
    End Get
  End Property

  ''' <summary>
  ''' Formats the HomeSeer device last change
  ''' </summary>
  ''' <param name="Date2"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function GetSnapshotAge(ByRef Date2 As DateTime) As String

    Try
      Dim Date1 As DateTime = DateTime.Now
      Dim ts As TimeSpan = Date1.Subtract(Date2)

      If ts.Days > 0 Then
        Return String.Format("{0}d,{1}h,{2}m", ts.Days.ToString.PadLeft(2, "0"), ts.Hours.ToString.PadLeft(2, "0"), ts.Minutes.ToString.PadLeft(2, "0"))
      Else
        Return String.Format("{0}d,{1}h,{2}m", ts.Days.ToString.PadLeft(2, "0"), ts.Hours.ToString.PadLeft(2, "0"), ts.Minutes.ToString.PadLeft(2, "0"))
      End If

    Catch pEx As Exception
      Return "?:?"
    End Try

  End Function

End Class
