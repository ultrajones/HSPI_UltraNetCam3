Public Class NetCamEvent

  Private m_netcam_id As Integer = 0
  Private m_frames_requested As Integer
  Private m_email_recipient As String = ""
  Private m_email_attachment As String = ""
  Private m_emailed As Integer = 0
  Private m_compressed As Integer
  Private m_upload As Integer = 0
  Private m_archived As Integer = 0

  Public Sub New()

  End Sub

  Public Sub New(ByVal netcam_id As Integer, _
                 ByVal frames_requested As Integer, _
                 ByVal strEmailRecipient As String, _
                 ByVal strEmailAttachment As String, _
                 ByVal emailed As Integer, _
                 ByVal compressed As Integer, _
                 ByVal upload As Integer, _
                 ByVal archived As Integer)

    Me.m_netcam_id = netcam_id
    Me.m_frames_requested = frames_requested
    Me.m_email_recipient = strEmailRecipient
    Me.m_email_attachment = strEmailAttachment
    Me.m_emailed = emailed
    Me.m_compressed = compressed
    Me.m_upload = upload
    Me.m_archived = archived

  End Sub

  Public Property NetCamId() As Integer
    Get
      Return Me.m_netcam_id
    End Get
    Set(ByVal value As Integer)
      Me.m_netcam_id = value
    End Set
  End Property

  Public Property FramesRequested() As Integer
    Get
      Return Me.m_frames_requested
    End Get
    Set(ByVal value As Integer)
      Me.m_frames_requested = value
    End Set
  End Property

  Public Property EmailRecipient() As String
    Get
      Return Me.m_email_recipient
    End Get
    Set(ByVal value As String)
      Me.m_email_recipient = value
    End Set
  End Property

  Public Property EmailAttachment() As String
    Get
      Return Me.m_email_attachment
    End Get
    Set(ByVal value As String)
      Me.m_email_attachment = value
    End Set
  End Property

  Public Property Emailed() As Integer
    Get
      Return Me.m_emailed
    End Get
    Set(ByVal value As Integer)
      Me.m_emailed = value
    End Set
  End Property

  Public Property Compressed() As Integer
    Get
      Return Me.m_compressed
    End Get
    Set(ByVal value As Integer)
      Me.m_compressed = value
    End Set
  End Property

  Public Property Uploaded() As Integer
    Get
      Return Me.m_upload
    End Get
    Set(ByVal value As Integer)
      Me.m_upload = value
    End Set
  End Property

  Public Property Archived() As Integer
    Get
      Return Me.m_archived
    End Get
    Set(ByVal value As Integer)
      Me.m_archived = value
    End Set
  End Property

End Class