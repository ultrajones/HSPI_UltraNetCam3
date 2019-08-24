Imports System.Threading
Imports System.Net
Imports System.IO
Imports System.Drawing
Imports System.Text.RegularExpressions
Imports System.Text
Imports System.Web

Public Class NetCamDevice

  Protected SendCommandThread As Thread

  Protected CommandQueue As New Queue

  Protected m_Id As Integer = 0
  Protected m_NetCamId As String = ""
  Protected m_NetCamName As String = ""
  Protected m_NetCamAddress As String = ""
  Protected m_NetCamPort As Integer = 80
  Protected m_SnapshotPath As String = ""
  Protected m_SnapshotPathNext As String = ""
  Protected m_VideostreamPath As String = ""
  Protected m_AuthUser As String = ""
  Protected m_AuthPass As String = ""
  Protected m_SnapshotDirectory As String = ""
  Protected m_IsFoscam As Boolean = False

#Region "NetCam Object"

  Public ReadOnly Property Id() As Integer
    Get
      Return Me.m_Id
    End Get
  End Property

  Public ReadOnly Property NetCamId() As String
    Get
      Return Me.m_NetCamId
    End Get
  End Property

  Public ReadOnly Property Name() As String
    Get
      Return Me.m_NetCamName
    End Get
  End Property

  Public ReadOnly Property Address() As String
    Get
      Return Me.m_NetCamAddress
    End Get
  End Property

  Public ReadOnly Property Port() As Integer
    Get
      Return Me.m_NetCamPort
    End Get
  End Property

  Public ReadOnly Property SnapshotPath() As Integer
    Get
      Return Me.m_SnapshotPath
    End Get
  End Property

  Public ReadOnly Property VideostreamPath() As Integer
    Get
      Return Me.m_VideostreamPath
    End Get
  End Property

  Public ReadOnly Property AuthUser() As String
    Get
      Return Me.m_AuthUser
    End Get
  End Property

  Public ReadOnly Property AuthPass() As String
    Get
      Return Me.m_AuthPass
    End Get
  End Property

  Public ReadOnly Property SnapshotDirectory() As String
    Get
      Return Me.m_SnapshotDirectory
    End Get
  End Property

  Public ReadOnly Property IsFoscam() As Boolean
    Get
      Return Me.m_IsFoscam
    End Get
  End Property

  Public ReadOnly Property CommandQueueCount() As Integer
    Get
      Return Me.CommandQueue.Count
    End Get
  End Property

  Public Sub New(ByVal NetCamId As Integer, _
                 ByVal NetCamName As String, _
                 ByVal NetCamAddress As String, _
                 ByVal NetCamPort As Integer, _
                 ByVal SnapshotPath As String, _
                 ByVal VideostreamPath As String, _
                 ByVal AuthUser As String, _
                 ByVal AuthPass As String)

    MyBase.New()

    Try

      '
      ' Set the variables for this object
      '
      Me.m_Id = NetCamId

      Me.m_NetCamId = String.Format("NetCam{0}", NetCamId.ToString.PadLeft(3, "0"))

      Me.m_NetCamName = NetCamName
      Me.m_NetCamAddress = NetCamAddress
      Me.m_NetCamPort = NetCamPort
      Me.m_SnapshotPath = SnapshotPath
      Me.m_VideostreamPath = VideostreamPath

      Me.m_AuthUser = AuthUser
      Me.m_AuthPass = AuthPass

      Dim strMessage As String = ""

      '
      ' Start the process command queue thread
      '
      SendCommandThread = New Thread(New ThreadStart(AddressOf ProcessCommandQueue))
      SendCommandThread.Name = "ProcessCommandQueue"
      SendCommandThread.Start()

      strMessage = SendCommandThread.Name & " Thread Started"
      WriteMessage(strMessage, MessageType.Debug)

      '
      ' Make sure Snapshot directories exist
      '
      If Directory.Exists(gSnapshotDirectory) = False Then
        Directory.CreateDirectory(gSnapshotDirectory)
      End If

      m_SnapshotDirectory = FixPath(String.Format("{0}\{1}", gSnapshotDirectory, m_NetCamId))
      If Directory.Exists(m_SnapshotDirectory) = False Then
        Directory.CreateDirectory(m_SnapshotDirectory)
      End If

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "New()")
    End Try

  End Sub

  Protected Overrides Sub Finalize()

    Try
      '
      ' Abort SendCommandThread
      '
      If SendCommandThread.IsAlive = True Then
        SendCommandThread.Abort()
      End If

    Catch pEx As Exception

    End Try

    MyBase.Finalize()

  End Sub

#End Region

#Region "NetCam Command Processing"

  ''' <summary>
  ''' Adds command to command buffer for processing
  ''' </summary>
  ''' <param name="NetCamCommand"></param>
  ''' <remarks></remarks>
  Public Sub AddCommand(ByVal NetCamCommand As NetCamCommand)

    Try
      '
      ' Add NetCamCommand to CommandQueue
      '
      CommandQueue.Enqueue(NetCamCommand)

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "AddCommand()")
    End Try

  End Sub

  ''' <summary>
  ''' Processes commands and waits for the response
  ''' </summary>
  ''' <remarks></remarks>
  Protected Sub ProcessCommandQueue()

    Dim NetCamCommand As NetCamCommand
    Dim bAbortThread As Boolean = False

    Try

      While bAbortThread = False
        '
        ' Process commands in command queue
        '
        While CommandQueue.Count > 0 And gIOEnabled = True

          '
          ' Set the command response we are waiting for
          '
          NetCamCommand = CommandQueue.Peek

          If NetCamCommand.Command = "TakeEventSnapshot" Then
            '
            ' Take a snapshot
            '
            Dim strEventName As String = String.Format("Event{0}", NetCamCommand.EventId)
            Dim strEventDirectory As String = FixPath(String.Format("{0}\{1}", m_SnapshotDirectory, strEventName))
            Dim strIdentifier As String = NetCamCommand.SnapshotId.ToString.PadLeft(4, "0")

            Dim strSnapshotFilename As String = FixPath(String.Format("{0}\{1}_{2}_snapshot.jpg", strEventDirectory, m_NetCamId, strIdentifier))
            Dim strThumbnailFilename As String = FixPath(String.Format("{0}\{1}_{2}_thumbnail.jpg", strEventDirectory, m_NetCamId, strIdentifier))

            WriteMessage(String.Format("ProcessCommandQueue thread is running command: {0}.", NetCamCommand.Command), MessageType.Debug)
            Call TakeSnapshot(strSnapshotFilename, strThumbnailFilename)

            If m_SnapshotPathNext.Length Then
              m_SnapshotPath = m_SnapshotPathNext
              m_SnapshotPathNext = String.Empty
              Call TakeSnapshot(strSnapshotFilename, strThumbnailFilename)
            End If

          ElseIf NetCamCommand.Command = "CameraCGI" Then
            '
            ' Run Camera CGI
            '
            WriteMessage(String.Format("ProcessCommandQueue thread is running command: {0}.", NetCamCommand.Command), MessageType.Debug)
            Call NetCamCGI(NetCamCommand.QueryString)

          ElseIf NetCamCommand.Command = "StartEvent" Then
            '
            ' Mark the start of the Snapshot Event
            '
            StartSnapshotEvent(NetCamCommand.EventId)

          ElseIf NetCamCommand.Command = "EndEvent" Then
            '
            ' Mark the end of the Snapshot Event
            '
            EndSnapshotEvent(NetCamCommand.EventId)

          ElseIf NetCamCommand.Command = "TakeSnapshot" Then
            '
            ' Take a snapshot
            '
            Dim strSnapshotFilename As String = FixPath(String.Format("{0}\last_snapshot.jpg", m_SnapshotDirectory))
            Dim strThumbnailFilename As String = FixPath(String.Format("{0}\last_thumbnail.jpg", m_SnapshotDirectory))

            WriteMessage(String.Format("ProcessCommandQueue thread is running command: {0}.", NetCamCommand.Command), MessageType.Debug)
            Call TakeSnapshot(strSnapshotFilename, strThumbnailFilename)

            If m_SnapshotPathNext.Length Then
              m_SnapshotPath = m_SnapshotPathNext
              m_SnapshotPathNext = String.Empty
              Call TakeSnapshot(strSnapshotFilename, strThumbnailFilename)
            End If

          Else
            '
            ' Invalid command was received
            '
            WriteMessage(String.Format("ProcessCommandQueue thread received invalid command: {0}.", NetCamCommand.Command), MessageType.Error)
          End If

          '
          ' Command delay
          '
          Dim iWaitSeconds As Double = NetCamCommand.Delay

          WriteMessage(String.Format("ProcessCommandQueue thread is sleeping for {0} seconds.", iWaitSeconds.ToString), MessageType.Debug)
          Thread.Sleep(iWaitSeconds * 1000)

          CommandQueue.Dequeue()

        End While ' Done with all commands in queue

        '
        ' Give up some time to allow the main thread to populate the command queue with more commands
        '
        Thread.Sleep(50)

      End While ' Stay in thread until we get an abort/exit request

    Catch pEx As ThreadAbortException
      ' 
      ' There was a normal request to terminate the thread.  
      '
      bAbortThread = True      ' Not actually needed
      WriteMessage(String.Format("ProcessCommandQueue thread received abort request, terminating normally."), MessageType.Debug)

    Catch pEx As Exception
      '
      ' Return message
      '
      ProcessError(pEx, "ProcessCommandQueue()")

    Finally
      '
      ' Notify that we are exiting the thread 
      '
      WriteMessage(String.Format("ProcessCommandQueue terminated."), MessageType.Debug)
    End Try

  End Sub

#End Region

#Region "NetCam CGI"

  ''' <summary>
  ''' Calls Camera CGI
  ''' </summary>
  ''' <param name="strQueryString"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function NetCamCGI(ByVal strQueryString As String, Optional ByVal iTimeoutSec As Integer = 60) As String

    Dim strResult As String = String.Empty

    Try
      '
      ' Format the URL
      '
      Dim netcam_url As String = String.Format("http://{0}:{1}{2}", Me.m_NetCamAddress, Me.m_NetCamPort.ToString, strQueryString)

      '
      ' Replace HomeSeer IP Address if required
      '
      If netcam_url.Contains("$homeseer_ip") Then
        Dim strLanIPAddress As String = Regex.Match(hs.LANIP(), "\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}").Value
        netcam_url = netcam_url.Replace("$homeseer_ip", strLanIPAddress)
      End If
      WriteMessage(String.Format("NetCamCGI is running CGI command: {0}.", netcam_url), MessageType.Debug)

      '
      ' Replace User and Password if required
      '
      netcam_url = netcam_url.Replace("$user", Me.m_AuthUser).Replace("$pass", Me.m_AuthPass)
      If netcam_url.Contains("$camname") Then
        netcam_url = netcam_url.Replace("$camname", HttpUtility.UrlEncode(Me.m_NetCamName))
      End If

      '
      ' Build the HTTP Web Request
      '
      Dim lxRequest As HttpWebRequest = DirectCast(WebRequest.Create(netcam_url), HttpWebRequest)
      lxRequest.Timeout = iTimeoutSec * 1000
      lxRequest.Credentials = New NetworkCredential(Me.m_AuthUser, Me.m_AuthPass)

      '
      ' Process the HTTP Web Response
      '
      Using lxResponse As HttpWebResponse = DirectCast(lxRequest.GetResponse(), HttpWebResponse)
        Using responseStream As Stream = lxResponse.GetResponseStream()
          Using readStream As New StreamReader(responseStream, Encoding.UTF8)
            strResult = readStream.ReadToEnd
          End Using
        End Using

        lxResponse.Close()
      End Using

    Catch pEx As WebException
      '
      ' Process the error
      '
      Dim strErrorMessage As String = String.Format("The CGI request sent to {0} [{1}] failed: {2}", Me.m_NetCamId, Me.m_NetCamName, pEx.Message)
      WriteMessage(strErrorMessage, MessageType.Error)

    Catch pEx As Exception
      '
      ' Process the error
      '
      Dim strErrorMessage As String = String.Format("The CGI request sent to {0} [{1}] failed: {2}", Me.m_NetCamId, Me.m_NetCamName, pEx.Message)
      WriteMessage(strErrorMessage, MessageType.Error)

    End Try

    Return strResult

  End Function

#End Region

#Region "NetCam Snapshots"

  ''' <summary>
  ''' Takes snapshot from camera
  ''' </summary>
  ''' <param name="strSnapshotFilename"></param>
  ''' <param name="strThumbnailFilename"></param>
  ''' <param name="iTimeoutSec"></param>
  ''' <param name="strSnapshotPath"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Protected Friend Function TakeSnapshot(ByVal strSnapshotFilename As String, _
                                         ByVal strThumbnailFilename As String, _
                                         Optional ByVal iTimeoutSec As Integer = 60, _
                                         Optional ByVal strSnapshotPath As String = "") As Boolean

    Try
      '
      ' Format the URL
      '
      If strSnapshotPath.Length = 0 Then strSnapshotPath = Me.m_SnapshotPath
      Dim netcam_url As String = String.Format("http://{0}:{1}{2}", Me.m_NetCamAddress, Me.m_NetCamPort.ToString, strSnapshotPath)
      WriteMessage(String.Format("TakeSnapshot is running command: {0}.", netcam_url), MessageType.Debug)

      '
      ' Replace User and Password if required
      '
      netcam_url = netcam_url.Replace("$user", Me.m_AuthUser).Replace("$pass", Me.m_AuthPass)
      If netcam_url.Contains("$camname") Then
        netcam_url = netcam_url.Replace("$camname", HttpUtility.UrlEncode(Me.m_NetCamName))
      End If

      '
      ' Build the HTTP Web Request
      '
      Dim lxRequest As HttpWebRequest = DirectCast(WebRequest.Create(netcam_url), HttpWebRequest)
      lxRequest.Timeout = iTimeoutSec * 1000
      lxRequest.Credentials = New NetworkCredential(Me.m_AuthUser, Me.m_AuthPass)

      '
      ' Process the HTTP Web Response
      '
      Using lxResponse As HttpWebResponse = DirectCast(lxRequest.GetResponse(), HttpWebResponse)

        Dim lnBuffer As Byte()
        Dim lnFile As Byte()

        Using lxBR As New BinaryReader(lxResponse.GetResponseStream())

          Using lxMS As New MemoryStream()

            lnBuffer = lxBR.ReadBytes(1024)

            While lnBuffer.Length > 0
              lxMS.Write(lnBuffer, 0, lnBuffer.Length)
              lnBuffer = lxBR.ReadBytes(1024)
            End While

            lnFile = New Byte(CInt(lxMS.Length) - 1) {}
            lxMS.Position = 0
            lxMS.Read(lnFile, 0, lnFile.Length)

            Try

              Using image As Image = image.FromStream(lxMS)

                If Not image Is Nothing Then

                  If strSnapshotFilename.EndsWith("last_snapshot.jpg") Then
                    If File.Exists(strSnapshotFilename) = True Then
                      File.SetCreationTime(strSnapshotFilename, DateTime.Now)
                    End If
                  End If

                  If strThumbnailFilename.EndsWith("last_thumbnail.jpg") Then
                    If File.Exists(strThumbnailFilename) = True Then
                      File.SetCreationTime(strThumbnailFilename, DateTime.Now)
                    End If
                  End If

                  image.Save(strSnapshotFilename, image.RawFormat)

                  Dim iWidth As Integer = image.Width
                  Dim iHeight As Integer = image.Height
                  For i = 1 To 10 Step 1
                    iWidth = image.Width / i
                    iHeight = image.Height / i
                    If iWidth <= 200 Then Exit For
                  Next

                  Dim imageThumb As Image = image.GetThumbnailImage(iWidth, iHeight, Nothing, New IntPtr())
                  imageThumb.Save(strThumbnailFilename, image.RawFormat)

                End If

                image.Dispose()

              End Using

            Catch pEx As ArgumentException
              '
              ' We got here because the data was not an image
              '
              Dim strErrorMessage As String = String.Format("The Snapshot request sent to {0} [{1}] failed: {2}", Me.m_NetCamId, Me.m_NetCamName, pEx.Message)
              WriteMessage(strErrorMessage, MessageType.Warning)

            End Try

            lxMS.Close()
            lxBR.Close()

          End Using

        End Using

        lxResponse.Close()

      End Using

      Return True

    Catch pEx As System.Net.WebException
      '
      ' Process the WebException
      '
      Dim strErrorMessage As String = String.Format("The Snapshot request sent to {0} [{1}] failed: {2}", Me.m_NetCamId, Me.m_NetCamName, pEx.Message)
      WriteMessage(strErrorMessage, MessageType.Error)

    Catch pEx As Exception
      '
      ' Process the error
      '
      Dim strErrorMessage As String = String.Format("The Snapshot request sent to {0} [{1}] failed: {2}", Me.m_NetCamId, Me.m_NetCamName, pEx.Message)
      WriteMessage(strErrorMessage, MessageType.Error)
    End Try

    Return False

  End Function

  ''' <summary>
  ''' Starts the Snapshot Event
  ''' </summary>
  ''' <param name="EventId"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function StartSnapshotEvent(ByVal EventId As Integer) As Boolean

    Try

      Dim strEventName As String = String.Format("Event{0}", EventId.ToString)
      Dim strEventDirectory As String = FixPath(String.Format("{0}\{1}", m_SnapshotDirectory, strEventName))

      If Directory.Exists(strEventDirectory) = False Then
        Directory.CreateDirectory(strEventDirectory)
      End If

      Return True

    Catch pEx As Exception
      Return False
    End Try

  End Function

  ''' <summary>
  ''' Ends Snapshot Event
  ''' </summary>
  ''' <param name="EventId"></param>
  ''' <remarks></remarks>
  Private Sub EndSnapshotEvent(ByVal EventId As Integer)

    Try

      Dim strEventName As String = String.Format("Event{0}", EventId.ToString)
      Dim strEventDirectory As String = FixPath(String.Format("{0}\{1}", m_SnapshotDirectory, strEventName))
      Dim strEventINIFile As String = FixPath(String.Format("{0}\{1}", strEventDirectory, "Event.ini"))

      Dim iEndtimestamp As Long = ConvertDateTimeToEpoch(DateTime.Now)
      Dim strFirstSnapshot As String = ""
      Dim strLastSnapshot As String = ""

      Dim EventSnapshots As ArrayList = GetCameraEventSnapshots(m_NetCamId, strEventName)
      Dim iFramesCompleted As Integer = EventSnapshots.Count

      If iFramesCompleted > 0 Then
        Dim FirstSnapshot As SnapshotInfo = EventSnapshots.Item(0)
        strFirstSnapshot = FirstSnapshot.SnapshotFileName

        If EventSnapshots.Count > 1 Then
          Dim LastSnapshot As SnapshotInfo = EventSnapshots.Item(EventSnapshots.Count - 1)
          strLastSnapshot = LastSnapshot.SnapshotFileName
        Else
          strLastSnapshot = strFirstSnapshot
        End If

        If File.Exists(strLastSnapshot) = True Then
          Try
            Dim strSnapshotDest As String = FixPath(String.Format("{0}\last_snapshot.jpg", m_SnapshotDirectory))
            File.Copy(strLastSnapshot, strSnapshotDest, True)

            Dim strThumbnailDest As String = FixPath(String.Format("{0}\last_thumbnail.jpg", m_SnapshotDirectory))
            File.Copy(strLastSnapshot.Replace("_snapshot", "_thumbnail"), strThumbnailDest, True)

            ' Update Creation Time
            If File.Exists(strSnapshotDest) = True Then
              File.SetCreationTime(strSnapshotDest, DateTime.Now)
            End If

            ' Update Creation Time
            If File.Exists(strThumbnailDest) = True Then
              File.SetCreationTime(strThumbnailDest, DateTime.Now)
            End If

          Catch pEx As Exception
            '
            ' Process the error
            '
            Call ProcessError(pEx, "EndSnapshotEvent()")
          End Try
        End If
      End If

      Using objFile As New System.IO.StreamWriter(strEventINIFile, False)

        objFile.WriteLine(String.Format("[{0}]", strEventName))
        objFile.WriteLine(String.Format("EventDir={0}", strEventDirectory))
        objFile.WriteLine(String.Format("NetCamId={0}", m_NetCamId))
        objFile.WriteLine(String.Format("NetCamName={0}", m_NetCamName))
        objFile.WriteLine(String.Format("FirstSnapshot={0}", strFirstSnapshot))
        objFile.WriteLine(String.Format("LastSnapshot={0}", strLastSnapshot))

        objFile.Dispose()
      End Using

      '
      ' Update the databasse indicating the event is complete
      '
      EndNetCamEvent(EventId, iEndtimestamp, iFramesCompleted)

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "EndSnapshotEvent()")
    End Try

  End Sub

#End Region

End Class

