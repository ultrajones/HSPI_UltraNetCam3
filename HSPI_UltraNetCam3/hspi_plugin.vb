Imports System.Threading
Imports System.Text.RegularExpressions
Imports System.Net.Sockets
Imports System.Text
Imports System.Net
Imports System.Data.Common
Imports HomeSeerAPI
Imports Scheduler
Imports System.IO
Imports System.Drawing
Imports System.ComponentModel
Imports System.Data.SQLite
Imports HSPI_ULTRANETCAM3.Utilities.FTP

Module hspi_plugin

  '
  ' Declare public objects, not required by HomeSeer
  '
  Dim actions As New hsCollection
  Dim triggers As New hsCollection
  Dim conditions As New Hashtable
  Const Pagename = "Events"

  Public Devices As New Queue
  Public NetCams As New Hashtable
  Public Foscams As New Hashtable
  Public AttachmentTypes As New SortedList

  Public Const IFACE_NAME As String = "UltraNetCam3"

  Public Const LINK_TARGET As String = "hspi_ultranetcam3/hspi_ultranetcam3.aspx"
  Public Const LINK_URL As String = "hspi_ultranetcam3.html"
  Public Const LINK_TEXT As String = "UltraNetCam3"
  Public Const LINK_PAGE_TITLE As String = "UltraNetCam3 HSPI"
  Public Const LINK_HELP As String = "/hspi_ultranetcam3/UltraNetCam3_HSPI_Users_Guide.pdf"

  Public gBaseCode As String = ""
  Public gIOEnabled As Boolean = True
  Public gImageDir As String = "/images/hspi_ultranetcam3/"
  Public gHSInitialized As Boolean = False
  Public gINIFile As String = "hspi_" & IFACE_NAME.ToLower & ".ini"

  Public gFoscamAutoDiscovery As Boolean = False        ' Indicates if Foscam autodiscovery should be enabled
  Public gEventArchiveToDir As Boolean = False          ' Indicates if events should be archived to archive directory
  Public gEventArchiveToFTP As Boolean = False          ' Indicates if events should be archived to FTP directory
  Public gEventEmailNotification As Boolean = False     ' Indicates if events should generate an e-mail notification
  Public gEventCompressToZip As Boolean = True          ' Indicates if event snapshots should be compressed
  Public gSnapshotEventMax As Integer = 25              ' The default number of events to store per NetCam
  Public gSnapshotsMaxWidth As String = "Auto"          ' The size of the snapshot image
  Public gSnapshotRefreshInterval As Integer = 0        ' The default snapshot refresh interval
  Public gEventSnapshotCount As Integer = 10            ' The number of snapshots needed to create a video

  Public HSAppPath As String = ""
  Public gSnapshotDirectory As String = ""

  Public EMAIL_SUBJECT As String = String.Format("{0} - {1} - {2}", IFACE_NAME, "$camera_name", "$event_id")
  Public Const EMAIL_BODY_TEMPLATE As String = "HomeSeer captured snapshot $event_id from $camera_name [$camera_id] on " & _
                                               "$date $time~~" & _
                                               "http://$lan_ip:$lan_port/$snapshot_path~" & _
                                               "https://$wan_ip:$wan_port/$snapshot_path~"

#Region "HSPI - Public Sub/Functions"

#Region "HSPI - NetCam Controls"

  ''' <summary>
  ''' Sends CGI command to NetCam
  ''' </summary>
  ''' <param name="netcam_id"></param>
  ''' <param name="strQueryString"></param>
  ''' <param name="iWaitSeconds"></param>
  ''' <remarks></remarks>
  Public Sub NetCamCGI(ByVal netcam_id As String, ByVal strQueryString As String, Optional ByVal iWaitSeconds As Integer = 0)

    Try
      '
      ' Format the NetCam Id
      '
      Dim strNetCamId As String = netcam_id
      If Regex.IsMatch(strNetCamId, "NetCam\d\d\d") = False Then
        strNetCamId = String.Format("NetCam{0}", netcam_id.PadLeft(3, "0"))
      End If

      '
      ' Add CGI commands if NetCam exists
      '
      If NetCams.ContainsKey(strNetCamId) = True Then
        Dim NetCamDevice As NetCamDevice = NetCams(strNetCamId)

        Dim NetCamCommand As New NetCamCommand("CameraCGI", iWaitSeconds, 0, 0, strQueryString)
        NetCamDevice.AddCommand(NetCamCommand)

      Else
        '
        ' Unable to find selected NetCam
        '
        WriteMessage(String.Format("Unable to find {0}!", strNetCamId), MessageType.Warning)
      End If

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "NetCamCGI()")
    End Try

  End Sub

  ''' <summary>
  ''' NetCam Control URL
  ''' </summary>
  ''' <param name="control_id"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetNetCamControlURL(ByVal control_id As Integer) As String

    Dim strControlURL As String = ""

    Try

      Dim strSQL As String = String.Format("SELECT control_url FROM tblNetCamControls WHERE control_id={0}", control_id)

      '
      ' Execute the data reader
      '
      Using MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

        MyDbCommand.Connection = DBConnectionMain
        MyDbCommand.CommandType = CommandType.Text
        MyDbCommand.CommandText = strSQL

        SyncLock SyncLockMain
          Dim dtrResults As IDataReader = MyDbCommand.ExecuteReader()

          '
          ' Process the resutls
          '
          If dtrResults.Read() Then
            strControlURL = dtrResults("control_url")
          End If

          dtrResults.Close()
        End SyncLock

        MyDbCommand.Dispose()

      End Using

    Catch pEx As Exception

    End Try

    Return strControlURL

  End Function

  ''' <summary>
  ''' Gets the NetCam Types from the underlying database
  ''' </summary>
  ''' <param name="netcam_type"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetNetCamControlsFromDB(ByVal netcam_type As Integer) As DataTable

    Dim ResultsDT As New DataTable
    Dim strMessage As String = ""

    strMessage = "Entered GetNetCamControlsFromDB() function."
    Call WriteMessage(strMessage, MessageType.Debug)

    Try
      '
      ' Make sure the datbase is open before attempting to use it
      '
      Select Case DBConnectionMain.State
        Case ConnectionState.Broken, ConnectionState.Closed
          strMessage = "Unable to complete database query because the database " _
                     & "connection has not been initialized."
          Throw New System.Exception(strMessage)
      End Select

      Dim strSQL As String = String.Format("SELECT control_id, netcam_type, control_name, control_url FROM tblNetCamControls WHERE netcam_type={0} ORDER BY control_name", netcam_type)

      '
      ' Initialize the command object
      '
      Dim MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

      MyDbCommand.Connection = DBConnectionMain
      MyDbCommand.CommandType = CommandType.Text
      MyDbCommand.CommandText = strSQL

      '
      ' Initialize the dataset, then populate it
      '
      Dim MyDS As DataSet = New DataSet

      Dim MyDA As System.Data.IDbDataAdapter = New SQLiteDataAdapter(MyDbCommand)
      MyDA.SelectCommand = MyDbCommand

      SyncLock SyncLockMain
        MyDA.Fill(MyDS)
      End SyncLock

      '
      ' Get our DataTable
      '
      Dim MyDT As DataTable = MyDS.Tables(0)

      '
      ' Get record count
      '
      Dim iRecordCount As Integer = MyDT.Rows.Count

      If iRecordCount > 0 Then
        '
        ' Build field names
        '
        Dim iFieldCount As Integer = MyDS.Tables(0).Columns.Count() - 1
        For iFieldNum As Integer = 0 To iFieldCount
          '
          ' Create the columns
          '
          Dim ColumnName As String = MyDT.Columns.Item(iFieldNum).ColumnName
          Dim MyDataColumn As New DataColumn(ColumnName, GetType(String))

          '
          ' Add the columns to the DataTable's Columns collection
          '
          ResultsDT.Columns.Add(MyDataColumn)
        Next

        '
        ' Let's output our records	
        '
        Dim i As Integer = 0
        For i = 0 To iRecordCount - 1
          '
          ' Create the rows
          '
          Dim dr As DataRow
          dr = ResultsDT.NewRow()
          For iFieldNum As Integer = 0 To iFieldCount
            dr(iFieldNum) = MyDT.Rows(i)(iFieldNum)
          Next
          ResultsDT.Rows.Add(dr)
        Next

      End If

    Catch pEx As Exception
      '
      ' Process Exception
      '
      Call ProcessError(pEx, "GetNetCamControlsFromDB()")

    End Try

    Return ResultsDT

  End Function

  ''' <summary>
  ''' Inserts a new NetCam Controls into the database
  ''' </summary>
  ''' <param name="netcam_type"></param>
  ''' <param name="control_name"></param>
  ''' <param name="control_url"></param>
  ''' <param name="bRefreshNetCamList"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function InsertNetCamControls(ByVal netcam_type As String, _
                                       ByVal control_name As String, _
                                       ByVal control_url As String, _
                                       Optional ByVal bRefreshNetCamList As Boolean = True) As Integer

    Dim strMessage As String = ""
    Dim control_id As Integer = 0

    Try

      Select Case DBConnectionMain.State
        Case ConnectionState.Broken, ConnectionState.Closed
          strMessage = "Unable to complete database transaction because the database " _
                     & "connection has not been initialized."
          Throw New System.Exception(strMessage)
      End Select

      If netcam_type.Length = 0 Or control_name.Length = 0 Or control_url.Length = 0 Then
        Throw New Exception("One or more required fields are empty.  Unable to insert new NetCam control into the database.")
      End If

      Dim strSQL As String = String.Format("INSERT INTO tblNetCamControls (" _
                                           & " netcam_type, control_name, control_url" _
                                           & ") VALUES (" _
                                           & "'{0}', '{1}', '{2}' );", _
                                           netcam_type, control_name, control_url)
      strSQL &= "SELECT last_insert_rowid();"

      Using dbcmd As DbCommand = DBConnectionMain.CreateCommand()

        dbcmd.Connection = DBConnectionMain
        dbcmd.CommandType = CommandType.Text
        dbcmd.CommandText = strSQL

        SyncLock SyncLockMain
          control_id = dbcmd.ExecuteScalar()
        End SyncLock

        dbcmd.Dispose()

      End Using

    Catch pEx As Exception
      Call ProcessError(pEx, "InsertNetCamControls()")
      Return False
    Finally
      If bRefreshNetCamList = True Then
        RefreshNetCamList()
      End If
    End Try

    Return control_id

  End Function

  ''' <summary>
  ''' Updates existing NetCam Controls stored in the database
  ''' </summary>
  ''' <param name="control_id"></param>
  ''' <param name="control_name"></param>
  ''' <param name="control_url"></param>
  ''' <param name="bRefreshNetCamList"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function UpdateNetCamControls(ByVal control_id As Integer, _
                                       ByVal control_name As String, _
                                       ByVal control_url As String, _
                                       Optional ByVal bRefreshNetCamList As Boolean = True) As Boolean

    Dim strMessage As String = ""

    Try

      Select Case DBConnectionMain.State
        Case ConnectionState.Broken, ConnectionState.Closed
          strMessage = "Unable to complete database transaction because the database " _
                     & "connection has not been initialized."
          Throw New System.Exception(strMessage)
      End Select

      Dim strSql As String = String.Format("UPDATE tblNetCamControls SET " _
                                          & " control_name='{0}', " _
                                          & " control_url='{1}' " _
                                          & "WHERE control_id={2}", control_name, control_url, control_id)

      Dim iRecordsAffected As Integer = 0

      '
      ' Build the insert/update/delete query
      '
      Using MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

        MyDbCommand.Connection = DBConnectionMain
        MyDbCommand.CommandType = CommandType.Text
        MyDbCommand.CommandText = strSql

        SyncLock SyncLockMain
          iRecordsAffected = MyDbCommand.ExecuteNonQuery()
        End SyncLock

        MyDbCommand.Dispose()

      End Using

      strMessage = "UpdateNetCamControls() updated " & iRecordsAffected & " row(s)."
      Call WriteMessage(strMessage, MessageType.Debug)

      If iRecordsAffected > 0 Then
        Return True
      Else
        Return False
      End If

    Catch pEx As Exception
      Call ProcessError(pEx, "UpdateNetCamControls()")
      Return False
    Finally
      If bRefreshNetCamList = True Then
        RefreshNetCamList()
      End If
    End Try

  End Function

  ''' <summary>
  ''' Removes existing NetCam Controls stored in the database
  ''' </summary>
  ''' <param name="control_id"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function DeleteNetCamControls(ByVal control_id As Integer) As Boolean

    Dim strMessage As String = ""

    Try

      Select Case DBConnectionMain.State
        Case ConnectionState.Broken, ConnectionState.Closed
          strMessage = "Unable to complete database transaction because the database " _
                     & "connection has not been initialized."
          Throw New System.Exception(strMessage)
      End Select

      Dim iRecordsAffected As Integer = 0

      '
      ' Build the insert/update/delete query
      '
      Using MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

        MyDbCommand.Connection = DBConnectionMain
        MyDbCommand.CommandType = CommandType.Text
        MyDbCommand.CommandText = String.Format("DELETE FROM tblNetCamControls WHERE control_id={0}", control_id.ToString)

        SyncLock SyncLockMain
          iRecordsAffected = MyDbCommand.ExecuteNonQuery()
        End SyncLock

        MyDbCommand.Dispose()

      End Using

      strMessage = "DeleteNetCamType() removed " & iRecordsAffected & " row(s)."
      Call WriteMessage(strMessage, MessageType.Debug)

      If iRecordsAffected > 0 Then
        Return True
      Else
        Return False
      End If

      Return True

    Catch pEx As Exception
      Call ProcessError(pEx, "DeleteNetCamControls()")
      Return False
    Finally
      RefreshNetCamList()
    End Try

  End Function

  ''' <summary>
  ''' Load NetCam Controls
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub LoadNetCamControls()

    Try

      Dim ControlFiles As String() = {"Foscam[JPEG].txt", "Foscam[H.264].txt"}
      Dim netcam_type As Integer = 1

      For Each strFileName As String In ControlFiles

        Dim HTTPWebRequest As HttpWebRequest = System.Net.WebRequest.Create(String.Format("http://automatedhomeonline.com/HomeSeer3/hspi_netcam3/{0}", strFileName))
        HTTPWebRequest.Timeout = 1000 * 15

        Using response As HttpWebResponse = DirectCast(HTTPWebRequest.GetResponse(), HttpWebResponse)

          Using responseStream As Stream = response.GetResponseStream()
            ImportNetCamControls(netcam_type, responseStream)
            netcam_type += 1
          End Using

        End Using

      Next

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "LoadNetCamControls()")
    End Try

  End Sub

  ''' <summary>
  ''' Processes Upload File
  ''' </summary>
  ''' <param name="netcam_type"></param>
  ''' <param name="MyFileStream"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ImportNetCamControls(ByVal netcam_type As Integer, ByRef MyFileStream As System.IO.Stream) As Boolean

    Dim MyDataTable As New DataTable
    Dim MyHashTable As New Hashtable

    Try

      Dim bDeleted As Boolean = TruncateNetCamControls(netcam_type)

      Using sr As New System.IO.StreamReader(MyFileStream)

        Dim line As String = sr.ReadLine()

        Dim colMatches As MatchCollection

        While Not line Is Nothing

          colMatches = Regex.Matches(line, "^\[(?<NAME>(.+))\]\s+(?<URL>.+)\s*")

          If colMatches.Count > 0 Then

            For Each objMatch As Match In colMatches

              Dim strControlName As String = objMatch.Groups("NAME").Value
              Dim strControlURL As String = objMatch.Groups("URL").Value

              If MyHashTable.ContainsKey(strControlName) = False Then
                MyHashTable.Add(strControlName, "")

                Dim bInserted As Boolean = InsertNetCamControls(netcam_type, strControlName, strControlURL, False)

              End If

            Next

          End If

          line = sr.ReadLine()

        End While

      End Using

      Return True

    Catch pEx As Exception
      Call ProcessError(pEx, "ImportNetCamControls()")
      Return False
    Finally
      RefreshNetCamList()
    End Try

  End Function

  ''' <summary>
  ''' Truncates NetCam Controls
  ''' </summary>
  ''' <param name="netcam_type"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TruncateNetCamControls(ByVal netcam_type As Integer) As Boolean

    Dim strMessage As String = ""

    Try

      Select Case DBConnectionMain.State
        Case ConnectionState.Broken, ConnectionState.Closed
          strMessage = "Unable to complete database transaction because the database " _
                     & "connection has not been initialized."
          Throw New System.Exception(strMessage)
      End Select

      '
      ' Build the insert/update/delete query
      '
      Using MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

        MyDbCommand.Connection = DBConnectionMain
        MyDbCommand.CommandType = CommandType.Text
        MyDbCommand.CommandText = String.Format("DELETE FROM tblNetCamControls WHERE netcam_type='{0}'", netcam_type.ToString)

        Dim iRecordsAffected As Integer = 0
        SyncLock SyncLockMain
          iRecordsAffected = MyDbCommand.ExecuteNonQuery()
        End SyncLock

        strMessage = "TruncateNetCamControls() removed " & iRecordsAffected & " row(s)."
        Call WriteMessage(strMessage, MessageType.Debug)

        MyDbCommand.Dispose()

      End Using

      Return True

    Catch pEx As Exception
      Call ProcessError(pEx, "TruncateNetCamControls()")
      Return False
    Finally
      RefreshNetCamList()
    End Try

  End Function

#End Region

#Region "HSPI - Snapshots"

  ''' <summary>
  ''' Takes a single snapshot
  ''' </summary>
  ''' <param name="netcam_id"></param>
  ''' <remarks></remarks>
  Public Sub NetCamSnapshot(ByVal netcam_id As String)

    Try

      '
      ' Format the NetCam Id
      '
      Dim strNetCamId As String = netcam_id
      If Regex.IsMatch(strNetCamId, "NetCam\d\d\d") = False Then
        strNetCamId = String.Format("NetCam{0}", netcam_id.PadLeft(3, "0"))
      End If

      '
      ' Add snapshot commands if NetCam exists
      '
      If NetCams.ContainsKey(strNetCamId) = True Then
        Dim NetCamDevice As NetCamDevice = NetCams(strNetCamId)

        Dim NetCamCommand As New NetCamCommand("TakeSnapshot", 0, 0, 0, "")
        NetCamDevice.AddCommand(NetCamCommand)

      Else
        WriteMessage(String.Format("Unable to find {0}!", strNetCamId), MessageType.Warning)
      End If

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "NetCamSnapshot()")
    End Try

  End Sub

  ''' <summary>
  ''' Refreshes the latest snapshots
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub RefreshLatestSnapshots()

    Try

      SyncLock NetCams.SyncRoot

        For Each strKey As String In NetCams.Keys
          Dim NetCamDevice As NetCamDevice = NetCams(strKey)

          If NetCamDevice.CommandQueueCount = 0 Then
            Dim strSnapshotFilename As String = FixPath(String.Format("{0}\last_snapshot.jpg", NetCamDevice.SnapshotDirectory))
            Dim strThumbnailFilename As String = FixPath(String.Format("{0}\last_thumbnail.jpg", NetCamDevice.SnapshotDirectory))

            NetCamDevice.TakeSnapshot(strSnapshotFilename, strThumbnailFilename, 5)
          End If

        Next

      End SyncLock

    Catch pEx As Exception
      '
      ' Process Program Exception
      '
      Call ProcessError(pEx, "RefreshLatestSnapshots()")
    End Try

  End Sub

  ''' <summary>
  ''' Get the camera snapshot minimum width
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetCameraSnapshotMinWidth() As Integer

    Dim iMinWidth As Integer = 640

    Try

      Dim DirectoryInfo As New System.IO.DirectoryInfo(gSnapshotDirectory)
      Dim Files() As System.IO.FileInfo = DirectoryInfo.GetFiles("last_thumbnail.jpg", SearchOption.AllDirectories)

      For Each MyFileInfo As FileInfo In Files
        Dim img As Image = Image.FromFile(MyFileInfo.FullName)
        If img.Size.Width < iMinWidth Then iMinWidth = img.Size.Width
      Next
    Catch pEx As Exception

    End Try

    Return iMinWidth

  End Function

  ''' <summary>
  ''' Get the camera viewer filenames
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetCameraSnapshotViewer() As DataTable

    Dim FileNames As New DataTable

    Try

      FileNames.TableName = "tblSnapshotInfo"

      Dim ColumNames As String() = {"netcam_id", "netcam_name", "event_id", "creation_date", "snapshot_age", "snapshot_filename", "thumbnail_filename"}
      For Each ColumnName As String In ColumNames

        '
        ' Create the columns
        '
        Dim MyDataColumn As New DataColumn(ColumnName, GetType(String))

        '
        ' Add the columns to the DataTable's Columns collection
        '
        FileNames.Columns.Add(MyDataColumn)
      Next

      Dim DirectoryInfo As New System.IO.DirectoryInfo(gSnapshotDirectory)
      Dim Files() As System.IO.FileInfo = DirectoryInfo.GetFiles("last_thumbnail.jpg", SearchOption.AllDirectories)

      Dim comparer As IComparer = New FilenameComparer()
      Array.Sort(Files, comparer)

      For Each MyFileInfo As FileInfo In Files
        Dim CreationDate As DateTime = MyFileInfo.CreationTime
        Dim strNetCamName = "Unknown"
        Dim strEventId As String = 0

        Dim ThumbnailFileName As String = Regex.Replace(MyFileInfo.FullName, ".*images", "\images").Replace("\", "/")
        Dim SnapshotFileName As String = ThumbnailFileName.Replace("thumbnail", "snapshot")

        Dim strNetCamId As String = Regex.Match(ThumbnailFileName, "NetCam\d\d\d").Value
        If NetCams.ContainsKey(strNetCamId) Then
          Dim NetCamDevice As NetCamDevice = NetCams(strNetCamId)
          strNetCamName = NetCamDevice.Name
        End If

        Dim SnapshotInfo As SnapshotInfo = New SnapshotInfo(strNetCamId, strEventId, ThumbnailFileName, SnapshotFileName, CreationDate)

        Dim MyDataRow As DataRow = FileNames.NewRow()

        MyDataRow.Item("netcam_name") = strNetCamName
        MyDataRow.Item("netcam_id") = SnapshotInfo.NetCamId
        MyDataRow.Item("event_id") = SnapshotInfo.EventId
        MyDataRow.Item("creation_date") = SnapshotInfo.CreationDate
        MyDataRow.Item("snapshot_age") = SnapshotInfo.SnapshotAge
        MyDataRow.Item("snapshot_filename") = SnapshotInfo.SnapshotFileName
        MyDataRow.Item("thumbnail_filename") = SnapshotInfo.ThumbnailFileName

        FileNames.Rows.Add(MyDataRow)
      Next

    Catch pEx As Exception
      '
      ' Process Program Exception
      '
      Call ProcessError(pEx, "GetCameraSnapshotViewer()")
    End Try

    Return FileNames

  End Function

  ''' <summary>
  ''' Get the camera picture filenames
  ''' </summary>
  ''' <param name="strNetCamName"></param>
  ''' <param name="strEventId"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetCameraSnapshotInfo(ByVal strNetCamName As String, Optional ByVal strEventId As String = "") As DataTable

    Dim FileNames As New DataTable

    Try

      FileNames.TableName = "tblSnapshotInfo"

      Dim ColumNames As String() = {"netcam_id", "event_id", "creation_date", "snapshot_age", "snapshot_filename", "thumbnail_filename"}
      For Each ColumnName As String In ColumNames

        '
        ' Create the columns
        '
        Dim MyDataColumn As New DataColumn(ColumnName, GetType(String))

        '
        ' Add the columns to the DataTable's Columns collection
        '
        FileNames.Columns.Add(MyDataColumn)
      Next

      Dim DirectoryPath As String = FixPath(String.Format("{0}\{1}", gSnapshotDirectory, strNetCamName))
      If strEventId.Length > 0 Then
        DirectoryPath = FixPath(String.Format("{0}\{1}\{2}", gSnapshotDirectory, strNetCamName, strEventId))
      End If

      If Directory.Exists(DirectoryPath) = True Then
        Dim DirectoryInfo As New System.IO.DirectoryInfo(DirectoryPath)
        Dim Files() As System.IO.FileInfo = DirectoryInfo.GetFiles(String.Format("{0}_*_thumbnail.jpg", strNetCamName), SearchOption.AllDirectories)

        Dim comparer As IComparer = New DateComparer()
        Array.Sort(Files, comparer)

        For Each MyFileInfo As FileInfo In Files
          Dim CreationDate As DateTime = MyFileInfo.CreationTime
          strEventId = Regex.Match(MyFileInfo.FullName, "(Event\d+)").ToString

          If strEventId.Length > 0 Then
            Dim ThumbnailFileName As String = Regex.Replace(MyFileInfo.FullName, ".*images", "\images").Replace("\", "/")
            Dim SnapshotFileName As String = ThumbnailFileName.Replace("thumbnail", "snapshot")

            Dim SnapshotInfo As SnapshotInfo = New SnapshotInfo(strNetCamName, strEventId, ThumbnailFileName, SnapshotFileName, CreationDate)

            Dim MyDataRow As DataRow = FileNames.NewRow()
            MyDataRow.Item("netcam_id") = SnapshotInfo.NetCamId
            MyDataRow.Item("event_id") = SnapshotInfo.EventId
            MyDataRow.Item("creation_date") = SnapshotInfo.CreationDate
            MyDataRow.Item("snapshot_age") = SnapshotInfo.SnapshotAge
            MyDataRow.Item("snapshot_filename") = SnapshotInfo.SnapshotFileName
            MyDataRow.Item("thumbnail_filename") = SnapshotInfo.ThumbnailFileName
            FileNames.Rows.Add(MyDataRow)
          End If

        Next
      End If

    Catch pEx As Exception
      '
      ' Process Program Exception
      '
      Call ProcessError(pEx, "GetCameraSnapshotInfo()")
    End Try

    Return FileNames

  End Function

  ''' <summary>
  ''' Removes the snapshot directory
  ''' </summary>
  ''' <param name="netcam_id"></param>
  ''' <remarks></remarks>
  Public Sub DeleteNetCamSnapshotDir(ByVal netcam_id As String)

    Try
      '
      ' Format the NetCam Id
      '
      Dim strNetCamId As String = netcam_id
      If Regex.IsMatch(strNetCamId, "NetCam\d\d\d") = False Then
        strNetCamId = String.Format("NetCam{0}", netcam_id.PadLeft(3, "0"))
      End If

      '
      ' Remove the directory
      '
      Dim strDirectory As String = FixPath(String.Format("{0}/{1}", gSnapshotDirectory, strNetCamId))
      If Directory.Exists(strDirectory) = True Then
        WriteMessage(String.Format("Removing {0} because the camera is no longer defined.", strNetCamId), MessageType.Warning)
        Directory.Delete(strDirectory, True)
      End If

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "DeleteNetCamSnapshotDir()")
    End Try

  End Sub

  ''' <summary>
  ''' Purges snapshots from filesystem
  ''' </summary>
  ''' <param name="strNetCamName"></param>
  ''' <param name="strEventId"></param>
  ''' <remarks></remarks>
  Public Sub PurgeSnapshotEvent(ByVal strNetCamName As String, ByVal strEventId As String)

    Try

      Call WriteMessage("Entered PurgeSnapshotEvent() subroutine.", MessageType.Debug)

      If strEventId.Length > 0 Then
        '
        ' Purge a single event for camera
        '
        Dim SnapshotPath As String = FixPath(String.Format("{0}\{1}\{2}", gSnapshotDirectory, strNetCamName, strEventId))

        If Directory.Exists(SnapshotPath) = True Then
          WriteMessage(String.Format("Deleting {0}", SnapshotPath), MessageType.Debug)
          Directory.Delete(SnapshotPath, True)
        End If
      Else
        '
        ' Purge all events for selected Camera
        '
        Dim SnapshotPath As String = FixPath(String.Format("{0}\{1}", gSnapshotDirectory, strNetCamName))

        If Directory.Exists(SnapshotPath) = True Then

          Dim RootDirInfo As New IO.DirectoryInfo(SnapshotPath)

          For Each EventDirInfo As DirectoryInfo In RootDirInfo.GetDirectories
            If Directory.Exists(EventDirInfo.FullName) = True Then
              WriteMessage(String.Format("Deleting {0}", EventDirInfo.FullName), MessageType.Debug)
              Directory.Delete(EventDirInfo.FullName, True)
              Thread.Sleep(0)
            End If
          Next

        End If
      End If

    Catch pEx As Exception
      '
      ' Return message
      '
      ProcessError(pEx, "PurgeSnapshotEvent()")
    End Try

  End Sub

#End Region

#Region "HSPI - NetCam Events"

  ''' <summary>
  ''' GetNetCamEventSummary
  ''' </summary>
  ''' <param name="strNetCamName"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetNetCamEventSummary(ByVal strNetCamName As String) As SortedList

    Dim NetCamEvents As New SortedList
    If strNetCamName.Length = 0 Then Return NetCamEvents

    Try

      Dim strRootDir As String = FixPath(String.Format("{0}\{1}", gSnapshotDirectory, strNetCamName))
      If Directory.Exists(strRootDir) = True Then

        Dim RootDirInfo As New IO.DirectoryInfo(strRootDir)
        For Each EventDirInfo As DirectoryInfo In RootDirInfo.GetDirectories
          If Directory.Exists(EventDirInfo.FullName) = True Then
            Dim Files() As System.IO.FileInfo = EventDirInfo.GetFiles(String.Format("{0}_*_thumbnail.jpg", strNetCamName), SearchOption.AllDirectories)

            Dim strEventId As String = Regex.Match(EventDirInfo.FullName, "(Event\d+)").ToString
            If strEventId.Length > 0 Then
              If NetCamEvents.ContainsKey(strEventId) = False Then
                NetCamEvents.Add(strEventId, Files.Length)
              End If
            End If

          End If

        Next

      End If

    Catch pEx As Exception
      '
      ' Process Program Exception
      '
      Call ProcessError(pEx, "GetNetCamEventSummary()")
    End Try

    Return NetCamEvents

  End Function

  ''' <summary>
  ''' Gets the NetCam Events from the underlying database
  ''' </summary>
  ''' <param name="netcam_id"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetNetCamEvents(ByVal netcam_id As Integer) As DataTable

    Dim ResultsDT As New DataTable
    Dim strMessage As String = ""

    strMessage = "Entered GetNetCamEvents() function."
    Call WriteMessage(strMessage, MessageType.Debug)

    Try
      '
      ' Make sure the datbase is open before attempting to use it
      '
      Select Case DBConnectionMain.State
        Case ConnectionState.Broken, ConnectionState.Closed
          strMessage = "Unable to complete database query because the database " _
                     & "connection has not been initialized."
          Throw New System.Exception(strMessage)
      End Select

      Dim strSQL As String = String.Format("SELECT * FROM tblNetCamEvents WHERE netcam_id = {0}", netcam_id.ToString)

      '
      ' Initialize the command object
      '
      Dim MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

      MyDbCommand.Connection = DBConnectionMain
      MyDbCommand.CommandType = CommandType.Text
      MyDbCommand.CommandText = strSQL

      '
      ' Initialize the dataset, then populate it
      '
      Dim MyDS As DataSet = New DataSet

      Dim MyDA As System.Data.IDbDataAdapter = New SQLiteDataAdapter(MyDbCommand)
      MyDA.SelectCommand = MyDbCommand

      SyncLock SyncLockMain
        MyDA.Fill(MyDS)
      End SyncLock

      '
      ' Get our DataTable
      '
      Dim MyDT As DataTable = MyDS.Tables(0)

      '
      ' Get record count
      '
      Dim iRecordCount As Integer = MyDT.Rows.Count

      If iRecordCount > 0 Then
        '
        ' Build field names
        '
        Dim iFieldCount As Integer = MyDS.Tables(0).Columns.Count() - 1
        For iFieldNum As Integer = 0 To iFieldCount
          '
          ' Create the columns
          '
          Dim ColumnName As String = MyDT.Columns.Item(iFieldNum).ColumnName
          Dim MyDataColumn As New DataColumn(ColumnName, GetType(String))

          '
          ' Add the columns to the DataTable's Columns collection
          '
          ResultsDT.Columns.Add(MyDataColumn)
        Next

        '
        ' Let's output our records	
        '
        Dim i As Integer = 0
        For i = 0 To iRecordCount - 1
          '
          ' Create the rows
          '
          Dim dr As DataRow
          dr = ResultsDT.NewRow()
          For iFieldNum As Integer = 0 To iFieldCount
            dr(iFieldNum) = MyDT.Rows(i)(iFieldNum)
          Next
          ResultsDT.Rows.Add(dr)
        Next

      End If

    Catch pEx As Exception
      '
      ' Process Exception
      '
      Call ProcessError(pEx, "GetNetCamEvents()")

    End Try

    Return ResultsDT

  End Function

  ''' <summary>
  ''' Takes an event snapshot
  ''' </summary>
  ''' <param name="netcam_id"></param>
  ''' <param name="iSnapshotCount"></param>
  ''' <param name="iWaitSeconds"></param>
  ''' <param name="strEmailRecipient"></param>
  ''' <param name="iEmailAttachment"></param>
  ''' <param name="bEmailed"></param>
  ''' <param name="bCompressed"></param>
  ''' <param name="bUploaded"></param>
  ''' <param name="bArchived"></param>
  ''' <remarks></remarks>
  Public Sub NetCamEventSnapshot(ByVal netcam_id As String, _
                                 ByVal iSnapshotCount As Integer, _
                                 ByVal iWaitSeconds As Integer, _
                                 ByVal strEmailRecipient As String, _
                                 ByVal iEmailAttachment As Integer, _
                                 ByVal bEmailed As Boolean, _
                                 ByVal bCompressed As Boolean, _
                                 ByVal bUploaded As Boolean, _
                                 ByVal bArchived As Boolean)

    Try
      '
      ' Format the NetCam Id
      '
      Dim strNetCamId As String = netcam_id
      If Regex.IsMatch(strNetCamId, "NetCam\d\d\d") = False Then
        strNetCamId = String.Format("NetCam{0}", netcam_id.PadLeft(3, "0"))
      End If

      '
      ' Start NetCam Event
      '
      Dim iNetCamId As Integer = Val(Regex.Match(strNetCamId, "\d\d\d").Value)
      Dim iEmailed As Integer = IIf(bEmailed = True, 1, 0)
      Dim iUploaded As Integer = IIf(bUploaded = True, 1, 0)
      Dim iCompressed As Integer = IIf(bCompressed = True, 1, 0)
      Dim iArchived As Integer = IIf(bArchived = True, 1, 0)

      Dim NetCamEvent As New NetCamEvent(iNetCamId, _
                                         iSnapshotCount, _
                                         strEmailRecipient, _
                                         iEmailAttachment, _
                                         iEmailed, _
                                         iCompressed, _
                                         iUploaded, _
                                         iArchived)

      Dim event_id As Integer = StartNetCamEvent(NetCamEvent)

      '
      ' Add snapshot commands if NetCam exists
      '
      If NetCams.ContainsKey(strNetCamId) = True Then
        Dim NetCamDevice As NetCamDevice = NetCams(strNetCamId)

        NetCamDevice.AddCommand(New NetCamCommand("StartEvent", 0, event_id, 0, ""))
        For i As Integer = 1 To iSnapshotCount
          Dim NetCamCommand As New NetCamCommand("TakeEventSnapshot", iWaitSeconds, event_id, i, "")
          NetCamDevice.AddCommand(NetCamCommand)
        Next
        NetCamDevice.AddCommand(New NetCamCommand("EndEvent", 0, event_id, 0, ""))

      Else
        WriteMessage(String.Format("Unable to find {0}!", strNetCamId), MessageType.Warning)
      End If

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "NetCamSnapshot()")
    End Try

  End Sub

#End Region

#Region "HSPI - NetCam Types"

  ''' <summary>
  ''' Gets the NetCam Types from the underlying database
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetNetCamTypesFromDB() As DataTable

    Dim ResultsDT As New DataTable
    Dim strMessage As String = ""

    strMessage = "Entered GetNetCamTypesFromDB() function."
    Call WriteMessage(strMessage, MessageType.Debug)

    Try
      '
      ' Make sure the datbase is open before attempting to use it
      '
      Select Case DBConnectionMain.State
        Case ConnectionState.Broken, ConnectionState.Closed
          strMessage = "Unable to complete database query because the database " _
                     & "connection has not been initialized."
          Throw New System.Exception(strMessage)
      End Select

      Dim strSQL As String = String.Format("SELECT netcam_type, netcam_vendor, netcam_model, snapshot_path, videostream_path, " & _
                                           "netcam_vendor || ' [' || netcam_model || ']' as netcam_name FROM tblNetCamTypes")

      '
      ' Initialize the command object
      '
      Dim MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

      MyDbCommand.Connection = DBConnectionMain
      MyDbCommand.CommandType = CommandType.Text
      MyDbCommand.CommandText = strSQL

      '
      ' Initialize the dataset, then populate it
      '
      Dim MyDS As DataSet = New DataSet

      Dim MyDA As System.Data.IDbDataAdapter = New SQLiteDataAdapter(MyDbCommand)
      MyDA.SelectCommand = MyDbCommand

      SyncLock SyncLockMain
        MyDA.Fill(MyDS)
      End SyncLock

      '
      ' Get our DataTable
      '
      Dim MyDT As DataTable = MyDS.Tables(0)

      '
      ' Get record count
      '
      Dim iRecordCount As Integer = MyDT.Rows.Count

      If iRecordCount > 0 Then
        '
        ' Build field names
        '
        Dim iFieldCount As Integer = MyDS.Tables(0).Columns.Count() - 1
        For iFieldNum As Integer = 0 To iFieldCount
          '
          ' Create the columns
          '
          Dim ColumnName As String = MyDT.Columns.Item(iFieldNum).ColumnName
          Dim MyDataColumn As New DataColumn(ColumnName, GetType(String))

          '
          ' Add the columns to the DataTable's Columns collection
          '
          ResultsDT.Columns.Add(MyDataColumn)
        Next

        '
        ' Let's output our records	
        '
        Dim i As Integer = 0
        For i = 0 To iRecordCount - 1
          '
          ' Create the rows
          '
          Dim dr As DataRow
          dr = ResultsDT.NewRow()
          For iFieldNum As Integer = 0 To iFieldCount
            dr(iFieldNum) = MyDT.Rows(i)(iFieldNum)
          Next
          ResultsDT.Rows.Add(dr)
        Next

      End If

    Catch pEx As Exception
      '
      ' Process Exception
      '
      Call ProcessError(pEx, "GetNetCamTypesFromDB()")

    End Try

    Return ResultsDT

  End Function

  ''' <summary>
  ''' Inserts a new NetCam Type into the database
  ''' </summary>
  ''' <param name="netcam_vendor"></param>
  ''' <param name="netcam_model"></param>
  ''' <param name="snapshot_path"></param>
  ''' <param name="videostream_path"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function InsertNetCamType(ByVal netcam_vendor As String, _
                                   ByVal netcam_model As String, _
                                   ByVal snapshot_path As String, _
                                   ByVal videostream_path As String) As Integer

    Dim strMessage As String = ""
    Dim netcam_type As Integer = 0

    Try

      Select Case DBConnectionMain.State
        Case ConnectionState.Broken, ConnectionState.Closed
          strMessage = "Unable to complete database transaction because the database " _
                     & "connection has not been initialized."
          Throw New System.Exception(strMessage)
      End Select

      If netcam_vendor.Length = 0 Or netcam_model.Length = 0 Or snapshot_path.Length = 0 Then
        Throw New Exception("One or more required fields are empty.  Unable to insert new NetCam type into the database.")
      End If

      Dim strSQL As String = String.Format("INSERT INTO tblNetCamTypes (" _
                                           & " netcam_vendor, netcam_model, snapshot_path, videostream_path" _
                                           & ") VALUES (" _
                                           & "'{0}', '{1}', '{2}', '{3}' );", _
                                           netcam_vendor, netcam_model, snapshot_path, videostream_path)
      strSQL &= "SELECT last_insert_rowid();"

      Using dbcmd As DbCommand = DBConnectionMain.CreateCommand()

        dbcmd.Connection = DBConnectionMain
        dbcmd.CommandType = CommandType.Text
        dbcmd.CommandText = strSQL

        SyncLock SyncLockMain
          netcam_type = dbcmd.ExecuteScalar()
        End SyncLock

        dbcmd.Dispose()

      End Using

    Catch pEx As Exception
      Call ProcessError(pEx, "InsertNetCamType()")
      Return False
    Finally
      RefreshNetCamList()
    End Try

    Return netcam_type

  End Function

  ''' <summary>
  ''' Updates existing NetCam Type stored in the database
  ''' </summary>
  ''' <param name="netcam_type"></param>
  ''' <param name="netcam_vendor"></param>
  ''' <param name="netcam_model"></param>
  ''' <param name="snapshot_path"></param>
  ''' <param name="videostream_path"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function UpdateNetCamType(ByVal netcam_type As Integer, _
                                   ByVal netcam_vendor As String, _
                                   ByVal netcam_model As String, _
                                   ByVal snapshot_path As String, _
                                   ByVal videostream_path As String) As Boolean

    Dim strMessage As String = ""

    Try

      Select Case DBConnectionMain.State
        Case ConnectionState.Broken, ConnectionState.Closed
          strMessage = "Unable to complete database transaction because the database " _
                     & "connection has not been initialized."
          Throw New System.Exception(strMessage)
      End Select

      Dim strSql As String = String.Format("UPDATE tblNetCamTypes SET " _
                                          & " netcam_vendor='{0}', " _
                                          & " netcam_model='{1}', " _
                                          & " snapshot_path='{2}'," _
                                          & " videostream_path='{3}' " _
                                          & "WHERE netcam_type={4}", netcam_vendor, netcam_model, snapshot_path, videostream_path, netcam_type)

      Dim iRecordsAffected As Integer = 0

      '
      ' Build the insert/update/delete query
      '
      Using MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

        MyDbCommand.Connection = DBConnectionMain
        MyDbCommand.CommandType = CommandType.Text
        MyDbCommand.CommandText = strSql

        SyncLock SyncLockMain
          iRecordsAffected = MyDbCommand.ExecuteNonQuery()
        End SyncLock

        strMessage = "UpdateNetCamType() updated " & iRecordsAffected & " row(s)."
        Call WriteMessage(strMessage, MessageType.Debug)

        MyDbCommand.Dispose()

      End Using

      If iRecordsAffected > 0 Then
        Return True
      Else
        Return False
      End If

    Catch pEx As Exception
      Call ProcessError(pEx, "UpdateNetCamType()")
      Return False
    Finally
      RefreshNetCamList()
    End Try

  End Function

  ''' <summary>
  ''' Removes existing NetCam Type stored in the database
  ''' </summary>
  ''' <param name="netcam_type"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function DeleteNetCamType(ByVal netcam_type As Integer) As Boolean

    Dim strMessage As String = ""

    Try

      Select Case DBConnectionMain.State
        Case ConnectionState.Broken, ConnectionState.Closed
          strMessage = "Unable to complete database transaction because the database " _
                     & "connection has not been initialized."
          Throw New System.Exception(strMessage)
      End Select

      Dim iRecordsAffected As Integer = 0

      '
      ' Build the insert/update/delete query
      '
      Using MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

        MyDbCommand.Connection = DBConnectionMain
        MyDbCommand.CommandType = CommandType.Text
        MyDbCommand.CommandText = String.Format("DELETE FROM tblNetCamTypes WHERE netcam_type={0}", netcam_type.ToString)


        SyncLock SyncLockMain
          iRecordsAffected = MyDbCommand.ExecuteNonQuery()
        End SyncLock

        strMessage = "DeleteNetCamType() removed " & iRecordsAffected & " row(s)."
        Call WriteMessage(strMessage, MessageType.Debug)

        MyDbCommand.Dispose()

      End Using

      If iRecordsAffected > 0 Then
        Return True
      Else
        Return False
      End If

      Return True

    Catch pEx As Exception
      Call ProcessError(pEx, "DeleteNetCamType()")
      Return False
    Finally
      RefreshNetCamList()
    End Try

  End Function

#End Region

#Region "HSPI - NetCam Devices"

  ''' <summary>
  ''' Gets the NetCam Id from the database
  ''' </summary>
  ''' <param name="ip_address"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetNetCamDeviceId(ByVal ip_address As String) As Integer

    Dim netcam_id As Integer = 0

    Try

      Dim strSQL As String = String.Format("SELECT netcam_id FROM tblNetCamDevices WHERE netcam_address='{0}'", ip_address)

      '
      ' Execute the data reader
      '
      Using MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

        MyDbCommand.Connection = DBConnectionMain
        MyDbCommand.CommandType = CommandType.Text
        MyDbCommand.CommandText = strSQL

        SyncLock SyncLockMain
          Dim dtrResults As IDataReader = MyDbCommand.ExecuteReader()

          '
          ' Process the results
          '
          If dtrResults.Read() Then
            netcam_id = dtrResults("netcam_id")
          End If

          dtrResults.Close()
        End SyncLock

        MyDbCommand.Dispose()

      End Using

    Catch pEx As Exception

    End Try

    Return netcam_id

  End Function

  ''' <summary>
  ''' Gets the NetCam Devices from the underlying database
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetNetCamDevicesFromDB() As DataTable

    Dim ResultsDT As New DataTable
    Dim strMessage As String = ""

    strMessage = "Entered GetNetCamDevicesFromDB() function."
    Call WriteMessage(strMessage, MessageType.Debug)

    Try
      '
      ' Make sure the datbase is open before attempting to use it
      '
      Select Case DBConnectionMain.State
        Case ConnectionState.Broken, ConnectionState.Closed
          strMessage = "Unable to complete database query because the database " _
                     & "connection has not been initialized."
          Throw New System.Exception(strMessage)
      End Select

      Dim strSQL As String = String.Format("SELECT * FROM tblNetCamDevices")

      '
      ' Initialize the command object
      '
      Dim MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

      MyDbCommand.Connection = DBConnectionMain
      MyDbCommand.CommandType = CommandType.Text
      MyDbCommand.CommandText = strSQL

      '
      ' Initialize the dataset, then populate it
      '
      Dim MyDS As DataSet = New DataSet

      Dim MyDA As System.Data.IDbDataAdapter = New SQLiteDataAdapter(MyDbCommand)
      MyDA.SelectCommand = MyDbCommand

      SyncLock SyncLockMain
        MyDA.Fill(MyDS)
      End SyncLock

      '
      ' Get our DataTable
      '
      Dim MyDT As DataTable = MyDS.Tables(0)

      '
      ' Get record count
      '
      Dim iRecordCount As Integer = MyDT.Rows.Count

      If iRecordCount > 0 Then
        '
        ' Build field names
        '
        Dim iFieldCount As Integer = MyDS.Tables(0).Columns.Count() - 1
        For iFieldNum As Integer = 0 To iFieldCount
          '
          ' Create the columns
          '
          Dim ColumnName As String = MyDT.Columns.Item(iFieldNum).ColumnName
          Dim MyDataColumn As New DataColumn(ColumnName, GetType(String))

          '
          ' Add the columns to the DataTable's Columns collection
          '
          ResultsDT.Columns.Add(MyDataColumn)
        Next

        '
        ' Let's output our records	
        '
        Dim i As Integer = 0
        For i = 0 To iRecordCount - 1
          '
          ' Create the rows
          '
          Dim dr As DataRow
          dr = ResultsDT.NewRow()
          For iFieldNum As Integer = 0 To iFieldCount
            dr(iFieldNum) = MyDT.Rows(i)(iFieldNum)
          Next
          ResultsDT.Rows.Add(dr)
        Next

      End If

    Catch pEx As Exception
      '
      ' Process Exception
      '
      Call ProcessError(pEx, "GetNetCamDevicesFromDB()")

    End Try

    Return ResultsDT

  End Function

  ''' <summary>
  ''' Inserts a new NetCam Device into the database
  ''' </summary>
  ''' <param name="netcam_name"></param>
  ''' <param name="netcam_address"></param>
  ''' <param name="netcam_port"></param>
  ''' <param name="netcam_type"></param>
  ''' <param name="auth_user"></param>
  ''' <param name="auth_pass"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function InsertNetCamDevice(ByVal netcam_name As String, _
                                     ByVal netcam_address As String, _
                                     ByVal netcam_port As Integer, _
                                     ByVal netcam_type As Integer, _
                                     ByVal auth_user As String, _
                                     ByVal auth_pass As String) As Integer

    Dim strMessage As String = ""
    Dim netcam_id As Integer = 0

    Try

      Select Case DBConnectionMain.State
        Case ConnectionState.Broken, ConnectionState.Closed
          strMessage = "Unable to complete database transaction because the database " _
                     & "connection has not been initialized."
          Throw New System.Exception(strMessage)
      End Select

      If netcam_name.Length = 0 Or netcam_address.Length = 0 Or netcam_port = 0 Or netcam_type = 0 Then
        Throw New Exception("One or more required fields are empty.  Unable to insert new NetCam device into the database.")
      End If

      auth_pass = hs.EncryptString(auth_pass, "&Cul8r#1")
      Dim strSQL As String = String.Format("INSERT INTO tblNetCamDevices (" _
                                           & " netcam_name, netcam_address, netcam_port, netcam_type, auth_user, auth_pass" _
                                           & ") VALUES (" _
                                           & "'{0}', '{1}', {2}, {3}, '{4}', '{5}' );", _
                                           netcam_name, netcam_address, netcam_port, netcam_type, auth_user, auth_pass)
      strSQL &= "SELECT last_insert_rowid();"

      Using dbcmd As DbCommand = DBConnectionMain.CreateCommand()

        dbcmd.Connection = DBConnectionMain
        dbcmd.CommandType = CommandType.Text
        dbcmd.CommandText = strSQL

        SyncLock SyncLockMain
          netcam_id = dbcmd.ExecuteScalar()
        End SyncLock

        dbcmd.Dispose()

      End Using

    Catch pEx As Exception
      Call ProcessError(pEx, "InsertNetCamDevice()")
      Return False
    Finally
      RefreshNetCamList()
    End Try

    Return netcam_id

  End Function

  ''' <summary>
  ''' Updates existing NetCam Profile stored in the database
  ''' </summary>
  ''' <param name="netcam_id"></param>
  ''' <param name="netcam_name"></param>
  ''' <param name="netcam_address"></param>
  ''' <param name="netcam_port"></param>
  ''' <param name="netcam_type"></param>
  ''' <param name="auth_user"></param>
  ''' <param name="auth_pass"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function UpdateNetCamDevice(ByVal netcam_id As Integer, _
                                     ByVal netcam_name As String, _
                                     ByVal netcam_address As String, _
                                     ByVal netcam_port As Integer, _
                                     ByVal netcam_type As Integer, _
                                     ByVal auth_user As String, _
                                     ByVal auth_pass As String) As Boolean

    Dim strMessage As String = ""

    Try

      Select Case DBConnectionMain.State
        Case ConnectionState.Broken, ConnectionState.Closed
          strMessage = "Unable to complete database transaction because the database " _
                     & "connection has not been initialized."
          Throw New System.Exception(strMessage)
      End Select

      If netcam_name.Length = 0 Or netcam_address.Length = 0 Or netcam_port = 0 Then
        Throw New Exception("One or more required fields are empty.  Unable to save NetCam profile update to database.")
      End If

      Dim strSql As String = ""

      If auth_user.Length > 0 And auth_pass.Length > 0 Then

        auth_pass = hs.EncryptString(auth_pass, "&Cul8r#1")
        strSql = String.Format("UPDATE tblNetCamDevices SET " _
                            & " netcam_name='{0}', " _
                            & " netcam_address='{1}', " _
                            & " netcam_port={2}," _
                            & " netcam_type={3}," _
                            & " auth_user='{4}'," _
                            & " auth_pass='{5}' " _
                            & "WHERE netcam_id={6}", netcam_name, netcam_address, netcam_port.ToString, netcam_type.ToString, auth_user, auth_pass, netcam_id.ToString)

      ElseIf auth_user.Length > 0 Then

        strSql = String.Format("UPDATE tblNetCamDevices SET " _
                            & " netcam_name='{0}', " _
                            & " netcam_address='{1}', " _
                            & " netcam_port={2}," _
                            & " netcam_type={3}," _
                            & " auth_user='{4}' " _
                            & "WHERE netcam_id={5}", netcam_name, netcam_address, netcam_port.ToString, netcam_type.ToString, auth_user, netcam_id.ToString)

      Else

        auth_user = ""
        auth_pass = ""

        strSql = String.Format("UPDATE tblNetCamDevices SET " _
                            & " netcam_name='{0}', " _
                            & " netcam_address='{1}', " _
                            & " netcam_port={2}," _
                            & " netcam_type={3}," _
                            & " auth_user='{4}'," _
                            & " auth_pass='{5}' " _
                            & "WHERE netcam_id={6}", netcam_name, netcam_address, netcam_port.ToString, netcam_type.ToString, auth_user, auth_pass, netcam_id.ToString)

      End If

      Dim iRecordsAffected As Integer = 0

      '
      ' Build the insert/update/delete query
      '
      Using MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

        MyDbCommand.Connection = DBConnectionMain
        MyDbCommand.CommandType = CommandType.Text
        MyDbCommand.CommandText = strSql


        SyncLock SyncLockMain
          iRecordsAffected = MyDbCommand.ExecuteNonQuery()
        End SyncLock

        MyDbCommand.Dispose()

      End Using

      strMessage = "UpdateNetCamDevice() updated " & iRecordsAffected & " row(s)."
      Call WriteMessage(strMessage, MessageType.Debug)

      If iRecordsAffected > 0 Then
        Return True
      Else
        Return False
      End If

    Catch pEx As Exception
      Call ProcessError(pEx, "UpdateNetCamDevice()")
      Return False
    Finally
      RefreshNetCamList()
    End Try

  End Function

  ''' <summary>
  ''' Removes existing NetCam Profile stored in the database
  ''' </summary>
  ''' <param name="netcam_id"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function DeleteNetCamDevice(ByVal netcam_id As Integer) As Boolean

    Dim strMessage As String = ""

    Try

      Select Case DBConnectionMain.State
        Case ConnectionState.Broken, ConnectionState.Closed
          strMessage = "Unable to complete database transaction because the database " _
                     & "connection has not been initialized."
          Throw New System.Exception(strMessage)
      End Select

      Dim iRecordsAffected As Integer = 0

      '
      ' Build the insert/update/delete query
      '
      Using MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

        MyDbCommand.Connection = DBConnectionMain
        MyDbCommand.CommandType = CommandType.Text
        MyDbCommand.CommandText = String.Format("DELETE FROM tblNetCamDevices WHERE netcam_id={0}", netcam_id.ToString)

        SyncLock SyncLockMain
          iRecordsAffected = MyDbCommand.ExecuteNonQuery()
        End SyncLock

        MyDbCommand.Dispose()

      End Using

      strMessage = "DeleteNetCamDevice() removed " & iRecordsAffected & " row(s)."
      Call WriteMessage(strMessage, MessageType.Debug)

      If iRecordsAffected > 0 Then
        '
        ' Remove the camera snapshot directory
        '
        DeleteNetCamSnapshotDir(netcam_id.ToString)

        Return True
      Else
        Return False
      End If

    Catch pEx As Exception
      Call ProcessError(pEx, "DeleteNetCamDevice()")
      Return False
    Finally
      RefreshNetCamList()
    End Try

  End Function

#End Region

#Region "HSPI - Misc"

  ''' <summary>
  ''' Returns the FFmpeg Install Status
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetFFmpegStatus() As String

    Try

      Dim strProgram As String = FixPath(String.Format("{0}\{1}", HSAppPath, "ffmpeg.exe"))
      If File.Exists(strProgram) = True Then
        Return "Installed"
      Else
        Return "Not Installed"
      End If

    Catch pEx As Exception
      Return "Not Installed"
    End Try

  End Function

  ''' <summary>
  ''' Returns number of NetCams
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetNetCamCount() As Integer

    Try
      Return NetCams.Count
    Catch pEx As Exception
      Return 0
    End Try

  End Function

  ''' <summary>
  ''' Returns number of snapshots in directory
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetSnapshotCount() As Integer

    Try
      Return Directory.GetFiles(gSnapshotDirectory, "*_snapshot.jpg", SearchOption.AllDirectories).Length()
    Catch pEx As Exception
      Return 0
    End Try

  End Function

  ''' <summary>
  ''' Alarm Trigger
  ''' </summary>
  ''' <param name="strRemoteIP"></param>
  ''' <remarks></remarks>
  Public Sub AlarmTrigger(ByVal strRemoteIP As String)

    Try

      WriteMessage(String.Format("Alarm motion trigger received from {0}!", strRemoteIP), MessageType.Debug)

      SyncLock NetCams.SyncRoot

        For Each strNetCamId As String In NetCams.Keys
          Dim NetCamDevice As NetCamDevice = NetCams(strNetCamId)

          If NetCamDevice.Address = strRemoteIP Then
            'UltraNetCamUltraNetCam Alarm TriggerNetCam001
            Dim strTrigger As String = String.Format("{0},{1}", "Alarm Trigger", strNetCamId)
            hspi_plugin.CheckTrigger(IFACE_NAME, NetCamTriggers.AlarmTrigger, -1, strTrigger)

            Exit For
          End If

        Next

      End SyncLock

    Catch pEx As Exception
      '
      ' Process Program Exception
      '
      Call ProcessError(pEx, "AlarmTrigger()")
    End Try

  End Sub

  ''' <summary>
  ''' Execute Raw SQL
  ''' </summary>
  ''' <param name="strSQL"></param>
  ''' <param name="iRecordCount"></param>
  ''' <param name="iPageSize"></param>
  ''' <param name="iPageCount"></param>
  ''' <param name="iPageCur"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ExecuteSQL(ByVal strSQL As String, _
                             ByRef iRecordCount As Integer, _
                             ByVal iPageSize As Integer, _
                             ByRef iPageCount As Integer, _
                             ByRef iPageCur As Integer) As DataTable

    Dim ResultsDT As New DataTable
    Dim strMessage As String = ""

    strMessage = "Entered ExecuteSQL() function."
    Call WriteMessage(strMessage, MessageType.Debug)

    Try
      '
      ' Make sure the datbase is open before attempting to use it
      '
      Select Case DBConnectionMain.State
        Case ConnectionState.Broken, ConnectionState.Closed
          strMessage = "Unable to complete database query because the database " _
                     & "connection has not been initialized."
          Throw New System.Exception(strMessage)
      End Select

      '
      ' Determine Requested database action
      '
      If strSQL.StartsWith("SELECT", StringComparison.CurrentCultureIgnoreCase) Then

        '
        ' Populate the DataSet
        '

        '
        ' Initialize the command object
        '
        Dim MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

        MyDbCommand.Connection = DBConnectionMain
        MyDbCommand.CommandType = CommandType.Text
        MyDbCommand.CommandText = strSQL

        '
        ' Initialize the dataset, then populate it
        '
        Dim MyDS As DataSet = New DataSet

        Dim MyDA As System.Data.IDbDataAdapter = New SQLiteDataAdapter(MyDbCommand)
        MyDA.SelectCommand = MyDbCommand

        SyncLock SyncLockMain
          MyDA.Fill(MyDS)
        End SyncLock

        '
        ' Get our DataTable
        '
        Dim MyDT As DataTable = MyDS.Tables(0)

        '
        ' Get record count
        '
        iRecordCount = MyDT.Rows.Count

        If iRecordCount > 0 Then
          '
          ' Determine the number of pages available
          '
          iPageSize = IIf(iPageSize <= 0, 1, iPageSize)
          iPageCount = iRecordCount \ iPageSize
          If iRecordCount Mod iPageSize > 0 Then
            iPageCount += 1
          End If

          '
          ' Find starting and ending record
          '
          Dim nStart As Integer = iPageSize * (iPageCur - 1)
          Dim nEnd As Integer = nStart + iPageSize - 1
          If nEnd > iRecordCount - 1 Then
            nEnd = iRecordCount - 1
          End If

          '
          ' Build field names
          '
          Dim iFieldCount As Integer = MyDS.Tables(0).Columns.Count() - 1
          For iFieldNum As Integer = 0 To iFieldCount
            '
            ' Create the columns
            '
            Dim ColumnName As String = MyDT.Columns.Item(iFieldNum).ColumnName
            Dim MyDataColumn As New DataColumn(ColumnName, GetType(String))

            '
            ' Add the columns to the DataTable's Columns collection
            '
            ResultsDT.Columns.Add(MyDataColumn)
          Next

          '
          ' Let's output our records	
          '
          Dim i As Integer = 0
          For i = nStart To nEnd
            'Add some rows
            Dim dr As DataRow
            dr = ResultsDT.NewRow()
            For iFieldNum As Integer = 0 To iFieldCount
              dr(iFieldNum) = MyDT.Rows(i)(iFieldNum)
            Next
            ResultsDT.Rows.Add(dr)
          Next

          '
          ' Make sure current page count is valid
          '
          If iPageCur > iPageCount Then iPageCur = iPageCount

        Else
          '
          ' Query succeeded, but returned 0 records
          '
          strMessage = "Your query executed and returned 0 record(s)."
          Call WriteMessage(strMessage, MessageType.Debug)

        End If

      Else
        '
        ' Execute query (does not return recordset)
        '
        Try
          '
          ' Build the insert/update/delete query
          '
          Dim MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

          MyDbCommand.Connection = DBConnectionMain
          MyDbCommand.CommandType = CommandType.Text
          MyDbCommand.CommandText = strSQL

          Dim iRecordsAffected As Integer = 0
          SyncLock SyncLockMain
            iRecordsAffected = MyDbCommand.ExecuteNonQuery()
          End SyncLock

          strMessage = "Your query executed and affected " & iRecordsAffected & " row(s)."
          Call WriteMessage(strMessage, MessageType.Debug)

          MyDbCommand.Dispose()

        Catch pEx As Common.DbException
          '
          ' Process Database Error
          '
          strMessage = "Your query failed for the following reason(s):  "
          strMessage &= "[Error Source: " & pEx.Source & "] " _
                      & "[Error Number: " & pEx.ErrorCode & "] " _
                      & "[Error Desciption:  " & pEx.Message & "] "
          Call WriteMessage(strMessage, MessageType.Error)
        End Try

      End If

      Call WriteMessage("SQL: " & strSQL, MessageType.Debug)
      Call WriteMessage("Record Count: " & iRecordCount, MessageType.Debug)
      Call WriteMessage("Page Count: " & iPageCount, MessageType.Debug)
      Call WriteMessage("Page Current: " & iPageCur, MessageType.Debug)

    Catch pEx As Exception
      '
      ' Error:  Query error
      '
      strMessage = "Your query failed for the following reason:  " _
                  & Err.Source & " function/subroutine:  [" _
                  & Err.Number & " - " & Err.Description _
                  & "]"

      Call ProcessError(pEx, "ExecuteSQL()")

    End Try

    Return ResultsDT

  End Function

  ''' <summary>
  ''' Gets plug-in setting from INI file
  ''' </summary>
  ''' <param name="strSection"></param>
  ''' <param name="strKey"></param>
  ''' <param name="strValueDefault"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetSetting(ByVal strSection As String, _
                             ByVal strKey As String, _
                             ByVal strValueDefault As String) As String

    Dim strMessage As String = ""

    Try
      strMessage = "Entered GetSetting() function."
      Call WriteMessage(strMessage, MessageType.Debug)

      '
      ' Get the ini settings
      '
      Dim strValue As String = hs.GetINISetting(strSection, strKey, strValueDefault, gINIFile)

      strMessage = String.Format("Section: {0}, Key: {1}, Value: {2}", strSection, strKey, strValue)
      Call WriteMessage(strMessage, MessageType.Debug)

      '
      ' Check to see if we need to decrypt the data
      '
      If strSection = "FTPClient" And strKey = "UserPass" Then
        strValue = hs.DecryptString(strValue, "&Cul8r#1")
      ElseIf strKey = "UserPass" Then
        strValue = hs.DecryptString(strValue, "&Cul8r#1")
      End If

      Return strValue

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "GetSetting()")
      Return ""
    End Try

  End Function

  ''' <summary>
  ''' Saves plug-in setting to INI file
  ''' </summary>
  ''' <param name="strSection"></param>
  ''' <param name="strKey"></param>
  ''' <param name="strValue"></param>
  ''' <remarks></remarks>
  Public Sub SaveSetting(ByVal strSection As String, _
                         ByVal strKey As String, _
                         ByVal strValue As String)

    Dim strMessage As String = ""

    Try
      strMessage = "Entered SaveSetting() subroutine."
      Call WriteMessage(strMessage, MessageType.Debug)

      '
      ' Check to see if we need to encrypt the data
      '
      If strKey = "UserPass" Then
        If strValue.Length = 0 Then Exit Sub
        strValue = hs.EncryptString(strValue, "&Cul8r#1")
      End If

      strMessage = String.Format("Section: {0}, Key: {1}, Value: {2}", strSection, strKey, strValue)
      Call WriteMessage(strMessage, MessageType.Debug)

      '
      ' Save selected settings to global variables
      '
      If strSection = "Options" And strKey = "SnapshotEventMax" Then
        If IsNumeric(strValue) Then
          gSnapshotEventMax = CInt(Val(strValue))
        End If
      End If

      If strSection = "Options" And strKey = "SnapshotsMaxWidth" Then
        gSnapshotsMaxWidth = strValue
      End If

      If strSection = "Options" And strKey = "SnapshotRefreshInterval" Then
        If IsNumeric(strValue) Then
          gSnapshotRefreshInterval = CInt(Val(strValue))
        End If
      End If

      If strSection = "Archive" And strKey = "ArchiveEnabled" Then
        gEventArchiveToDir = CBool(strValue)
      End If

      If strSection = "FTPArchive" And strKey = "ArchiveEnabled" Then
        gEventArchiveToFTP = CBool(strValue)
      End If

      If strSection = "EmailNotification" And strKey = "EmailEnabled" Then
        gEventEmailNotification = CBool(strValue)
      End If

      '
      ' Save the settings
      '
      hs.SaveINISetting(strSection, strKey, strValue, gINIFile)
    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "SaveSetting()")
    End Try

  End Sub

#End Region

#End Region

#Region "NetCam Threads"

  ''' <summary>
  ''' Refresh Snapshot Thread
  ''' </summary>
  ''' <remarks></remarks>
  Friend Sub RefreshSnapshotsThread()

    Dim strMessage As String = ""

    Dim bAbortThread As Boolean = False

    Try
      '
      ' Begin Main While Loop
      '
      While bAbortThread = False

        If gSnapshotRefreshInterval > 0 Then
          RefreshLatestSnapshots()
        End If

        '
        ' Sleep the requested number of minutes between runs
        '
        If gSnapshotRefreshInterval > 0 Then
          Thread.Sleep(1000 * gSnapshotRefreshInterval)
        Else
          Thread.Sleep(5000)
        End If

      End While ' Stay in thread until we get an abort/exit request

    Catch pEx As ThreadAbortException
      ' 
      ' There was a normal request to terminate the thread.  
      '
      bAbortThread = True      ' Not actually needed
      WriteMessage(String.Format("RefreshSnapshotsThread received abort request, terminating normally."), MessageType.Debug)
    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "RefreshSnapshotsThread()")
    End Try

  End Sub

  ''' <summary>
  ''' Refreshes the NetCam List
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub RefreshNetCamList()

    Dim NetCamList As New SortedList

    Try
      WriteMessage("Refreshing list of network cameras ...", MessageType.Debug)

      Try

        SyncLock NetCams.SyncRoot

          NetCams.Clear()

          NetCamList = GetNetCamList()
          For Each strNetCamId As String In NetCamList.Keys

            Dim NetCamDevice As Hashtable = NetCamList(strNetCamId)
            AddNetCamDevice(NetCamDevice)
            hspi_devices.CreateNetCamDevice(NetCamDevice)

          Next

        End SyncLock

      Catch pEx As Exception
        '
        ' Return message
        '
        ProcessError(pEx, "RefreshNetCamList()")
      End Try

    Catch pEx As Exception
      '
      ' Return message
      '
      ProcessError(pEx, "RefreshNetCamList()")
    End Try

  End Sub

  ''' <summary>
  ''' Get the camera picture filenames
  ''' </summary>
  ''' <param name="strNetCamName"></param>
  ''' <param name="strEventId"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetCameraEventSnapshots(ByVal strNetCamName As String, ByVal strEventId As String) As ArrayList

    Dim FileNames As New ArrayList

    Try
      Dim SnapshotPath As String = FixPath(String.Format("{0}\{1}\{2}", gSnapshotDirectory, strNetCamName, strEventId))
      Dim DirectoryInfo As New System.IO.DirectoryInfo(SnapshotPath)
      Dim Files() As System.IO.FileInfo = DirectoryInfo.GetFiles("*_thumbnail.jpg", SearchOption.TopDirectoryOnly)

      Dim comparer As IComparer = New DateComparer()
      Array.Sort(Files, comparer)

      For Each MyFileInfo As FileInfo In Files
        Dim CreationDate As DateTime = MyFileInfo.CreationTime

        Dim ThumbnailFileName As String = MyFileInfo.FullName
        Dim SnapshotFileName As String = ThumbnailFileName.Replace("thumbnail", "snapshot")

        FileNames.Add(New SnapshotInfo(strNetCamName, strEventId, ThumbnailFileName, SnapshotFileName, CreationDate))
      Next

    Catch pEx As Exception
      '
      ' Process Program Exception
      '
      Call ProcessError(pEx, "GetCameraEventSnapshots()")
    End Try

    Return FileNames

  End Function

  ''' <summary>
  ''' Gets new NetCam Event Id
  ''' </summary>
  ''' <param name="NetCamEvent"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function StartNetCamEvent(ByRef NetCamEvent As NetCamEvent) As Integer

    Dim strMessage As String = ""
    Dim iInsertId As Integer = 0

    Try

      Select Case DBConnectionMain.State
        Case ConnectionState.Broken, ConnectionState.Closed
          strMessage = "Unable to complete database transaction because the database " _
                     & "connection has not been initialized."
          Throw New System.Exception(strMessage)
      End Select

      Dim start_ts As Long = ConvertDateTimeToEpoch(DateTime.Now)

      Dim strSQL As String = String.Format("INSERT INTO tblNetCamEvents (" _
                                           & " netcam_id, start_ts, frames_requested, email_rcpt, email_attach, emailed, compressed, uploaded, archived" _
                                           & ") VALUES (" _
                                           & "{0}, {1}, {2}, '{3}', {4}, {5}, {6}, {7}, {8});", _
                                           NetCamEvent.NetCamId, _
                                           start_ts, _
                                           NetCamEvent.FramesRequested, _
                                           NetCamEvent.EmailRecipient, _
                                           NetCamEvent.EmailAttachment, _
                                           NetCamEvent.Emailed, _
                                           NetCamEvent.Compressed, _
                                           NetCamEvent.Uploaded, _
                                           NetCamEvent.Compressed, _
                                           NetCamEvent.Archived)

      strSQL &= "SELECT last_insert_rowid();"

      Using dbcmd As DbCommand = DBConnectionMain.CreateCommand()

        dbcmd.Connection = DBConnectionMain
        dbcmd.CommandType = CommandType.Text
        dbcmd.CommandText = strSQL

        SyncLock SyncLockMain
          iInsertId = dbcmd.ExecuteScalar()
        End SyncLock

        dbcmd.Dispose()

      End Using

    Catch pEx As Exception
      Call ProcessError(pEx, "StartNetCamEvent()")
    End Try

    Return iInsertId

  End Function

  ''' <summary>
  ''' Updates existing NetCam Profile stored in the database
  ''' </summary>
  ''' <param name="EventId"></param>
  ''' <param name="EndTimestamp"></param>
  ''' <param name="FramesCompleted"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function EndNetCamEvent(ByVal EventId As Integer, _
                                 ByVal EndTimestamp As Long, _
                                 ByVal FramesCompleted As Integer) As Boolean

    Dim strMessage As String = ""

    Try

      Select Case DBConnectionMain.State
        Case ConnectionState.Broken, ConnectionState.Closed
          strMessage = "Unable to complete database transaction because the database " _
                     & "connection has not been initialized."
          Throw New System.Exception(strMessage)
      End Select

      Dim strSQL As String = String.Format("UPDATE tblNetCamEvents SET " & _
                                             "end_ts={0}, " & _
                                             "frames_completed={1} " & _
                                           "WHERE event_id={2};", EndTimestamp, FramesCompleted, EventId)

      Dim iRecordsAffected As Integer = 0

      '
      ' Build the insert/update/delete query
      '
      Using MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

        MyDbCommand.Connection = DBConnectionMain
        MyDbCommand.CommandType = CommandType.Text
        MyDbCommand.CommandText = strSQL

        SyncLock SyncLockMain
          iRecordsAffected = MyDbCommand.ExecuteNonQuery()
        End SyncLock

        MyDbCommand.Dispose()

      End Using

      strMessage = "UpdateNetCamEventId() updated " & iRecordsAffected & " row(s)."
      Call WriteMessage(strMessage, MessageType.Debug)

      If iRecordsAffected > 0 Then
        Return True
      Else
        Return False
      End If

    Catch pEx As Exception
      Call ProcessError(pEx, "UpdateNetCamEventId()")
      Return False
    End Try

  End Function

  ''' <summary>
  ''' Gets the NetCam Event details from the database
  ''' </summary>
  ''' <param name="event_id"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetNetCamEvent(ByVal event_id As Integer) As NetCamEvent

    Dim NetCamEvent As New NetCamEvent

    Try

      Dim strSQL As String = String.Format("SELECT event_id, email_rcpt, email_attach, emailed, compressed, uploaded, archived " & _
                                           "FROM tblNetCamEvents WHERE event_id={0}", event_id)

      '
      ' Initialize the command object
      '
      Using MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

        MyDbCommand.Connection = DBConnectionMain
        MyDbCommand.CommandType = CommandType.Text
        MyDbCommand.CommandText = strSQL

        SyncLock SyncLockMain
          Dim dtrResults As IDataReader = MyDbCommand.ExecuteReader()

          '
          ' Process the resutls
          '
          If dtrResults.Read() Then

            NetCamEvent.NetCamId = dtrResults("event_id")
            NetCamEvent.EmailRecipient = dtrResults("email_rcpt")
            NetCamEvent.EmailAttachment = dtrResults("email_attach")
            NetCamEvent.Emailed = dtrResults("emailed")
            NetCamEvent.Compressed = dtrResults("compressed")
            NetCamEvent.Uploaded = dtrResults("uploaded")
            NetCamEvent.Archived = dtrResults("archived")

          End If

          dtrResults.Close()
        End SyncLock

      End Using

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "GetNetCamEvent()")
    End Try

    Return NetCamEvent

  End Function

  ''' <summary>
  ''' Discovers Foscam Cameras installed on the network
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub DiscoveryBeacon()

    Dim subscriber As UdpClient = New UdpClient()
    Dim bAbortThread As Boolean = False

    Try
      WriteMessage("The IP Camera Device Search Protocol receive routine has started ...", MessageType.Debug)

      subscriber = New UdpClient()

      Dim localEP As IPEndPoint = New IPEndPoint(IPAddress.Any, 10000)

      subscriber.Client.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.ReuseAddress, 1)
      subscriber.Client.Bind(localEP)

      While bAbortThread = False

        Try

          Dim pdata As Byte() = subscriber.Receive(localEP)
          Dim strIPAddress As String = localEP.Address.ToString
          Dim resp As String = Encoding.ASCII.GetString(pdata)

          If pdata.Length >= 88 Then

            Dim strTCPHeader As String = Encoding.ASCII.GetString(pdata, 0, 4)                                                      ' 0-3
            Dim operation As Int16 = CInt(pdata(5)) << 8 Or CInt(pdata(4))                                                          ' 4-5
            '                                                                                                                       ' 6
            '                                                                                                                       ' 7-14
            Dim operationCode As String = Encoding.ASCII.GetString(pdata, 15, 1).TrimEnd(ControlChars.NullChar)                     ' 15
            Dim netcam_mac As String = Encoding.ASCII.GetString(pdata, 23, 13).TrimEnd(ControlChars.NullChar)                       ' 22-35

            If Foscams.ContainsKey(netcam_mac) = False Then
              Dim netcam_name As String = Encoding.ASCII.GetString(pdata, 36, 21).TrimEnd(ControlChars.NullChar)                      ' 36-57
              Dim lngAddress As ULong = CULng(pdata(60)) << 24 Or CULng(pdata(59)) << 16 Or CULng(pdata(58)) << 8 Or CULng(pdata(57)) ' 57-60

              Dim netcam_address As String = LongToIPAddress(lngAddress)
              Dim net_camid As Integer = GetNetCamDeviceId(netcam_address)

              '
              ' Add Foscam Camera to database
              '
              If net_camid = 0 Then
                Dim netcam_type As Integer = IIf(operationCode = "A", 1, 2)
                Dim lngMask As ULong = CULng(pdata(64)) << 24 Or CULng(pdata(63)) << 16 Or CULng(pdata(62)) << 8 Or CULng(pdata(61))    ' 61-64
                Dim lngGateway As ULong = CULng(pdata(68)) << 24 Or CULng(pdata(67)) << 16 Or CULng(pdata(66)) << 8 Or CULng(pdata(65)) ' 65-68
                Dim lngDNS As ULong = CULng(pdata(72)) << 24 Or CULng(pdata(71)) << 16 Or CULng(pdata(70)) << 8 Or CULng(pdata(69))     ' 69-72
                'Dim reserve As ULong = CULng(pdata(72)) << 24 Or CULng(pdata(71)) << 16 Or CULng(pdata(70)) << 8 Or CULng(pdata(69))   ' 73-76
                Dim strSoftwareVersion As String = String.Format("{0}.{1}.{2}.{3}",
                                                                 pdata(77).ToString,
                                                                 pdata(78).ToString,
                                                                 pdata(79).ToString,
                                                                 pdata(80).ToString)                                                    ' 77-80

                Dim strAppVersion As String = String.Format("{0}.{1}.{2}.{3}",
                                                             pdata(81).ToString,
                                                             pdata(82).ToString,
                                                             pdata(83).ToString,
                                                             pdata(84).ToString)                                                         ' 81-84

                Dim netcam_port As Long = CInt(pdata(85)) << 8 Or CInt(pdata(86))                                                       ' 85-86
                Dim dhcpEnabled As Int16 = CInt(pdata(87))

                WriteMessage(String.Format("Discovered Foscam Camera {0} [{1}] at {2}.  Software Version is {3}, Application Version: {4}",
                                           netcam_mac,
                                           netcam_name,
                                           strIPAddress,
                                           strSoftwareVersion,
                                           strAppVersion), MessageType.Informational)

                InsertNetCamDevice(netcam_name, netcam_address, netcam_port, netcam_type, "admin", "")
              End If

              Foscams.Add(netcam_mac, netcam_address)
            End If

          End If

        Catch pEx As Exception
          '
          ' Return message
          '
          WriteMessage("An error occured while processing the IP Camera Device Search Protocol beacon message.", MessageType.Error)
          ProcessError(pEx, "DiscoveryBeacon()")
        End Try

        '
        ' Give up some time
        '
        Thread.Sleep(50)

      End While ' Stay in thread until we get an abort/exit request

    Catch pEx As ThreadAbortException
      ' 
      ' There was a normal request to terminate the thread.  
      '
      bAbortThread = True      ' Not actually needed
      WriteMessage(String.Format("DiscoveryBeacon thread received abort request, terminating normally."), MessageType.Informational)

    Catch pEx As Exception
      '
      ' Return message
      '
      ProcessError(pEx, "DiscoveryBeacon()")

    Finally
      '
      ' Notify that we are exiting the thread 
      '
      WriteMessage(String.Format("DiscoveryBeacon terminated."), MessageType.Debug)

      subscriber.Close()
    End Try

  End Sub

  ''' <summary>
  ''' Discovers Foscam Network Cameras on the network
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub SendDiscoveryBeacon()

    Dim bAbortThread As Boolean = False
    Dim udpSocket As New System.Net.Sockets.Socket(Net.Sockets.AddressFamily.InterNetwork, _
                                                   Net.Sockets.SocketType.Dgram, _
                                                   Net.Sockets.ProtocolType.Udp)

    Try

      WriteMessage("The IP Camera Device Search Protocol send routine has started ...", MessageType.Debug)

      ' 4d:4f:5f:49:00:00:00:00:00:00:00:00:00:00:00:04:00:00:00:04:00:00:00:00:00:00:01
      Dim sendbuf As Byte() = {77, 79, 95, 73, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 4, 0, 0, 0, 4, 0, 0, 0, 0, 0, 0, 1}

      Dim remoteEP As New System.Net.IPEndPoint(System.Net.IPAddress.Broadcast, 10000)
      Dim localEP As IPEndPoint = New IPEndPoint(IPAddress.Any, 10000)

      udpSocket.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.ReuseAddress, 1)
      udpSocket.SetSocketOption(System.Net.Sockets.SocketOptionLevel.Socket, System.Net.Sockets.SocketOptionName.Broadcast, 1)
      udpSocket.Bind(localEP)

      While bAbortThread = False

        '
        ' Transmit the UDP Broadcast Packet
        '
        Try
          udpSocket.SendTo(sendbuf, remoteEP)
        Catch pEx As Exception
          WriteMessage("An error occured while sending the IP Camera Device Search Protocol beacon message.", MessageType.Debug)
          ProcessError(pEx, "DiscoveryBeacon()")
        End Try

        '
        ' Give up some time
        '
        Thread.Sleep(1000 * 30)

      End While ' Stay in thread until we get an abort/exit request

    Catch pEx As ThreadAbortException
      ' 
      ' There was a normal request to terminate the thread.  
      '
      bAbortThread = True      ' Not actually needed
      WriteMessage(String.Format("SendDiscoveryBeacon thread received abort request, terminating normally."), MessageType.Informational)

      udpSocket.Close()

    Catch pEx As Exception
      '
      ' Return message
      '
      ProcessError(pEx, "SendDiscoveryBeacon()")

    Finally
      '
      ' Notify that we are exiting the thread 
      '
      WriteMessage(String.Format("SendDiscoveryBeacon terminated."), MessageType.Debug)

      udpSocket.Dispose()

    End Try

  End Sub

  ''' <summary>
  ''' Converts Long to an IP address
  ''' </summary>
  ''' <param name="ip"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function LongToIPAddress(ByVal ip As Long) As String

    Return New System.Net.IPAddress(ip).ToString

  End Function

  ''' <summary>
  ''' Performs Snapshot File Maintenance
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub SnapshotFileMaintenance()

    Dim bAbortThread As Boolean = False
    Static Loopcount As Integer = 12

    Try
      WriteMessage("The SnapshotFileMaintenance routine has started ...", MessageType.Debug)

      While bAbortThread = False

        Try
          '
          ' Increment loopcount
          '
          Loopcount += 1

          '
          ' Begin the purge and archive routines (need to not do a for each)
          '
          Dim arrKeys As New ArrayList(NetCams.Keys)

          For Each strNetCamId As String In arrKeys
            PurgeSnapshotEvents(strNetCamId, gSnapshotEventMax)
            ProcessSnapshotEvents(strNetCamId)
          Next

          '
          ' Purge the FTP Archives
          '
          If Loopcount >= 12 Then
            Loopcount = 0
            PurgeFTPArchives()
          End If

        Catch pEx As Exception
          '
          ' Return message
          '
          ProcessError(pEx, "SnapshotFileMaintenance()")
        End Try

        '
        ' Give up some time
        '
        Thread.Sleep(1000 * 60)

      End While ' Stay in thread until we get an abort/exit request

    Catch pEx As ThreadAbortException
      ' 
      ' There was a normal request to terminate the thread.  
      '
      bAbortThread = True      ' Not actually needed
      WriteMessage(String.Format("SnapshotFileMaintenance thread received abort request, terminating normally."), MessageType.Informational)

    Catch pEx As Exception
      '
      ' Return message
      '
      ProcessError(pEx, "SnapshotFileMaintenance()")
    Finally
      '
      ' Notify that we are exiting the thread
      '
      WriteMessage(String.Format("SnapshotFileMaintenance terminated."), MessageType.Debug)
    End Try

  End Sub

  ''' <summary>
  ''' Trims the number of snapshots kept per camera
  ''' </summary>
  ''' <param name="strNetCamId"></param>
  ''' <param name="iEventMax"></param>
  ''' <remarks></remarks>
  Public Sub PurgeSnapshotEvents(ByVal strNetCamId As String, Optional ByVal iEventMax As Integer = 25)

    Try

      Dim LastDate As Date = Now
      Dim LastEvent As String = ""
      Dim iEventCount As Long = 0

      Dim SnapshotPath As String = FixPath(String.Format("{0}\{1}", gSnapshotDirectory, strNetCamId))
      Dim RootDirInfo As New IO.DirectoryInfo(SnapshotPath)
      For Each EventDirInfo As DirectoryInfo In RootDirInfo.GetDirectories
        If EventDirInfo.LastWriteTime < LastDate Then
          LastDate = EventDirInfo.LastWriteTime
          LastEvent = EventDirInfo.FullName
        End If
        iEventCount += 1
      Next

      If iEventCount > iEventMax Then
        If LastEvent <> "" Then

          If Directory.Exists(LastEvent) = True Then
            Directory.Delete(LastEvent, True)
          End If

          PurgeSnapshotEvents(strNetCamId, iEventMax)
        End If
      End If

    Catch pEx As Exception
      '
      ' Return message
      '
      ProcessError(pEx, "PurgeSnapshotEvents()")
    End Try

  End Sub

  ''' <summary>
  ''' Compresses and archives Snapshot Events
  ''' </summary>
  ''' <param name="strNetCamId"></param>
  ''' <remarks></remarks>
  Private Sub ProcessSnapshotEvents(ByVal strNetCamId As String)

    Try
      '
      ' Process Snapshot Events
      '
      Dim SnapshotEventPath As String = FixPath(String.Format("{0}\{1}", gSnapshotDirectory, strNetCamId))
      Dim DirectoryInfo As New System.IO.DirectoryInfo(SnapshotEventPath)
      Dim Files() As System.IO.FileInfo = DirectoryInfo.GetFiles("Event.ini", SearchOption.AllDirectories)

      Dim comparer As IComparer = New DateComparer()
      Array.Sort(Files, comparer)

      For Each MyFileInfo As FileInfo In Files

        Dim strSourceDir As String = MyFileInfo.DirectoryName

        Dim iEventSnapshotCount As Integer = Int16.Parse(GetSetting("ConvertToVideo", "EventSnapshotCount", gEventSnapshotCount))

        Dim strEventName As String = Regex.Match(MyFileInfo.FullName, "(Event\d+)").ToString
        Dim event_id As Integer = Val(Regex.Match(strEventName, "\d+").Value)
        Dim NetCamEvent As NetCamEvent = GetNetCamEvent(event_id)

        Dim strCompressedFilename As String = FixPath(String.Format("{0}\{1}-{2}.zip", MyFileInfo.DirectoryName, strEventName, strNetCamId))
        Dim strVideoFilename As String = FixPath(String.Format("{0}\{1}-{2}.mp4", MyFileInfo.DirectoryName, strEventName, strNetCamId))

        Dim bCompressed As Boolean = False
        Dim bVideoCreated As Boolean = False

        '
        ' Compress the NetCam Event
        '
        If NetCamEvent.Compressed = 1 Then
          '
          ' Compress the NetCam Event Snapshots
          '
          bCompressed = CompressFolder(strSourceDir, strCompressedFilename)
        End If

        '
        ' Create the Video file from static images
        '
        If Directory.GetFiles(MyFileInfo.DirectoryName, "*_snapshot.jpg", SearchOption.TopDirectoryOnly).Length() > iEventSnapshotCount Then
          bVideoCreated = CreateVideoFile(strSourceDir, strNetCamId, strEventName)
        End If

        '
        ' Upload the compressed file to the FTP directory
        '
        If gEventArchiveToFTP = True And NetCamEvent.Uploaded = 1 Then
          '
          ' Determine what to send to the FTP archive
          '
          Dim iArchiveFiles As Integer = Int16.Parse(GetSetting("Archive", "ArchiveFiles", "1"))

          If (iArchiveFiles And EventArchveFile.Compressed) And bCompressed = True Then
            ArchiveToFTP(strCompressedFilename)
          End If

          If (iArchiveFiles And EventArchveFile.Video) And bVideoCreated = True Then
            ArchiveToFTP(strVideoFilename)
          End If

        End If

        '
        ' Archive NetCam Event to archive directory
        '
        If gEventArchiveToDir = True And NetCamEvent.Archived = 1 Then
          '
          ' Determine what to send to the archive
          '
          Dim iArchiveFiles As Integer = Int16.Parse(GetSetting("Archive", "ArchiveFiles", "1"))

          If (iArchiveFiles And EventArchveFile.Compressed) And bCompressed = True Then
            ArchiveToDir(strCompressedFilename)
          End If

          If (iArchiveFiles And EventArchveFile.Video) And bVideoCreated = True Then
            ArchiveToDir(strVideoFilename)
          End If

        End If

        '
        ' Send the e-mail notification
        '
        If gEventEmailNotification = True And NetCamEvent.Emailed = 1 Then
          Dim Attachments() As String = {""}
          Select Case NetCamEvent.EmailAttachment
            Case "0"
              '
              ' No attachment
              '

            Case "1"
              '
              ' Last snapshot
              '
              Dim strLastSnapshot As String = FixPath(String.Format("{0}\{1}", MyFileInfo.Directory.Parent.FullName, "last_snapshot.jpg"))
              Attachments.SetValue(strLastSnapshot, 0)

            Case "2"
              '
              ' Compressed
              '
              If bCompressed = True Then
                Attachments.SetValue(strCompressedFilename, 0)
              End If

            Case "4"
              '
              ' Video
              '
              If bVideoCreated = True Then
                Attachments.SetValue(strVideoFilename, 0)
              End If

          End Select

          SendEmailNotification(NetCamEvent.EmailRecipient, strEventName, strNetCamId, Attachments)
        End If

        MyFileInfo.Delete()

      Next

    Catch pEx As Exception
      '
      ' Return message
      '
      ProcessError(pEx, "ProcessSnapshotEvents()")
    End Try

  End Sub

  ''' <summary>
  ''' Creates video from snapshots
  ''' </summary>
  ''' <param name="strSourceDir"></param>
  ''' <param name="strNetCamId"></param>
  ''' <param name="strEventName"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function CreateVideoFile(ByVal strSourceDir As String, _
                                   ByVal strNetCamId As String, _
                                   ByVal strEventName As String) As Boolean

    Try

      Call WriteMessage("Entered CreateVideoFile() subroutine.", MessageType.Debug)

      Dim strProgram As String = String.Format("{0}\{1}", HSAppPath, "ffmpeg.exe")

      If File.Exists(strProgram) = True Then
        '
        ' "C:\Program Files (x86)\HomeSeer HS2\ffmpeg.exe" -r 1/1 -i "NetCam002_%04d_snapshot.jpg" -vcodec libx264 out.mp4
        '
        Dim strInputFiles As String = String.Format("{0}_%04d_snapshot.jpg", strNetCamId)
        Dim strCmdArguments As String = String.Format("-r 1/1 -i ""{0}"" -n -vcodec mpeg4 {1}-{2}.mp4", strInputFiles, strEventName, strNetCamId)


        WriteMessage(String.Format("Running {0} {1}", strProgram, strCmdArguments), MessageType.Debug)

        Dim CmdShell As New CmdShell
        CmdShell.DoCommand(strProgram, strCmdArguments, strSourceDir)

      End If

    Catch pEx As Exception
      '
      ' Return message
      '
      ProcessError(pEx, "CreateVideoFile()")
    End Try

    Return True

  End Function

  ''' <summary>
  ''' Compresses Snapshot Event Folder
  ''' </summary>
  ''' <param name="strSourceDir"></param>
  ''' <param name="strOutputFile"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function CompressFolder(ByVal strSourceDir As String, _
                                  ByVal strOutputFile As String) As Boolean

    Try

      Call WriteMessage("Entered CompressFolder() subroutine.", MessageType.Debug)

      If File.Exists(strOutputFile) = False Then

        Dim zipResult As String = hs.Zip(strSourceDir, strOutputFile)

        If zipResult = "" Then
          WriteMessage(String.Format("Compressing {0} complete.", strSourceDir), MessageType.Debug)
          Return True
        Else
          WriteMessage(String.Format("Compressing {0} failed due to error: ", strSourceDir, zipResult), MessageType.Error)
        End If

      End If

    Catch pEx As Exception
      '
      ' Process Error Message
      '
      ProcessError(pEx, "CompressFolder()")
    End Try

    Return False

  End Function

  ''' <summary>
  ''' Purgest FTP Archive files
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub PurgeFTPArchives()

    Try

      Call WriteMessage("Entered PurgeFTPArchives() subroutine.", MessageType.Debug)

      Dim bFTPArchive As Boolean = CBool(GetSetting("FTPArchive", "ArchiveEnabled", "False"))
      If bFTPArchive = False Then Exit Sub

      Dim FTPFiles As New SortedList()

      Dim strArchiveDir As String = GetSetting("FTPArchive", "ArchiveDir", "/ultranetcam")
      Dim strArchiveLimit As String = GetSetting("FTPArchive", "ArchiveLimit", "50")
      Dim iArhiveLimit As Integer = IIf(IsNumeric(strArchiveLimit) = True, CInt(strArchiveLimit), 50)

      '
      ' Check if user has selected unlimited
      '
      If iArhiveLimit = 0 Then Exit Sub
      iArhiveLimit *= 1048576

      Dim strHostName As String = GetSetting("FTPClient", "ServerName", "")
      Dim strUserName As String = GetSetting("FTPClient", "UserName", "")
      Dim strUserPass As String = GetSetting("FTPClient", "UserPass", "")

      If strHostName.Length = 0 Then
        Throw New Exception("Unable to check archive because the FTP server hostname is emtpy.")
      ElseIf strUserName.Length = 0 Then
        Throw New Exception("Unable to check archive because the FTP user name is emtpy.")
      ElseIf strUserPass.Length = 0 Then
        Throw New Exception("Unable to check archive because the FTP user name is emtpy.")
      End If

      Dim ftp As New FTPclient(strHostName, strUserName, strUserPass)

      '
      ' Get a listing of available files
      '
      Dim iArchiveSize As Long = 0
      For Each file As FTPfileInfo In ftp.ListDirectoryDetail(strArchiveDir).GetFiles("zip")

        ' Event1-NetCam001
        If Regex.IsMatch(file.FullName, "Event\d+-NetCam\d\d\d\.zip") = True Then
          iArchiveSize += file.Size
          If FTPFiles.ContainsKey(file.FullName) = False Then
            WriteMessage(String.Format("{0} is {1} bytes.", file.Filename, file.Size), MessageType.Debug)
            FTPFiles.Add(file.Filename, file.Size)
          End If
        End If

      Next file

      For Each file As FTPfileInfo In ftp.ListDirectoryDetail(strArchiveDir).GetFiles("mp4")

        ' NetCam001-Event1
        If Regex.IsMatch(file.FullName, "Event\d+-NetCam\d\d\d\.mp4") = True Then
          iArchiveSize += file.Size
          If FTPFiles.ContainsKey(file.FullName) = False Then
            WriteMessage(String.Format("{0} is {1} bytes.", file.Filename, file.Size), MessageType.Debug)
            FTPFiles.Add(file.Filename, file.Size)
          End If
        End If

      Next file

      If iArchiveSize < iArhiveLimit Then
        WriteMessage(String.Format("FTP Archive size is {0} bytes which is less than the limit of {1} bytes.", iArchiveSize.ToString, iArhiveLimit.ToString), MessageType.Debug)
        Exit Sub
      Else
        '
        ' Begin purge routine
        '
        WriteMessage(String.Format("FTP Archive size is {0} bytes which is more than the limit of {1} bytes.", iArchiveSize.ToString, iArhiveLimit.ToString), MessageType.Debug)

        For Each strFileName As String In FTPFiles.Keys

          iArchiveSize -= FTPFiles(strFileName)

          Dim bSuccess As Boolean = ftp.FtpDelete(String.Format("{0}/{1}", strArchiveDir, strFileName))
          If bSuccess = True Then
            WriteMessage(String.Format("PurgeFTPArchives() purged file {0}.", strFileName), MessageType.Debug)
          Else
            WriteMessage(String.Format("PurgeFTPArchives() failed to delete {0} due to error.", strFileName), MessageType.Warning)
          End If

          If iArchiveSize < iArhiveLimit Then Exit For
          Thread.Sleep(1000)

        Next

      End If

    Catch pEx As Exception
      '
      ' Process Error Message
      '
      ProcessError(pEx, "PurgeFTPArchives()")
    End Try

  End Sub

  ''' <summary>
  ''' Archives File to FTP Archive
  ''' </summary>
  ''' <param name="strLocalFileName"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function ArchiveToFTP(ByVal strLocalFileName As String) As Boolean

    Try

      Call WriteMessage("Entered ArchiveToFTP() function.", MessageType.Debug)

      Dim strArchiveDir As String = GetSetting("FTPArchive", "ArchiveDir", "/ultranetcam")
      Dim strArchiveFilename As String = Regex.Match(strLocalFileName, "([^\\/]+)$").ToString
      Dim strRemoteFileName As String = String.Format("{0}/{1}", strArchiveDir, strArchiveFilename)

      Dim strHostName As String = GetSetting("FTPClient", "ServerName", "")
      Dim strUserName As String = GetSetting("FTPClient", "UserName", "")
      Dim strUserPass As String = GetSetting("FTPClient", "UserPass", "")

      If strHostName.Length = 0 Then
        Throw New Exception(String.Format("Unable to archive {0} because the FTP server hostname is emtpy.", strArchiveFilename))
      ElseIf strUserName.Length = 0 Then
        Throw New Exception(String.Format("Unable to archive {0} because the FTP user name is emtpy.", strArchiveFilename))
      ElseIf strUserPass.Length = 0 Then
        Throw New Exception(String.Format("Unable to archive {0} because the FTP user password is emtpy.", strArchiveFilename))
      ElseIf strLocalFileName.Length = 0 Then
        Throw New Exception(String.Format("Unable to archive {0} because the FTP local filename is emtpy.", strArchiveFilename))
      ElseIf strRemoteFileName.Length = 0 Then
        Throw New Exception(String.Format("Unable to archive {0} because the FTP remote filename is emtpy.", strArchiveFilename))
      End If

      Dim ftp As New FTPclient(strHostName, strUserName, strUserPass)

      ftp.FtpCreateDirectory(strArchiveDir)

      Return ftp.Upload(strLocalFileName, strRemoteFileName)

    Catch pEx As Exception
      '
      ' Process Error Message
      '
      ProcessError(pEx, "ArchiveToFTP()")
      Return False
    End Try

  End Function

  ''' <summary>
  ''' Archives File to FTP Archive
  ''' </summary>
  ''' <param name="strLocalFileName"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function ArchiveToDir(ByVal strLocalFileName As String) As Boolean

    Try

      Call WriteMessage("Entered ArchiveToDir() function.", MessageType.Debug)

      Dim strArchiveDir As String = GetSetting("Archive", "ArchiveDir", "")
      Dim strArchiveFilename As String = Regex.Match(strLocalFileName, "([^\\/]+)$").ToString
      Dim strRemoteFilename As String = String.Format("{0}/{1}", strArchiveDir, strArchiveFilename)

      If strLocalFileName.Length = 0 Then
        Throw New Exception(String.Format("Unable to archive {0} because the local filename is emtpy.", strArchiveFilename))
      ElseIf strRemoteFilename.Length = 0 Then
        Throw New Exception(String.Format("Unable to archive {0} because the archive filename is emtpy.", strArchiveFilename))
      End If

      If Directory.Exists(strArchiveDir) = True Then
        File.Copy(strLocalFileName, strRemoteFilename, False)
      End If

      Return True

    Catch pEx As Exception
      '
      ' Process Error Message
      '
      ProcessError(pEx, "ArchiveToDir()")
      Return False
    End Try

  End Function

  ''' <summary>
  ''' Sends Email Notification
  ''' </summary>
  ''' <param name="strAlternateEmailRcpt"></param>
  ''' <param name="strEventName"></param>
  ''' <param name="strNetCamId"></param>
  ''' <param name="Attachments"></param>
  ''' <remarks></remarks>
  Private Sub SendEmailNotification(ByVal strAlternateEmailRcpt As String, _
                                    ByVal strEventName As String, _
                                    ByVal strNetCamId As String, _
                                    ByVal Attachments() As String)

    Try

      Call WriteMessage("Entered SendEmailNotification() function.", MessageType.Debug)

      '
      ' Get E-mail Settings
      '
      Dim strEmailFromDefault As String = hs.GetINISetting("Settings", "gSMTPFrom", "")
      Dim strEmailRcptTo As String = hs.GetINISetting("Settings", "gSMTPTo", "")

      Dim EmailRcptTo As String = GetSetting("EmailNotification", "EmailRcptTo", strEmailRcptTo)
      Dim EmailFrom As String = GetSetting("EmailNotification", "EmailFrom", strEmailFromDefault)
      Dim EmailSubject As String = GetSetting("EmailNotification", "EmailSubject", EMAIL_SUBJECT)
      Dim EmailBodyTemplate As String = GetSetting("EmailNotification", "EmailBody", EMAIL_BODY_TEMPLATE)

      '
      ' Change the recipient e-mail address
      '
      If Regex.IsMatch(strAlternateEmailRcpt, ".+@.+") = True Then
        EmailRcptTo = strAlternateEmailRcpt
      End If

      If EmailRcptTo.Length = 0 Then
        Throw New Exception(String.Format("Unable to sent e-mail notification for {0} because the e-mail recipient is emtpy.", strEventName))
      ElseIf EmailFrom.Length = 0 Then
        Throw New Exception(String.Format("Unable to sent e-mail notification for {0} because the e-mail sender is emtpy.", strEventName))
      ElseIf EmailSubject.Length = 0 Then
        Throw New Exception(String.Format("Unable to sent e-mail notification for {0} because the e-mail subject is emtpy.", strEventName))
      ElseIf EmailBodyTemplate.Length = 0 Then
        Throw New Exception(String.Format("Unable to sent e-mail notification for {0} because the e-mail body is emtpy.", strEventName))
      End If

      Dim EmailBody As String = EmailBodyTemplate

      '
      ' Get the NetCam Name for the e-mail notification 
      '
      Dim strNetCamName As String = strNetCamId
      SyncLock NetCams.SyncRoot
        If NetCams.ContainsKey(strNetCamId) = True Then
          Dim NetCamDevice As NetCamDevice = NetCams(strNetCamId)
          strNetCamName = NetCamDevice.Name
        End If
      End SyncLock

      '
      ' Get the URL data
      '
      Dim strLanIPAddress As String = hs.LANIP()
      Dim strWanIPAddress As String = hs.WANIP()
      Dim iWebServerPort As String = hs.WebServerPort()
      Dim iWebServerSSLPort As Integer = hs.WebServerSSLPort()

      strLanIPAddress = Regex.Match(strLanIPAddress, "\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}").Value
      strWanIPAddress = Regex.Match(strWanIPAddress, "\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}").Value

      Dim URLPath As String = "UltraNetCam3"

      EmailBody = EmailBody.Replace("$event_id", strEventName)
      EmailBody = EmailBody.Replace("$camera_id", strNetCamId)
      EmailBody = EmailBody.Replace("$camera_name", strNetCamName)

      EmailBody = EmailBody.Replace("$snapshot_path", URLPath)

      EmailBody = EmailBody.Replace("$lan_ip", strLanIPAddress)
      EmailBody = EmailBody.Replace("$lan_port", iWebServerPort)

      EmailBody = EmailBody.Replace("$wan_ip", strWanIPAddress)
      EmailBody = EmailBody.Replace("$wan_port", iWebServerSSLPort)

      EmailBody = EmailBody.Replace("$lan_url_snapshots", String.Format("http://{0}:{1}/{2}", strLanIPAddress, iWebServerPort, URLPath))
      EmailBody = EmailBody.Replace("$wan_url_snapshots", String.Format("https://{0}:{1}/{2}", strWanIPAddress, iWebServerSSLPort, URLPath))

      EmailBody = EmailBody.Replace("~", vbCrLf)

      EmailSubject = EmailSubject.Replace("$event_id", strEventName)
      EmailSubject = EmailSubject.Replace("$camera_id", strNetCamId)
      EmailSubject = EmailSubject.Replace("$camera_name", strNetCamName)

      Dim List() As String = hs.GetPluginsList()
      If List.Contains("UltraSMTP3:") = True Then
        '
        ' Send e-mail using UltraSMTP3
        '
        hs.PluginFunction("UltraSMTP3", "", "SendMail", New Object() {EmailRcptTo, EmailSubject, EmailBody, Attachments})
      Else
        '
        ' Send e-mail using HomeSeer
        '
        hs.SendEmail(EmailRcptTo, EmailFrom, "", "", EmailSubject, EmailBody, Attachments(0))
      End If

    Catch pEx As Exception
      '
      ' Process Program Exception
      '
      ProcessError(pEx, "SendEmailNotification()")
    End Try

  End Sub

  ''' <summary>
  ''' Get the NetCam devices
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetNetCamList() As SortedList

    Dim NetCamDevices As New SortedList
    Dim strSQL As String = ""

    Try
      '
      ' Define the SQL Query
      '
      strSQL = "SELECT A.netcam_id, A.netcam_name, A.netcam_address, A.netcam_port, A.netcam_type, B.snapshot_path, B.videostream_path, A.auth_user, A.auth_pass FROM tblNetCamDevices as A LEFT JOIN tblNetCamTypes as B on A.netcam_type = B.netcam_type"

      '
      ' Execute the data reader
      '
      Using MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

        MyDbCommand.Connection = DBConnectionMain
        MyDbCommand.CommandType = CommandType.Text
        MyDbCommand.CommandText = strSQL

        SyncLock SyncLockMain
          Dim dtrResults As IDataReader = MyDbCommand.ExecuteReader()

          '
          ' Process the resutls
          '
          While dtrResults.Read()
            Dim NetCamDevice As New Hashtable

            NetCamDevice.Add("netcam_id", dtrResults("netcam_id"))
            NetCamDevice.Add("netcam_name", dtrResults("netcam_name"))
            NetCamDevice.Add("netcam_address", dtrResults("netcam_address"))
            NetCamDevice.Add("netcam_port", dtrResults("netcam_port"))
            NetCamDevice.Add("netcam_type", dtrResults("netcam_type"))
            NetCamDevice.Add("snapshot_path", IIf(dtrResults("snapshot_path").Equals(System.DBNull.Value), "", dtrResults("snapshot_path")))
            NetCamDevice.Add("videostream_path", IIf(dtrResults("videostream_path").Equals(System.DBNull.Value), "", dtrResults("videostream_path")))
            NetCamDevice.Add("auth_user", dtrResults("auth_user"))
            NetCamDevice.Add("auth_pass", hs.DecryptString(dtrResults("auth_pass"), "&Cul8r#1"))

            Dim strNetCamId As String = String.Format("NetCam{0}", dtrResults("netcam_id").ToString.PadLeft(3, "0"))
            If NetCamDevices.ContainsKey(strNetCamId) = False Then
              NetCamDevices.Add(strNetCamId, NetCamDevice)
            End If
          End While

          dtrResults.Close()
        End SyncLock

        MyDbCommand.Dispose()

      End Using

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "GetNetCamList()")
    End Try

    Return NetCamDevices

  End Function

  ''' <summary>
  ''' Returns the netcam profile for the specified IP address
  ''' </summary>
  ''' <param name="netcam_id"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetNetCamDevice(ByVal netcam_id As Integer) As Hashtable

    Dim NetCamDevice As New Hashtable

    Try

      Dim strSQL As String = "SELECT netcam_id, netcam_name, netcam_address, netcam_port, snapshot_path, videostream_path, auth_user, auth_pass FROM tblNetCamDevices"

      '
      ' Initialize the command object
      '
      Using MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

        MyDbCommand.Connection = DBConnectionMain
        MyDbCommand.CommandType = CommandType.Text
        MyDbCommand.CommandText = strSQL

        SyncLock SyncLockMain
          Dim dtrResults As IDataReader = MyDbCommand.ExecuteReader()

          '
          ' Process the resutls
          '
          If dtrResults.Read() Then

            NetCamDevice.Add("netcam_id", dtrResults("netcam_id"))
            NetCamDevice.Add("netcam_name", dtrResults("netcam_name"))
            NetCamDevice.Add("netcam_address", dtrResults("netcam_address"))
            NetCamDevice.Add("netcam_port", dtrResults("netcam_port"))
            NetCamDevice.Add("snapshot_path", IIf(dtrResults("snapshot_path").Equals(System.DBNull.Value), "", dtrResults("snapshot_path")))
            NetCamDevice.Add("videostream_path", IIf(dtrResults("videostream_path").Equals(System.DBNull.Value), "", dtrResults("videostream_path")))
            NetCamDevice.Add("auth_user", dtrResults("auth_user"))
            NetCamDevice.Add("auth_pass", hs.DecryptString(dtrResults("auth_pass"), "&Cul8r#1"))

          End If

          dtrResults.Close()
        End SyncLock

      End Using

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "GetNetCamDevice()")
    End Try

    Return NetCamDevice

  End Function

  ''' <summary>
  ''' Add the NetCam to the NetCams hashtable
  ''' </summary>
  ''' <param name="NewNetCamDevice"></param>
  ''' <remarks></remarks>
  Private Sub AddNetCamDevice(ByRef NewNetCamDevice As Hashtable)

    Try

      If NewNetCamDevice.ContainsKey("netcam_id") = False Then Exit Sub

      SyncLock NetCams.SyncRoot

        Dim strNetCamId As String = String.Format("NetCam{0}", NewNetCamDevice("netcam_id").ToString.PadLeft(3, "0"))

        If NetCams.ContainsKey(strNetCamId) = False Then
          '
          ' Create the new PioneerAVR object
          '
          Dim NetCamDevice As New NetCamDevice(NewNetCamDevice("netcam_id"), _
                                               NewNetCamDevice("netcam_name"), _
                                               NewNetCamDevice("netcam_address"), _
                                               NewNetCamDevice("netcam_port"), _
                                               NewNetCamDevice("snapshot_path"), _
                                               NewNetCamDevice("videostream_path"), _
                                               NewNetCamDevice("auth_user"), _
                                               NewNetCamDevice("auth_pass"))

          '
          ' Add the NetCam object to global hashtable
          '
          WriteMessage(String.Format("Adding NetCam Object with Id {0}", strNetCamId), MessageType.Debug)
          NetCams.Add(strNetCamId, NetCamDevice)
        End If

      End SyncLock

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "AddNetCamDevice()")
    End Try

  End Sub

  <Flags()> Public Enum EventArchveFile
    Compressed = 1
    Video = 2
  End Enum

#End Region

#Region "UltraNetCam3 Actions/Triggers/Conditions"

#Region "Trigger Proerties"

  ''' <summary>
  ''' Defines the valid triggers for this plug-in
  ''' </summary>
  ''' <remarks></remarks>
  Sub SetTriggers()
    Dim o As Object = Nothing
    If triggers.Count = 0 Then
      triggers.Add(o, "NetCam Alarm Trigger")           ' 1
    End If
  End Sub

  ''' <summary>
  ''' Lets HomeSeer know our plug-in has triggers
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property HasTriggers() As Boolean
    Get
      SetTriggers()
      Return IIf(triggers.Count > 0, True, False)
    End Get
  End Property

  ''' <summary>
  ''' Returns the trigger count
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerCount() As Integer
    SetTriggers()
    Return triggers.Count
  End Function

  ''' <summary>
  ''' Returns the subtrigger count
  ''' </summary>
  ''' <param name="TriggerNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property SubTriggerCount(ByVal TriggerNumber As Integer) As Integer
    Get
      Dim trigger As trigger
      If ValidTrig(TriggerNumber) Then
        trigger = triggers(TriggerNumber - 1)
        If Not (trigger Is Nothing) Then
          Return 0
        Else
          Return 0
        End If
      Else
        Return 0
      End If
    End Get
  End Property

  ''' <summary>
  ''' Returns the trigger name
  ''' </summary>
  ''' <param name="TriggerNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property TriggerName(ByVal TriggerNumber As Integer) As String
    Get
      If Not ValidTrig(TriggerNumber) Then
        Return ""
      Else
        Return String.Format("{0}: {1}", IFACE_NAME, triggers.Keys(TriggerNumber - 1))
      End If
    End Get
  End Property

  ''' <summary>
  ''' Returns the subtrigger name
  ''' </summary>
  ''' <param name="TriggerNumber"></param>
  ''' <param name="SubTriggerNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property SubTriggerName(ByVal TriggerNumber As Integer, ByVal SubTriggerNumber As Integer) As String
    Get
      Dim trigger As trigger
      If ValidSubTrig(TriggerNumber, SubTriggerNumber) Then
        Return ""
      Else
        Return ""
      End If
    End Get
  End Property

  ''' <summary>
  ''' Determines if a trigger is valid
  ''' </summary>
  ''' <param name="TrigIn"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Friend Function ValidTrig(ByVal TrigIn As Integer) As Boolean
    SetTriggers()
    If TrigIn > 0 AndAlso TrigIn <= triggers.Count Then
      Return True
    End If
    Return False
  End Function

  ''' <summary>
  ''' Determines if the trigger is a valid subtrigger
  ''' </summary>
  ''' <param name="TrigIn"></param>
  ''' <param name="SubTrigIn"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ValidSubTrig(ByVal TrigIn As Integer, ByVal SubTrigIn As Integer) As Boolean
    Return False
  End Function

  ''' <summary>
  ''' Tell HomeSeer which triggers have conditions
  ''' </summary>
  ''' <param name="TriggerNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property HasConditions(ByVal TriggerNumber As Integer) As Boolean
    Get
      Select Case TriggerNumber
        Case 0
          Return True   ' Render trigger as IF / AND IF
        Case Else
          Return False  ' Render trigger as IF / OR IF
      End Select
    End Get
  End Property

  ''' <summary>
  ''' HomeSeer will set this to TRUE if the trigger is being used as a CONDITION.  
  ''' Check this value in BuildUI and other procedures to change how the trigger is rendered if it is being used as a condition or a trigger.
  ''' </summary>
  ''' <param name="TrigInfo"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Property Condition(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As Boolean
    Set(ByVal value As Boolean)

      Dim UID As String = TrigInfo.UID.ToString

      Dim trigger As New trigger
      If Not (TrigInfo.DataIn Is Nothing) Then
        DeSerializeObject(TrigInfo.DataIn, trigger)
      End If

      ' TriggerCondition(sKey) = value

    End Set
    Get

      Dim UID As String = TrigInfo.UID.ToString

      Dim trigger As New trigger
      If Not (TrigInfo.DataIn Is Nothing) Then
        DeSerializeObject(TrigInfo.DataIn, trigger)
      End If

      Return False

    End Get
  End Property

  ''' <summary>
  ''' Determines if a trigger is a condition
  ''' </summary>
  ''' <param name="sKey"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Property TriggerCondition(sKey As String) As Boolean
    Get

      If conditions.ContainsKey(sKey) = True Then
        Return conditions(sKey)
      Else
        Return False
      End If

    End Get
    Set(value As Boolean)

      If conditions.ContainsKey(sKey) = False Then
        conditions.Add(sKey, value)
      Else
        conditions(sKey) = value
      End If

    End Set
  End Property

  ''' <summary>
  ''' Called when HomeSeer wants to check if a condition is true
  ''' </summary>
  ''' <param name="TrigInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerTrue(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As Boolean

    Dim UID As String = TrigInfo.UID.ToString

    Dim trigger As New trigger
    If Not (TrigInfo.DataIn Is Nothing) Then
      DeSerializeObject(TrigInfo.DataIn, trigger)
    End If

    Return False
  End Function

#End Region

#Region "Trigger Interface"

  ''' <summary>
  ''' Builds the Trigger UI for display on the HomeSeer events page
  ''' </summary>
  ''' <param name="sUnique"></param>
  ''' <param name="TrigInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerBuildUI(ByVal sUnique As String, ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As String

    Dim UID As String = TrigInfo.UID.ToString
    Dim stb As New StringBuilder

    Dim trigger As New trigger
    If Not (TrigInfo.DataIn Is Nothing) Then
      DeSerializeObject(TrigInfo.DataIn, trigger)
    Else 'new event, so clean out the trigger object
      trigger = New trigger
    End If

    Select Case TrigInfo.TANumber
      Case NetCamTriggers.AlarmTrigger
        Dim triggerName As String = GetEnumName(NetCamTriggers.AlarmTrigger)

        Dim ActionSelected As String = trigger.Item("NetCam")

        Dim actionId As String = String.Format("{0}{1}_{2}_{3}", triggerName, "NetCam", UID, sUnique)

        Dim jqNetCam As New clsJQuery.jqDropList(actionId, Pagename, True)
        jqNetCam.autoPostBack = True

        jqNetCam.AddItem("(Select Network Camera)", "", (ActionSelected = ""))
        For Each strNetCamId As String In NetCams.Keys
          Dim NetCamDevice As NetCamDevice = NetCams(strNetCamId)

          Dim strOptionValue As String = strNetCamId
          Dim strOptionName As String = strOptionValue & " [" & NetCamDevice.Name & "]"
          jqNetCam.AddItem(strOptionName, strOptionValue, (ActionSelected = strOptionValue))
        Next

        stb.Append(jqNetCam.Build)

    End Select

    Return stb.ToString
  End Function

  ''' <summary>
  ''' Process changes to the trigger from the HomeSeer events page
  ''' </summary>
  ''' <param name="PostData"></param>
  ''' <param name="TrigInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerProcessPostUI(ByVal PostData As System.Collections.Specialized.NameValueCollection, _
                                       ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As HomeSeerAPI.IPlugInAPI.strMultiReturn

    Dim Ret As New HomeSeerAPI.IPlugInAPI.strMultiReturn

    Dim UID As String = TrigInfo.UID.ToString
    Dim TANumber As Integer = TrigInfo.TANumber

    ' When plug-in calls such as ...BuildUI, ...ProcessPostUI, or ...FormatUI are called and there is
    ' feedback or an error condition that needs to be reported back to the user, this string field 
    ' can contain the message to be displayed to the user in HomeSeer UI.  This field is cleared by
    ' HomeSeer after it is displayed to the user.
    Ret.sResult = ""

    ' We cannot be passed info ByRef from HomeSeer, so turn right around and return this same value so that if we want, 
    '   we can exit here by returning 'Ret', all ready to go.  If in this procedure we need to change DataOut or TrigInfo,
    '   we can still do that.
    Ret.DataOut = TrigInfo.DataIn
    Ret.TrigActInfo = TrigInfo

    If PostData Is Nothing Then Return Ret
    If PostData.Count < 1 Then Return Ret

    ' DeSerializeObject
    Dim trigger As New trigger
    If Not (TrigInfo.DataIn Is Nothing) Then
      DeSerializeObject(TrigInfo.DataIn, trigger)
    End If
    trigger.uid = UID

    Dim parts As Collections.Specialized.NameValueCollection = PostData

    Try

      Select Case TANumber
        Case NetCamTriggers.AlarmTrigger
          Dim triggerName As String = GetEnumName(NetCamTriggers.AlarmTrigger)

          For Each sKey As String In parts.Keys
            If sKey Is Nothing Then Continue For
            If String.IsNullOrEmpty(sKey.Trim) Then Continue For

            Select Case True
              Case InStr(sKey, triggerName & "NetCam_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                trigger.Item("NetCam") = ActionValue

            End Select
          Next

      End Select

      ' The serialization data for the plug-in object cannot be 
      ' passed ByRef which means it can be passed only one-way through the interface to HomeSeer.
      ' If the plug-in receives DataIn, de-serializes it into an object, and then makes a change 
      ' to the object, this is where the object can be serialized again and passed back to HomeSeer
      ' for storage in the HomeSeer database.

      ' SerializeObject
      If Not SerializeObject(trigger, Ret.DataOut) Then
        Ret.sResult = IFACE_NAME & " Error, Serialization failed. Signal Trigger not added."
        Return Ret
      End If

    Catch ex As Exception
      Ret.sResult = "ERROR, Exception in Trigger UI of " & IFACE_NAME & ": " & ex.Message
      Return Ret
    End Try

    ' All OK
    Ret.sResult = ""
    Return Ret

  End Function

  ''' <summary>
  ''' Determines if a trigger is properly configured
  ''' </summary>
  ''' <param name="TrigInfo"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property TriggerConfigured(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As Boolean
    Get
      Dim Configured As Boolean = True
      Dim UID As String = TrigInfo.UID.ToString

      Dim trigger As New trigger
      If Not (TrigInfo.DataIn Is Nothing) Then
        DeSerializeObject(TrigInfo.DataIn, trigger)
      End If

      Select Case TrigInfo.TANumber
        Case NetCamTriggers.AlarmTrigger
          If trigger.Item("NetCam") = "" Then Configured = False

      End Select

      Return Configured
    End Get
  End Property

  ''' <summary>
  ''' Formats the trigger for display
  ''' </summary>
  ''' <param name="TrigInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerFormatUI(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As String

    Dim stb As New StringBuilder

    Dim UID As String = TrigInfo.UID.ToString

    Dim trigger As New trigger
    If Not (TrigInfo.DataIn Is Nothing) Then
      DeSerializeObject(TrigInfo.DataIn, trigger)
    End If

    Select Case TrigInfo.TANumber
      Case NetCamTriggers.AlarmTrigger
        If trigger.uid <= 0 Then
          stb.Append("Trigger has not been properly configured.")
        Else
          Dim strTriggerName As String = "Alarm Trigger"
          Dim strNetCamId As String = trigger.Item("NetCam")

          If NetCams.ContainsKey(strNetCamId) Then
            Dim NetCamDevice As NetCamDevice = NetCams(strNetCamId)
            strNetCamId = String.Format("{0} [{1}]", strNetCamId, NetCamDevice.Name)
          End If

          stb.AppendFormat("{0} {1}", strNetCamId, strTriggerName)
        End If

    End Select

    Return stb.ToString
  End Function

  ''' <summary>
  ''' Checks to see if trigger should fire
  ''' </summary>
  ''' <param name="Plug_Name"></param>
  ''' <param name="TrigID"></param>
  ''' <param name="SubTrig"></param>
  ''' <param name="strTrigger"></param>
  ''' <remarks></remarks>
  Private Sub CheckTrigger(Plug_Name As String, TrigID As Integer, SubTrig As Integer, strTrigger As String)

    Try
      '
      ' Check HomeSeer Triggers
      '
      If Plug_Name.Contains(":") = False Then Plug_Name &= ":"
      Dim TrigsToCheck() As IAllRemoteAPI.strTrigActInfo = callback.TriggerMatches(Plug_Name, TrigID, SubTrig)

      Try

        For Each TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo In TrigsToCheck
          Dim UID As String = TrigInfo.UID.ToString

          If Not (TrigInfo.DataIn Is Nothing) Then

            Dim trigger As New trigger
            DeSerializeObject(TrigInfo.DataIn, trigger)

            Select Case TrigID

              Case NetCamTriggers.AlarmTrigger
                Dim strTriggerName As String = "Alarm Trigger"
                Dim strNetCam As String = trigger.Item("NetCam")

                Dim strTriggerCheck As String = String.Format("{0},{1}", strTriggerName, strNetCam)
                If Regex.IsMatch(strTrigger, strTriggerCheck) = True Then
                  callback.TriggerFire(IFACE_NAME, TrigInfo)
                End If

            End Select

          End If

        Next

      Catch pEx As Exception

      End Try

    Catch pEx As Exception

    End Try

  End Sub

#End Region

#Region "Action Properties"

  ''' <summary>
  ''' Defines the valid actions for this plug-in
  ''' </summary>
  ''' <remarks></remarks>
  Sub SetActions()
    Dim o As Object = Nothing
    If actions.Count = 0 Then
      actions.Add(o, "Event Snapshots")           ' 1
      actions.Add(o, "Take Snapshot")             ' 2
      actions.Add(o, "Purge Snapshots")           ' 2
      actions.Add(o, "CGI Action")                ' 4
    End If
  End Sub

  ''' <summary>
  ''' Returns the action count
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function ActionCount() As Integer
    SetActions()
    Return actions.Count
  End Function

  ''' <summary>
  ''' Returns the action name
  ''' </summary>
  ''' <param name="ActionNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  ReadOnly Property ActionName(ByVal ActionNumber As Integer) As String
    Get
      If Not ValidAction(ActionNumber) Then
        Return ""
      Else
        Return String.Format("{0}: {1}", IFACE_NAME, actions.Keys(ActionNumber - 1))
      End If
    End Get
  End Property

  ''' <summary>
  ''' Determines if an action is valid
  ''' </summary>
  ''' <param name="ActionIn"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Friend Function ValidAction(ByVal ActionIn As Integer) As Boolean
    SetActions()
    If ActionIn > 0 AndAlso ActionIn <= actions.Count Then
      Return True
    End If
    Return False
  End Function

#End Region

#Region "Action Interface"

  ''' <summary>
  ''' Builds the Action UI for display on the HomeSeer events page
  ''' </summary>
  ''' <param name="sUnique"></param>
  ''' <param name="ActInfo"></param>
  ''' <returns></returns>
  ''' <remarks>This function is called from the HomeSeer event page when an event is in edit mode.</remarks>
  Public Function ActionBuildUI(ByVal sUnique As String, ByVal ActInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As String

    Dim UID As String = ActInfo.UID.ToString
    Dim stb As New StringBuilder

    Dim action As New action
    If Not (ActInfo.DataIn Is Nothing) Then
      DeSerializeObject(ActInfo.DataIn, action)
    End If

    Select Case ActInfo.TANumber
      Case NetCamActions.EventSnapshots
        Dim actionName As String = GetEnumName(NetCamActions.EventSnapshots)

        Dim ActionSelected As String = action.Item("NetCam")

        Dim actionId As String = String.Format("{0}{1}_{2}_{3}", actionName, "NetCam", UID, sUnique)

        Dim jqNetCam As New clsJQuery.jqDropList(actionId, Pagename, True)
        jqNetCam.autoPostBack = True

        jqNetCam.AddItem("(Select Network Camera)", "", (ActionSelected = ""))
        For Each strNetCamId As String In NetCams.Keys
          Dim NetCamDevice As NetCamDevice = NetCams(strNetCamId)

          Dim strOptionValue As String = strNetCamId
          Dim strOptionName As String = strOptionValue & " [" & NetCamDevice.Name & "]"
          jqNetCam.AddItem(strOptionName, strOptionValue, (ActionSelected = strOptionValue))
        Next
        stb.Append("Select NetCam:")
        stb.Append(jqNetCam.Build)
        stb.Append("<br/>")

        '
        ' Start Snapshot Count
        '
        ActionSelected = IIf(action.Item("SnapshotCount").Length = 0, "1", action.Item("SnapshotCount"))
        actionId = String.Format("{0}{1}_{2}_{3}", actionName, "SnapshotCount", UID, sUnique)

        Dim jqParm2 As New clsJQuery.jqTextBox(actionId, "text", ActionSelected, Pagename, 10, True)
        stb.Append("Number of snapshots")
        stb.Append(jqParm2.Build)
        stb.Append("<br/>")

        '
        ' Start Snapshot Delay
        '
        ActionSelected = IIf(action.Item("SnapshotDelay").Length = 0, "1", action.Item("SnapshotDelay"))
        actionId = String.Format("{0}{1}_{2}_{3}", actionName, "SnapshotDelay", UID, sUnique)

        Dim jqParm3 As New clsJQuery.jqTextBox(actionId, "text", ActionSelected, Pagename, 10, True)
        stb.Append("Delay between snapshots (in seconds)")
        stb.Append(jqParm3.Build)
        stb.Append("<br/>")

        '
        ' Start E-mail Notification Recipient
        '
        ActionSelected = IIf(action.Item("EmailRecipient").Length = 0, "", action.Item("EmailRecipient"))
        actionId = String.Format("{0}{1}_{2}_{3}", actionName, "EmailRecipient", UID, sUnique)

        Dim jqParm4 As New clsJQuery.jqTextBox(actionId, "text", ActionSelected, Pagename, 25, True)
        stb.Append("E-mail Notification Recipient (leave blank to use the default)")
        stb.Append(jqParm4.Build)
        stb.Append("<br/>")

        '
        ' Start E-mail Attachment Options
        '
        ActionSelected = action.Item("Attachment")
        actionId = String.Format("{0}{1}_{2}_{3}", actionName, "Attachment", UID, sUnique)

        Dim jqAttachment As New clsJQuery.jqDropList(actionId, Pagename, True)
        jqAttachment.autoPostBack = True

        jqAttachment.AddItem("(Select Option)", "", (ActionSelected = ""))
        For Each strKey As String In AttachmentTypes.Keys
          Dim strOptionValue As String = strKey
          Dim strOptionName As String = AttachmentTypes(strKey)
          jqAttachment.AddItem(strOptionName, strOptionValue, (ActionSelected = strOptionValue))
        Next

        stb.Append("E-mail Notification Attachment Option")
        stb.Append(jqAttachment.Build)
        stb.Append("<br/>")

        '
        ' Start Send E-mail Notification
        '
        ActionSelected = action.Item("Notification")
        actionId = String.Format("{0}{1}_{2}_{3}", actionName, "Notification", UID, sUnique)

        Dim jqNotification As New clsJQuery.jqCheckBox(actionId, "Send E-mail Notification", Pagename, True, True)
        jqNotification.checked = ActionSelected = "on"
        stb.Append(jqNotification.Build)
        stb.Append("<br/>")

        '
        ' Start Compress Event Snapshots
        '
        ActionSelected = action.Item("Compress")
        actionId = String.Format("{0}{1}_{2}_{3}", actionName, "Compress", UID, sUnique)

        Dim jqCompress As New clsJQuery.jqCheckBox(actionId, "Compress Event Snapshots", Pagename, True, True)
        jqCompress.checked = ActionSelected = "on"
        stb.Append(jqCompress.Build)
        stb.Append("<br/>")

        '
        ' Start Upload Compressed File and/or Video to FTP Server
        '
        ActionSelected = action.Item("Upload")
        actionId = String.Format("{0}{1}_{2}_{3}", actionName, "Upload", UID, sUnique)

        Dim jqUpload As New clsJQuery.jqCheckBox(actionId, "Upload Compressed File and/or Video to FTP Server", Pagename, True, True)
        jqUpload.checked = ActionSelected = "on"
        stb.Append(jqUpload.Build)
        stb.Append("<br/>")

        '
        ' Start Archive Compressed File and/or Video
        '
        ActionSelected = action.Item("Archive")
        actionId = String.Format("{0}{1}_{2}_{3}", actionName, "Archive", UID, sUnique)

        Dim jqArchive As New clsJQuery.jqCheckBox(actionId, "Archive Compressed File and/or Video", Pagename, True, True)
        jqArchive.checked = ActionSelected = "on"
        stb.Append(jqArchive.Build)
        stb.Append("<br/>")

      Case NetCamActions.TakeSnapshot
        Dim actionName As String = GetEnumName(NetCamActions.TakeSnapshot)

        Dim ActionSelected As String = action.Item("NetCam")

        Dim actionId As String = String.Format("{0}{1}_{2}_{3}", actionName, "NetCam", UID, sUnique)

        Dim jqNetCam As New clsJQuery.jqDropList(actionId, Pagename, True)
        jqNetCam.autoPostBack = True

        jqNetCam.AddItem("(Select Network Camera)", "", (ActionSelected = ""))
        For Each strNetCamId As String In NetCams.Keys
          Dim NetCamDevice As NetCamDevice = NetCams(strNetCamId)

          Dim strOptionValue As String = strNetCamId
          Dim strOptionName As String = strOptionValue & " [" & NetCamDevice.Name & "]"
          jqNetCam.AddItem(strOptionName, strOptionValue, (ActionSelected = strOptionValue))
        Next
        stb.Append("Select NetCam:")
        stb.Append(jqNetCam.Build)
        stb.Append("<br/>")

      Case NetCamActions.CGIAction
        Dim actionName As String = GetEnumName(NetCamActions.CGIAction)

        Dim ActionSelected As String = action.Item("NetCam")

        Dim actionId As String = String.Format("{0}{1}_{2}_{3}", actionName, "NetCam", UID, sUnique)

        Dim jqNetCam As New clsJQuery.jqDropList(actionId, Pagename, True)
        jqNetCam.autoPostBack = True

        jqNetCam.AddItem("(Select Network Camera)", "", (ActionSelected = ""))
        For Each strNetCamId As String In NetCams.Keys
          Dim NetCamDevice As NetCamDevice = NetCams(strNetCamId)

          Dim strOptionValue As String = strNetCamId
          Dim strOptionName As String = strOptionValue & " [" & NetCamDevice.Name & "]"
          jqNetCam.AddItem(strOptionName, strOptionValue, (ActionSelected = strOptionValue))
        Next
        stb.Append("Select NetCam:")
        stb.Append(jqNetCam.Build)
        stb.Append("<br/>")

        '
        ' Start Query String
        '
        ActionSelected = IIf(action.Item("QueryString").Length = 0, "", action.Item("QueryString"))
        actionId = String.Format("{0}{1}_{2}_{3}", actionName, "QueryString", UID, sUnique)

        Dim jqQueryString As New clsJQuery.jqTextBox(actionId, "text", ActionSelected, Pagename, 25, True)
        stb.Append("CGI Command (Query String)")
        stb.Append(jqQueryString.Build)
        stb.Append("<br/>")

        '
        ' Start Wait Seconds
        '
        ActionSelected = IIf(action.Item("CommandWait").Length = 0, "0", action.Item("CommandWait"))
        actionId = String.Format("{0}{1}_{2}_{3}", actionName, "CommandWait", UID, sUnique)

        Dim jqWaitSeconds As New clsJQuery.jqTextBox(actionId, "text", ActionSelected, Pagename, 10, True)
        stb.Append("Wait for command to complete (in seconds)")
        stb.Append(jqWaitSeconds.Build)
        stb.Append("<br/>")

      Case NetCamActions.PurgeSnapshotEvents
        Dim actionName As String = GetEnumName(NetCamActions.PurgeSnapshotEvents)

        Dim ActionSelected As String = action.Item("NetCam")

        Dim actionId As String = String.Format("{0}{1}_{2}_{3}", actionName, "NetCam", UID, sUnique)

        Dim jqNetCam As New clsJQuery.jqDropList(actionId, Pagename, True)
        jqNetCam.autoPostBack = True

        jqNetCam.AddItem("(Select Network Camera)", "", (ActionSelected = ""))
        For Each strNetCamId As String In NetCams.Keys
          Dim NetCamDevice As NetCamDevice = NetCams(strNetCamId)

          Dim strOptionValue As String = strNetCamId
          Dim strOptionName As String = strOptionValue & " [" & NetCamDevice.Name & "]"
          jqNetCam.AddItem(strOptionName, strOptionValue, (ActionSelected = strOptionValue))
        Next
        stb.Append("Select NetCam:")
        stb.Append(jqNetCam.Build)
        stb.Append("<br/>")

        '
        ' Start Archive Compressed File and/or Video
        '
        ActionSelected = IIf(action.Item("EventsToKeep").Length = 0, "50", action.Item("EventsToKeep"))
        actionId = String.Format("{0}{1}_{2}_{3}", actionName, "EventsToKeep", UID, sUnique)

        Dim jqEventsToKeep As New clsJQuery.jqTextBox(actionId, "text", ActionSelected, Pagename, 10, True)
        stb.Append("Number of Snapshot Events to Keep")
        stb.Append(jqEventsToKeep.Build)

    End Select

    Return stb.ToString

  End Function

  ''' <summary>
  ''' When a user edits your event actions in the HomeSeer events, this function is called to process the selections.
  ''' </summary>
  ''' <param name="PostData"></param>
  ''' <param name="ActInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ActionProcessPostUI(ByVal PostData As Collections.Specialized.NameValueCollection, _
                                      ByVal ActInfo As IPlugInAPI.strTrigActInfo) As IPlugInAPI.strMultiReturn

    Dim Ret As New HomeSeerAPI.IPlugInAPI.strMultiReturn

    Dim UID As Integer = ActInfo.UID
    Dim TANumber As Integer = ActInfo.TANumber

    ' When plug-in calls such as ...BuildUI, ...ProcessPostUI, or ...FormatUI are called and there is
    ' feedback or an error condition that needs to be reported back to the user, this string field 
    ' can contain the message to be displayed to the user in HomeSeer UI.  This field is cleared by
    ' HomeSeer after it is displayed to the user.
    Ret.sResult = ""

    ' We cannot be passed info ByRef from HomeSeer, so turn right around and return this same value so that if we want, 
    '   we can exit here by returning 'Ret', all ready to go.  If in this procedure we need to change DataOut or TrigInfo,
    '   we can still do that.
    Ret.DataOut = ActInfo.DataIn
    Ret.TrigActInfo = ActInfo

    If PostData Is Nothing Then Return Ret
    If PostData.Count < 1 Then Return Ret

    '
    ' DeSerializeObject
    '
    Dim action As New action
    If Not (ActInfo.DataIn Is Nothing) Then
      DeSerializeObject(ActInfo.DataIn, action)
    End If
    action.uid = UID

    Dim parts As Collections.Specialized.NameValueCollection = PostData

    Try

      Select Case TANumber
        Case NetCamActions.EventSnapshots
          Dim actionName As String = GetEnumName(NetCamActions.EventSnapshots)

          For Each sKey As String In parts.Keys

            If sKey Is Nothing Then Continue For
            If String.IsNullOrEmpty(sKey.Trim) Then Continue For

            Select Case True
              Case InStr(sKey, actionName & "NetCam_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                action.Item("NetCam") = ActionValue

              Case InStr(sKey, actionName & "SnapshotCount_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                action.Item("SnapshotCount") = ActionValue

              Case InStr(sKey, actionName & "SnapshotDelay_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                action.Item("SnapshotDelay") = ActionValue

              Case InStr(sKey, actionName & "EmailRecipient_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                action.Item("EmailRecipient") = ActionValue

              Case InStr(sKey, actionName & "Attachment_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                action.Item("Attachment") = ActionValue

              Case InStr(sKey, actionName & "Notification_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                action.Item("Notification") = ActionValue

              Case InStr(sKey, actionName & "Compress_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                action.Item("Compress") = ActionValue

              Case InStr(sKey, actionName & "Upload_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                action.Item("Upload") = ActionValue

              Case InStr(sKey, actionName & "Archive_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                action.Item("Archive") = ActionValue

            End Select
          Next

        Case NetCamActions.TakeSnapshot
          Dim actionName As String = GetEnumName(NetCamActions.TakeSnapshot)

          For Each sKey As String In parts.Keys

            If sKey Is Nothing Then Continue For
            If String.IsNullOrEmpty(sKey.Trim) Then Continue For

            Select Case True
              Case InStr(sKey, actionName & "NetCam_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                action.Item("NetCam") = ActionValue

            End Select
          Next

        Case NetCamActions.CGIAction
          Dim actionName As String = GetEnumName(NetCamActions.CGIAction)

          For Each sKey As String In parts.Keys

            If sKey Is Nothing Then Continue For
            If String.IsNullOrEmpty(sKey.Trim) Then Continue For

            Select Case True
              Case InStr(sKey, actionName & "NetCam_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                action.Item("NetCam") = ActionValue

              Case InStr(sKey, actionName & "QueryString_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                action.Item("QueryString") = ActionValue

              Case InStr(sKey, actionName & "CommandWait_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                action.Item("CommandWait") = ActionValue

            End Select
          Next

        Case NetCamActions.PurgeSnapshotEvents
          Dim actionName As String = GetEnumName(NetCamActions.PurgeSnapshotEvents)

          For Each sKey As String In parts.Keys

            If sKey Is Nothing Then Continue For
            If String.IsNullOrEmpty(sKey.Trim) Then Continue For

            Select Case True
              Case InStr(sKey, actionName & "NetCam_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                action.Item("NetCam") = ActionValue

              Case InStr(sKey, actionName & "EventsToKeep_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                action.Item("EventsToKeep") = ActionValue

            End Select
          Next

      End Select

      ' The serialization data for the plug-in object cannot be 
      ' passed ByRef which means it can be passed only one-way through the interface to HomeSeer.
      ' If the plug-in receives DataIn, de-serializes it into an object, and then makes a change 
      ' to the object, this is where the object can be serialized again and passed back to HomeSeer
      ' for storage in the HomeSeer database.

      ' SerializeObject
      If Not SerializeObject(action, Ret.DataOut) Then
        Ret.sResult = IFACE_NAME & " Error, Serialization failed. Signal Action not added."
        Return Ret
      End If

    Catch ex As Exception
      Ret.sResult = "ERROR, Exception in Action UI of " & IFACE_NAME & ": " & ex.Message
      Return Ret
    End Try

    ' All OK
    Ret.sResult = ""
    Return Ret
  End Function

  ''' <summary>
  ''' Determines if our action is proplery configured
  ''' </summary>
  ''' <param name="ActInfo"></param>
  ''' <returns>Return TRUE if the given action is configured properly</returns>
  ''' <remarks>There may be times when a user can select invalid selections for the action and in this case you would return FALSE so HomeSeer will not allow the action to be saved.</remarks>
  Public Function ActionConfigured(ByVal ActInfo As IPlugInAPI.strTrigActInfo) As Boolean

    Dim Configured As Boolean = True
    Dim UID As String = ActInfo.UID.ToString

    Dim action As New action
    If Not (ActInfo.DataIn Is Nothing) Then
      DeSerializeObject(ActInfo.DataIn, action)
    End If

    Select Case ActInfo.TANumber
      Case NetCamActions.EventSnapshots
        If action.Item("NetCam") = "" Then Configured = False
        If action.Item("SnapshotCount") = "" Then Configured = False
        If action.Item("SnapshotDelay") = "" Then Configured = False
        'If action.Item("EmailRecipient") = "" Then Configured = False
        If action.Item("Attachment") = "" Then Configured = False
        'If action.Item("Notification") = "" Then Configured = False
        'If action.Item("Compress") = "" Then Configured = False
        'If action.Item("Upload") = "" Then Configured = False
        'If action.Item("Archive") = "" Then Configured = False

      Case NetCamActions.TakeSnapshot
        If action.Item("NetCam") = "" Then Configured = False

      Case NetCamActions.CGIAction
        If action.Item("NetCam") = "" Then Configured = False
        If action.Item("QueryString") = "" Then Configured = False
        If action.Item("CommandWait") = "" Then Configured = False

      Case NetCamActions.PurgeSnapshotEvents
        If action.Item("NetCam") = "" Then Configured = False
        If action.Item("EventsToKeep") = "" Then Configured = False

    End Select

    Return Configured

  End Function

  ''' <summary>
  ''' After the action has been configured, this function is called in your plugin to display the configured action
  ''' </summary>
  ''' <param name="ActInfo"></param>
  ''' <returns>Return text that describes the given action.</returns>
  ''' <remarks></remarks>
  Public Function ActionFormatUI(ByVal ActInfo As IPlugInAPI.strTrigActInfo) As String
    Dim stb As New StringBuilder

    Dim UID As String = ActInfo.UID.ToString

    Dim action As New action
    If Not (ActInfo.DataIn Is Nothing) Then
      DeSerializeObject(ActInfo.DataIn, action)
    End If

    Select Case ActInfo.TANumber
      Case NetCamActions.EventSnapshots
        If action.uid <= 0 Then
          stb.Append("Action has not been properly configured.")
        Else
          Dim strActionName = GetEnumDescription(NetCamActions.EventSnapshots)

          Dim strNetCamId As String = action.Item("NetCam")
          Dim strSnapshotCount As String = action.Item("SnapshotCount")
          Dim strSnapshotDelay As String = action.Item("SnapshotDelay")
          Dim strEmailRecipient As String = action.Item("EmailRecipient")
          Dim strEmailAttachment As String = action.Item("Attachment")
          Dim bEmailed As Boolean = IIf(action.Item("Notification") = "on", True, False)
          Dim bCompressed As Boolean = IIf(action.Item("Compress") = "on", True, False)
          Dim bUploaded As Boolean = IIf(action.Item("Upload") = "on", True, False)
          Dim bArchived As Boolean = IIf(action.Item("Archive") = "on", True, False)

          If NetCams.ContainsKey(strNetCamId) Then
            Dim NetCamDevice As NetCamDevice = NetCams(strNetCamId)
            strNetCamId = String.Format("{0} [{1}]", strNetCamId, NetCamDevice.Name)
          End If

          If AttachmentTypes.ContainsKey(strEmailAttachment) Then
            strEmailAttachment = AttachmentTypes(strEmailAttachment)
          End If

          stb.AppendFormat("{0}<br>" & _
                          "Network Camera: {1}<br>" & _
                          "Snapshot Count: {2}<br>" & _
                          "Delay Between Shapshots: {3} second(s)<br>" & _
                          "Alternate E-Mail Recipient: {4}<br>" & _
                          "E-Mail Attachment: {5}<br>" & _
                          "E-Mailed: {6}<br>" & _
                          "Compressed: {7}<br>" & _
                          "Uploaded: {8}<br>" & _
                          "Archived: {9}",
                          strActionName, _
                          strNetCamId, _
                          strSnapshotCount, _
                          strSnapshotDelay, _
                          strEmailRecipient, _
                          strEmailAttachment, _
                          bEmailed.ToString, _
                          bCompressed.ToString, _
                          bUploaded.ToString, _
                          bArchived.ToString)
        End If

      Case NetCamActions.TakeSnapshot
        If action.uid <= 0 Then
          stb.Append("Action has not been properly configured.")
        Else
          Dim strActionName = GetEnumDescription(NetCamActions.TakeSnapshot)

          Dim strNetCamId As String = action.Item("NetCam")

          If NetCams.ContainsKey(strNetCamId) Then
            Dim NetCamDevice As NetCamDevice = NetCams(strNetCamId)
            strNetCamId = String.Format("{0} [{1}]", strNetCamId, NetCamDevice.Name)
          End If

          stb.AppendFormat("{0} on {1}", strActionName, strNetCamId)
        End If

      Case NetCamActions.CGIAction
        If action.uid <= 0 Then
          stb.Append("Action has not been properly configured.")
        Else
          Dim strActionName = GetEnumDescription(NetCamActions.CGIAction)

          Dim strNetCamId As String = action.Item("NetCam")
          Dim strQueryString As String = action.Item("QueryString")
          Dim strCommandWait As String = action.Item("CommandWait")

          If NetCams.ContainsKey(strNetCamId) Then
            Dim NetCamDevice As NetCamDevice = NetCams(strNetCamId)
            strNetCamId = String.Format("{0} [{1}]", strNetCamId, NetCamDevice.Name)
          End If

          stb.AppendFormat("{0}<br>Network Camera: {1}<br>Command: {2}<br>Wait for command: {3} second(s)", strActionName, strNetCamId, strQueryString, strCommandWait)
        End If

      Case NetCamActions.PurgeSnapshotEvents
        If action.uid <= 0 Then
          stb.Append("Action has not been properly configured.")
        Else
          Dim strActionName = GetEnumDescription(NetCamActions.PurgeSnapshotEvents)

          Dim strNetCamId As String = action.Item("NetCam")
          Dim strEventsToKeep As String = action.Item("EventsToKeep")

          If NetCams.ContainsKey(strNetCamId) Then
            Dim NetCamDevice As NetCamDevice = NetCams(strNetCamId)
            strNetCamId = String.Format("{0} [{1}]", strNetCamId, NetCamDevice.Name)
          End If

          stb.AppendFormat("{0}<br>Network Camera: {1}<br>Keep {2} snapshot events", strActionName, strNetCamId, strEventsToKeep)
        End If

    End Select

    Return stb.ToString
  End Function

  ''' <summary>
  ''' Handles the HomeSeer Event Action
  ''' </summary>
  ''' <param name="ActInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function HandleAction(ByVal ActInfo As IPlugInAPI.strTrigActInfo) As Boolean

    Dim UID As String = ActInfo.UID.ToString

    Try

      Dim action As New action
      If Not (ActInfo.DataIn Is Nothing) Then
        DeSerializeObject(ActInfo.DataIn, action)
      Else
        Return False
      End If

      Select Case ActInfo.TANumber
        Case NetCamActions.EventSnapshots
          Dim strActionName = GetEnumDescription(NetCamActions.EventSnapshots)
          Dim strNetCamId As String = action.Item("NetCam")
          Dim strSnapshotCount As String = action.Item("SnapshotCount")
          Dim strSnapshotDelay As String = action.Item("SnapshotDelay")
          Dim strEmailRecipient As String = action.Item("EmailRecipient")
          Dim strEmailAttachment As String = action.Item("Attachment")
          Dim bEmailed As Boolean = IIf(action.Item("Notification") = "on", True, False)
          Dim bCompressed As Boolean = IIf(action.Item("Compress") = "on", True, False)
          Dim bUploaded As Boolean = IIf(action.Item("Upload") = "on", True, False)
          Dim bArchived As Boolean = IIf(action.Item("Archive") = "on", True, False)

          If NetCams.ContainsKey(strNetCamId) = True Then

            Dim iSnapshotCount As Integer = IIf(IsNumeric(strSnapshotCount) = True, Val(strSnapshotCount), 1)
            Dim iSnapshotDelay As Double = IIf(IsNumeric(strSnapshotDelay) = True, Val(strSnapshotDelay), 1)
            Dim iAttachmentType As Integer = IIf(IsNumeric(strEmailAttachment) = True, Val(strEmailAttachment), 1)

            NetCamEventSnapshot(strNetCamId, _
                                iSnapshotCount, _
                                iSnapshotDelay, _
                                strEmailRecipient, _
                                iAttachmentType, _
                                bEmailed, _
                                bCompressed, _
                                bUploaded, _
                                bArchived)

          End If

        Case NetCamActions.TakeSnapshot
          Dim strNetCamId As String = action.Item("NetCam")

          NetCamSnapshot(strNetCamId)

        Case NetCamActions.CGIAction
          Dim strNetCamId As String = action.Item("NetCam")
          Dim strQueryString As String = action.Item("QueryString")
          Dim strCommandWait As String = action.Item("CommandWait")

          If NetCams.ContainsKey(strNetCamId) = True Then

            Dim iCommandWait As Double = IIf(IsNumeric(strCommandWait) = True, Val(strCommandWait), 10)
            NetCamCGI(strNetCamId, strQueryString, iCommandWait)

          End If

        Case NetCamActions.PurgeSnapshotEvents
          Dim strNetCamId As String = action.Item("NetCam")
          Dim strEventsToKeep As String = action.Item("EventsToKeep")

          Dim iEventCount As Integer = IIf(IsNumeric(strEventsToKeep) = True, Val(strEventsToKeep), gSnapshotEventMax)
          PurgeSnapshotEvents(strNetCamId, iEventCount)

      End Select

    Catch pEx As Exception
      '
      ' Process Program Exception
      '
      hs.WriteLog(IFACE_NAME, "Error executing action: " & pEx.Message)
    End Try

    Return True

  End Function

#End Region

#End Region

End Module

Public Enum NetCamTriggers
  <Description("Alarm Trigger")> _
  AlarmTrigger = 1
End Enum

Public Enum NetCamActions
  <Description("Event Snapshots")> _
  EventSnapshots = 1
  <Description("Take Snapshot")> _
  TakeSnapshot = 2
  <Description("Purge Snapshots Events")> _
  PurgeSnapshotEvents = 3
  <Description("CGI Action")> _
  CGIAction = 4
End Enum