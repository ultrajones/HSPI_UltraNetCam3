Imports System.IO
Imports System.Data.Common
Imports System.Text.RegularExpressions

Module Database

  Public DBConnectionMain As SQLite.SQLiteConnection  ' Our main database connection
  Public DBConnectionTemp As SQLite.SQLiteConnection  ' Our temp database connection

  Public gDBInsertSuccess As ULong = 0            ' Tracks DB insert success
  Public gDBInsertFailure As ULong = 0            ' Tracks DB insert success

  Public bDBInitialized As Boolean = False        ' Indicates if database successfully initialized

  Public SyncLockMain As New Object
  Public SyncLockTemp As New Object

#Region "Database Initilization"

  ''' <summary>
  ''' Initializes the database
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function InitializeMainDatabase() As Boolean

    Dim strMessage As String = ""               ' Holds informational messages
    Dim bSuccess As Boolean = False             ' Indicate default success

    WriteMessage("Entered InitializeMainDatabase() function.", MessageType.Debug)

    Try
      '
      ' Close database if it's open
      '
      If Not DBConnectionMain Is Nothing Then
        If CloseDBConn(DBConnectionMain) = False Then
          Throw New Exception("An existing database connection could not be closed.")
        End If
      End If

      '
      ' Create the database directory if it does not exist
      '
      Dim databaseDir As String = FixPath(String.Format("{0}\Data\{1}\", hs.GetAppPath, IFACE_NAME.ToLower))
      If Directory.Exists(databaseDir) = False Then
        Directory.CreateDirectory(databaseDir)
      End If

      '
      ' Determine the database filename
      '
      Dim strDataSource As String = FixPath(String.Format("{0}\Data\{1}\{1}.db3", hs.GetAppPath(), IFACE_NAME.ToLower))

      '
      ' Determine the database provider factory and connection string
      '
      Dim strConnectionString As String = String.Format("Data Source={0}; Version=3;", strDataSource)

      '
      ' Attempt to open the database connection
      '
      bSuccess = OpenDBConn(DBConnectionMain, strConnectionString)

    Catch pEx As Exception
      '
      ' Process program exception
      '
      bSuccess = False
      Call ProcessError(pEx, "InitializeDatabase()")
    End Try

    Return bSuccess

  End Function

  '------------------------------------------------------------------------------------
  'Purpose: Initializes the temporary database
  'Inputs:  None
  'Outputs: True or False indicating if database was initialized
  '------------------------------------------------------------------------------------
  Public Function InitializeTempDatabase() As Boolean

    Dim strMessage As String = ""               ' Holds informational messages
    Dim bSuccess As Boolean = False             ' Indicate default success

    strMessage = "Entered InitializeChannelDatabase() function."
    Call WriteMessage(strMessage, MessageType.Debug)

    Try
      '
      ' Close database if it's open
      '
      If Not DBConnectionTemp Is Nothing Then
        If CloseDBConn(DBConnectionTemp) = False Then
          Throw New Exception("An existing database connection could not be closed.")
        End If
      End If

      '
      ' Determine the database filename
      '
      Dim dtNow As DateTime = DateTime.Now
      Dim iHour As Integer = dtNow.Hour
      Dim strDBDate As String = iHour.ToString.PadLeft(2, "0")
      Dim strDataSource As String = FixPath(String.Format("{0}\data\{1}\{1}_{2}.db3", hs.GetAppPath(), IFACE_NAME.ToLower, strDBDate))

      '
      ' Determine the database provider factory and connection string
      '
      Dim strConnectionString As String = String.Format("Data Source={0}; Version=3; Journal Mode=Off;", strDataSource)

      '
      ' Attempt to open the database connection
      '
      bSuccess = OpenDBConn(DBConnectionTemp, strConnectionString)

    Catch pEx As Exception
      '
      ' Process program exception
      '
      bSuccess = False
      Call ProcessError(pEx, "InitializeTempDatabase()")
    End Try

    Return bSuccess

  End Function

  ''' <summary>
  ''' Opens a connection to the database
  ''' </summary>
  ''' <param name="objConn"></param>
  ''' <param name="strConnectionString"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function OpenDBConn(ByRef objConn As SQLite.SQLiteConnection,
                              ByVal strConnectionString As String) As Boolean

    Dim strMessage As String = ""               ' Holds informational messages
    Dim bSuccess As Boolean = False             ' Indicate default success

    WriteMessage("Entered OpenDBConn() function.", MessageType.Debug)

    Try
      '
      ' Open database connection
      '
      objConn = New SQLite.SQLiteConnection()
      objConn.ConnectionString = strConnectionString
      objConn.Open()

      '
      ' Run database vacuum
      '
      WriteMessage("Running SQLite database vacuum.", MessageType.Debug)
      Using MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

        MyDbCommand.Connection = DBConnectionMain
        MyDbCommand.CommandType = CommandType.Text
        MyDbCommand.CommandText = "VACUUM"
        MyDbCommand.ExecuteNonQuery()

        MyDbCommand.Dispose()
      End Using

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "OpenDBConn()")
    End Try

    '
    ' Determine database connection status
    '
    bSuccess = DBConnectionMain.State = ConnectionState.Open

    '
    ' Record connection state to HomeSeer log
    '
    If bSuccess = True Then
      strMessage = "Database initialization complete."
      Call WriteMessage(strMessage, MessageType.Debug)
    Else
      strMessage = "Database initialization failed using [" & strConnectionString & "]."
      Call WriteMessage(strMessage, MessageType.Debug)
    End If

    Return bSuccess

  End Function

  ''' <summary>
  ''' Closes database connection
  ''' </summary>
  ''' <param name="objConn"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function CloseDBConn(ByRef objConn As SQLite.SQLiteConnection) As Boolean

    Dim strMessage As String = ""               ' Holds informational messages
    Dim bSuccess As Boolean = False             ' Indicate default success

    WriteMessage("Entered CloseDBConn() function.", MessageType.Debug)

    Try
      '
      ' Attempt to the database
      '
      If DBConnectionMain.State <> ConnectionState.Closed Then
        DBConnectionMain.Close()
      End If

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "CloseDBConn()")
    End Try

    '
    ' Determine database connection status
    '
    bSuccess = DBConnectionMain.State = ConnectionState.Closed

    '
    ' Record connection state to HomeSeer log
    '
    If bSuccess = True Then
      strMessage = "Database connection closed successfuly."
      Call WriteMessage(strMessage, MessageType.Debug)
    Else
      strMessage = "Unable to close database; Try restarting HomeSeer."
      Call WriteMessage(strMessage, MessageType.Debug)
    End If

    Return bSuccess

  End Function

  ''' <summary>
  ''' Checks to ensure a table exists
  ''' </summary>
  ''' <param name="strTableName"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function CheckDatabaseTable(ByVal strTableName As String) As Boolean

    Dim strMessage As String = ""
    Dim bSuccess As Boolean = False

    Try
      '
      ' Build SQL delete statement
      '
      If Regex.IsMatch(strTableName, "tblNetCamDevices") = True Then
        '
        ' Retrieve schema information about tblTableName
        '
        Dim SchemaTable As DataTable = DBConnectionMain.GetSchema("Columns", New String() {Nothing, Nothing, strTableName})

        If SchemaTable.Rows.Count <> 0 Then
          WriteMessage("Table " & SchemaTable.Rows(0)!TABLE_NAME.ToString & " exists.", MessageType.Debug)
        Else
          '
          ' Create the table
          '
          WriteMessage("Creating " & strTableName & " table ...", MessageType.Debug)

          Using dbcmd As DbCommand = DBConnectionMain.CreateCommand()

            Dim sqlQueue As New Queue

            sqlQueue.Enqueue("CREATE TABLE tblNetCamDevices(" _
                            & "netcam_id INTEGER PRIMARY KEY," _
                            & "netcam_name varchar(255) NOT NULL," _
                            & "netcam_address varchar(255) NOT NULL," _
                            & "netcam_port INTEGER NOT NULL," _
                            & "netcam_type INTEGER DEFAULT 1 NOT NULL," _
                            & "auth_user varchar(255)," _
                            & "auth_pass varchar(255)" _
                          & ")")

            sqlQueue.Enqueue("CREATE UNIQUE INDEX idxUNIQUE ON tblNetCamDevices(netcam_name, netcam_address, netcam_port, netcam_type)")

            While sqlQueue.Count > 0
              Dim strSQL As String = sqlQueue.Dequeue

              dbcmd.Connection = DBConnectionMain
              dbcmd.CommandType = CommandType.Text
              dbcmd.CommandText = strSQL

              Dim iRecordsAffected As Integer = dbcmd.ExecuteNonQuery()
              If iRecordsAffected <> 1 Then
                'Throw New Exception("Database schemea update failed due to error.")
              End If
            End While

            dbcmd.Dispose()
          End Using

        End If

      ElseIf String.Compare(strTableName, "tblNetCamTypes", True) = 0 Then
        '
        ' Retrieve schema information about tblTableName
        '
        Dim SchemaTable As DataTable = DBConnectionMain.GetSchema("Columns", New String() {Nothing, Nothing, strTableName})

        If SchemaTable.Rows.Count <> 0 Then
          WriteMessage("Table " & SchemaTable.Rows(0)!TABLE_NAME.ToString & " exists.", MessageType.Debug)
        Else
          WriteMessage("Creating " & strTableName & " table ...", MessageType.Informational)

          Dim dbcmd As DbCommand = DBConnectionMain.CreateCommand()

          Dim sqlQueue As New Queue

          sqlQueue.Enqueue("CREATE TABLE tblNetCamTypes(" _
                          & "netcam_type INTEGER PRIMARY KEY," _
                          & "netcam_vendor varchar(255) NOT NULL," _
                          & "netcam_model varchar(255) NOT NULL," _
                          & "snapshot_path varchar(255) NOT NULL," _
                          & "videostream_path varchar(255) NOT NULL" _
                        & ")")

          sqlQueue.Enqueue("CREATE UNIQUE INDEX idxUNIQUETYPE ON tblNetCamTypes(netcam_vendor, netcam_model)")
          sqlQueue.Enqueue("INSERT INTO tblNetCamTypes VALUES (1, 'Foscam', 'JPEG', '/snapshot.cgi', '/videostream.cgi')")
          sqlQueue.Enqueue("INSERT INTO tblNetCamTypes VALUES (2, 'Foscam', 'H.264', '/cgi-bin/CGIProxy.fcgi?cmd=snapPicture2&usr=$user&pwd=$pass', '/videostream.cgi')")

          While sqlQueue.Count > 0
            Dim strSQL As String = sqlQueue.Dequeue

            dbcmd.Connection = DBConnectionMain
            dbcmd.CommandType = CommandType.Text
            dbcmd.CommandText = strSQL

            Dim iRecordsAffected As Integer = dbcmd.ExecuteNonQuery()
            If iRecordsAffected <> 1 Then
              'Throw New Exception("Create table failed due to error.")
            End If
          End While

          dbcmd.Dispose()
        End If

      ElseIf String.Compare(strTableName, "tblNetCamControls", True) = 0 Then
        '
        ' Retrieve schema information about tblTableName
        '
        Dim SchemaTable As DataTable = DBConnectionMain.GetSchema("Columns", New String() {Nothing, Nothing, strTableName})

        If SchemaTable.Rows.Count <> 0 Then
          WriteMessage("Table " & SchemaTable.Rows(0)!TABLE_NAME.ToString & " exists.", MessageType.Debug)
        Else
          WriteMessage("Creating " & strTableName & " table ...", MessageType.Informational)

          Dim dbcmd As DbCommand = DBConnectionMain.CreateCommand()

          Dim sqlQueue As New Queue

          sqlQueue.Enqueue("CREATE TABLE tblNetCamControls(" _
                          & "control_id INTEGER PRIMARY KEY," _
                          & "netcam_type INTEGER DEFAULT 1 NOT NULL, " _
                          & "control_name varchar(255) NOT NULL," _
                          & "control_url varchar(255) NOT NULL" _
                        & ")")

          sqlQueue.Enqueue("CREATE UNIQUE INDEX idxUNIQUECONTROL ON tblNetCamControls(netcam_type, control_name)")

          While sqlQueue.Count > 0
            Dim strSQL As String = sqlQueue.Dequeue

            dbcmd.Connection = DBConnectionMain
            dbcmd.CommandType = CommandType.Text
            dbcmd.CommandText = strSQL

            Dim iRecordsAffected As Integer = dbcmd.ExecuteNonQuery()
            If iRecordsAffected <> 1 Then
              'Throw New Exception("Create table failed due to error.")
            End If
          End While

          dbcmd.Dispose()

          LoadNetCamControls()
        End If

      ElseIf String.Compare(strTableName, "tblNetCamEvents", True) = 0 Then
        '
        ' Retrieve schema information about tblTableName
        '
        Dim SchemaTable As DataTable = DBConnectionMain.GetSchema("Columns", New String() {Nothing, Nothing, strTableName})

        If SchemaTable.Rows.Count <> 0 Then
          WriteMessage("Table " & SchemaTable.Rows(0)!TABLE_NAME.ToString & " exists.", MessageType.Debug)
        Else
          WriteMessage("Creating " & strTableName & " table ...", MessageType.Informational)

          Dim dbcmd As DbCommand = DBConnectionMain.CreateCommand()

          Dim sqlQueue As New Queue

          sqlQueue.Enqueue("CREATE TABLE tblNetCamEvents(" _
                          & "event_id INTEGER PRIMARY KEY," _
                          & "event_flag INTEGER DEFAULT 0 NOT NULL," _
                          & "netcam_id INTEGER NOT NULL," _
                          & "start_ts INTEGER NOT NULL," _
                          & "end_ts INTEGER DEFAULT 0 NOT NULL," _
                          & "frames_requested INTEGER DEFAULT 0 NOT NULL," _
                          & "frames_completed INTEGER DEFAULT 0 NOT NULL," _
                          & "email_rcpt varchar(255)," _
                          & "email_attach INTEGER DEFAULT 0 NOT NULL," _
                          & "emailed INTEGER DEFAULT 0 NOT NULL," _
                          & "compressed INTEGER DEFAULT 0 NOT NULL," _
                          & "uploaded INTEGER DEFAULT 0 NOT NULL," _
                          & "archived INTEGER DEFAULT 0 NOT NULL" _
                        & ")")

          sqlQueue.Enqueue("CREATE INDEX idxSTARTDATE ON tblNetCamEvents(start_ts)")
          sqlQueue.Enqueue("CREATE INDEX idxFLAG ON tblNetCamEvents(event_flag)")
          sqlQueue.Enqueue("CREATE INDEX idxNETCAMID ON tblNetCamEvents(netcam_id)")

          While sqlQueue.Count > 0
            Dim strSQL As String = sqlQueue.Dequeue

            dbcmd.Connection = DBConnectionMain
            dbcmd.CommandType = CommandType.Text
            dbcmd.CommandText = strSQL

            Dim iRecordsAffected As Integer = dbcmd.ExecuteNonQuery()
            If iRecordsAffected <> 1 Then
              'Throw New Exception("Create table failed due to error.")
            End If

          End While

          dbcmd.Dispose()
        End If

      Else
        Throw New Exception(strTableName & " not currently supported.")
      End If

      bSuccess = True

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "CheckDatabaseTables()")
    End Try

    Return bSuccess

  End Function

#End Region

#Region "Database Date Formatting"

  ''' <summary>
  ''' dateTime as DateTime
  ''' </summary>
  ''' <param name="dateTime"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function ConvertDateTimeToEpoch(ByVal dateTime As DateTime) As Long

    Dim baseTicks As Long = 621355968000000000
    Dim tickResolution As Long = 10000000

    Return (dateTime.ToUniversalTime.Ticks - baseTicks) / tickResolution

  End Function

  ''' <summary>
  ''' Converts Epoch to datetime
  ''' </summary>
  ''' <param name="epochTicks"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function ConvertEpochToDateTime(ByVal epochTicks As Long) As DateTime

    '
    ' Create a new DateTime value based on the Unix Epoch
    '
    Dim converted As New DateTime(1970, 1, 1, 0, 0, 0, 0)

    '
    ' Return the value in string format
    '
    Return converted.AddSeconds(epochTicks).ToLocalTime

  End Function

#End Region

End Module

