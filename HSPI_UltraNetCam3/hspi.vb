Imports HomeSeerAPI
Imports Scheduler
Imports HSCF
Imports HSCF.Communication.ScsServices.Service

Imports System.Reflection
Imports System.Reflection.Assembly
Imports System.Diagnostics.FileVersionInfo

Imports System.Text
Imports System.Threading
Imports System.IO

Public Class HSPI
  Inherits ScsService
  Implements IPlugInAPI               ' This API is required for ALL plugins

  '
  ' Web Config Global Variables
  '
  Dim WebConfigPage As hspi_webconfig
  Dim WebPage As Object

  Public OurInstanceFriendlyName As String = ""
  Public instance As String = ""

  '
  ' Define Class Global Variables
  '
  Public DiscoveryThread As Thread
  Public SendDiscoveryThread As Thread
  Public SnapshotFileMaintenanceThread As Thread
  Public RefreshSnapshotThread As Thread

#Region "HSPI - Base Plugin API"

  ''' <summary>
  ''' Probably one of the most important properties, the Name function in your plug-in is what the plug-in is identified with at all times.  
  ''' The filename of the plug-in is irrelevant other than when HomeSeer is searching for plug-in files, but the Name property is key to many things, 
  ''' including how plug-in created triggers and actions are stored by HomeSeer.  
  ''' If this property is changed from one version of the plug-in to the next, all triggers, actions, and devices created by the plug-in will have to be re-created by the user.  
  ''' Please try to keep the Name property value short, e.g. 14 to 16 characters or less.  Web pages, trigger and action forms created by your plug-in can use a longer, 
  ''' more elaborate name if you so desire.  In the sample plug-ins, the constant IFACE_NAME is commonly used in the program to return the name of the plug-in. 
  ''' No spaces or special characters are allowed other than a dash or underscore.
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property Name As String Implements HomeSeerAPI.IPlugInAPI.Name
    Get
      Return IFACE_NAME
    End Get
  End Property

  ''' <summary>
  ''' Return the API's that this plug-in supports. This is a bit field. All plug-ins must have bit 3 set for I/O.
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function Capabilities() As Integer Implements HomeSeerAPI.IPlugInAPI.Capabilities
    Return HomeSeerAPI.Enums.eCapabilities.CA_IO
  End Function

  ''' <summary>
  ''' This determines whether the plug-in is free, or is a licensed plug-in using HomeSeer's licensing service.  
  ''' Return a value of 1 for a free plug-in, a value of 2 indicates that the plug-in is licensed using HomeSeer's licensing.
  ''' </summary>
  ''' <returns>Integer</returns>
  ''' <remarks></remarks>
  Public Function AccessLevel() As Integer Implements HomeSeerAPI.IPlugInAPI.AccessLevel
    Return 2
  End Function

  ''' <summary>
  ''' Returns the instance name of this instance of the plug-in. Only valid if SupportsMultipleInstances returns TRUE. 
  ''' The instance is set when the plug-in is started, it is passed as a command line parameter. 
  ''' The initial instance name is set when a new instance is created on the HomeSeer interfaces page. 
  ''' A plug-in needs to associate this instance name with any local status that it is keeping for this instance. 
  ''' See the multiple instances section for more information.
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function InstanceFriendlyName() As String Implements HomeSeerAPI.IPlugInAPI.InstanceFriendlyName

    '
    ' Write the debug message
    '
    WriteMessage("Entered InstanceFriendlyName() function.", MessageType.Debug)

    Return Me.instance
  End Function

  ''' <summary>
  ''' Return TRUE if the plug-in supports multiple instances. 
  ''' The plug-in may be launched multiple times and will be passed a unique instance name as a command line parameter to the Main function. 
  ''' The plug-in then needs to associate all local status with this particular instance.
  ''' This feature is ideal for cases where multiple hardware modules need to be supported. 
  ''' For example, an single irrigation controller supports 8 zones but the user needs 16. 
  ''' They can add a second controller as a new instance to control 8 more zones. 
  ''' This assumes that the second controller would use a different COM port or IP address.
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function SupportsMultipleInstances() As Boolean Implements HomeSeerAPI.IPlugInAPI.SupportsMultipleInstances
    Return False
  End Function

  ''' <summary>
  ''' No Summary in API Docs
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function SupportsMultipleInstancesSingleEXE() As Boolean Implements HomeSeerAPI.IPlugInAPI.SupportsMultipleInstancesSingleEXE
    Return False
  End Function

  ''' <summary>
  ''' Your plug-in should return True if you wish to use HomeSeer's interfaces page of the configuration for the user to enter a serial port number for your plug-in to use.  
  ''' If enabled, HomeSeer will return this COM port number to your plug-in in the InitIO call.  
  ''' If you wish to have your own configuration UI for the serial port, or if your plug-in does not require a serial port, return False
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property HSCOMPort As Boolean Implements HomeSeerAPI.IPlugInAPI.HSCOMPort
    Get
      Return False
    End Get
  End Property

  ''' <summary>
  ''' SetIOMulti is called by HomeSeer when a device that your plug-in owns is controlled.  Your plug-in owns a device when it's INTERFACE property is set to the name of your plug-in.
  ''' The parameters passed to SetIOMulti are as follows - depending upon what generated the SetIO call, not all parameters will contain data.  
  ''' Be sure to test for "Is Nothing" before testing for values or your plug-in may generate an exception error when a variable passed is uninitialized. 
  ''' </summary>
  ''' <param name="colSend">This is a collection of CAPIControl objects, one object for each device that needs to be controlled. Look at the ControlValue property to get the value that device needs to be set to.</param>
  ''' <remarks></remarks>
  Public Sub SetIOMulti(colSend As System.Collections.Generic.List(Of HomeSeerAPI.CAPI.CAPIControl)) Implements HomeSeerAPI.IPlugInAPI.SetIOMulti

    '
    ' Write the debug message
    '
    WriteMessage(String.Format("Entered {0} {1}", "SetIOMulti", "Subroutine"), MessageType.Debug)

    Dim CC As CAPIControl
    For Each CC In colSend
      If CC Is Nothing Then Continue For

      WriteMessage(String.Format("SetIOMulti set value: {0}, type {1}, ref:{2}", CC.ControlValue.ToString, CC.ControlType.ToString, CC.Ref.ToString), MessageType.Debug)

      Try
        '
        ' Device exists, so lets get a reference to it
        '
        Dim dv As Scheduler.Classes.DeviceClass = hs.GetDeviceByRef(CC.Ref)
        If dv Is Nothing Then Continue For

        '
        ' Get the device type
        '
        Dim dv_type As String = dv.Device_Type_String(Nothing)
        Dim dv_addr As String = dv.Address(Nothing)
        Dim dv_ref As Integer = dv.Ref(Nothing)

        '
        ' Process the SetIO action
        '
        Select Case dv_type
          Case "NetCam"
            Dim dv_value As Integer = Int32.Parse(CC.ControlValue)
            Select Case dv_value
              Case -1
                '
                ' Send NetCam Snapshot
                '
                Dim strNetCamId As String = dv_addr

                NetCamSnapshot(strNetCamId)
                'SetDeviceValue(dv_addr, 0)

              Case Is > 0
                '
                ' Send NetCam CGI command
                '
                Dim strNetCamId As String = dv_addr
                Dim strControlURL As String = GetNetCamControlURL(dv_value)
                NetCamCGI(strNetCamId, strControlURL, 0)
                'SetDeviceValue(dv_addr, 0)

            End Select

          Case Else
            '
            ' Write warning message
            '
            WriteMessage(String.Format("Received unsupported SetIOMulti action for {0}", dv_type), MessageType.Warning)

        End Select

      Catch pEx As Exception
        '
        ' Process program exception
        '
        ProcessError(pEx, "SetIOMulti")
      End Try

    Next

  End Sub

  ''' <summary>
  ''' HomeSeer may call this function at any time to get the status of the plug-in. Normally it is displayed on the Interfaces page.
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function InterfaceStatus() As HomeSeerAPI.IPlugInAPI.strInterfaceStatus Implements HomeSeerAPI.IPlugInAPI.InterfaceStatus

    Dim es As New IPlugInAPI.strInterfaceStatus
    es.intStatus = IPlugInAPI.enumInterfaceStatus.OK

    Return es

  End Function

  ''' <summary>
  ''' If your plugin is set to start when HomeSeer starts, or is enabled from the interfaces page, then this function will be called to initialize your plugin. 
  ''' If you returned TRUE from HSComPort then the port number as configured in HomeSeer will be passed to this function. 
  ''' Here you should initialize your plugin fully. The hs object is available to you to call the HomeSeer scripting API as well as the callback object so you can call into the HomeSeer plugin API.  
  ''' HomeSeer's startup routine waits for this function to return a result, so it is important that you try to exit this procedure quickly.  
  ''' If your hardware has a long initialization process, you can check the configuration in InitIO and if everything is set up correctly, start a separate thread to initialize the hardware and exit InitIO.  
  ''' If you encounter an error, you can always use InterfaceStatus to indicate this.
  ''' </summary>
  ''' <param name="port"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function InitIO(ByVal port As String) As String Implements HomeSeerAPI.IPlugInAPI.InitIO

    Dim InterfaceStatus As String = ""
    Dim logMessage As String = ""

    Try
      '
      ' Write startup message to disk
      '
      WriteToDisk("HomeSeerV", hs.Version(), False)
      WriteToDisk("HomeSeerL", hs.IsLicensed())
      WriteToDisk("HomeSeerU", hs.SystemUpTime())

      '
      ' Write the debug message
      '
      WriteMessage("Entered InitIO() function.", MessageType.Debug)

      '
      ' Write the startup message
      '
      logMessage = String.Format("{0} version {1} initializing...", IFACE_NAME, HSPI.Version())
      Call WriteMessage(logMessage, MessageType.Informational)

      '
      ' Let's set the initial HomeSeer options for our plug-in
      '
      gLogLevel = Int32.Parse(hs.GetINISetting("Options", "LogLevel", LogLevel.Informational, gINIFile))

      gEventArchiveToDir = CBool(hs.GetINISetting("Archive", "ArchiveEnabled", False, gINIFile))
      gEventArchiveToFTP = CBool(hs.GetINISetting("FTPArchive", "ArchiveEnabled", False, gINIFile))
      gEventEmailNotification = CBool(hs.GetINISetting("EmailNotification", "EmailEnabled", False, gINIFile))
      gSnapshotEventMax = CInt(Val(hs.GetINISetting("Options", "SnapshotEventMax", gSnapshotEventMax, gINIFile)))
      gSnapshotRefreshInterval = CInt(Val(hs.GetINISetting("Options", "SnapshotRefreshInterval", gSnapshotRefreshInterval, gINIFile)))
      gSnapshotsMaxWidth = hs.GetINISetting("Options", "SnapshotsMaxWidth", gSnapshotsMaxWidth, gINIFile)
      gFoscamAutoDiscovery = CBool(hs.GetINISetting("Options", "FoscamAutoDiscovery", False, gINIFile))

      HSAppPath = hs.GetAppPath
      gSnapshotDirectory = FixPath(String.Format("{0}\html\images\hspi_ultranetcam3\snapshots", HSAppPath))

      bDBInitialized = Database.InitializeMainDatabase()
      If bDBInitialized = True Then
        '
        ' Check to see if the devices need to be refreshed (the version # has changed)
        '
        Dim bVersionChanged As Boolean = False
        If GetSetting("Settings", "Version", "") <> HSPI.Version() Then
          bVersionChanged = True
        End If

        If bVersionChanged Then
          '
          ' Run maintenance tasks
          '
          WriteMessage("Running version has changed, performing maintenance ...", MessageType.Informational)
          SaveSetting("Settings", "Version", HSPI.Version())
        End If

        '
        ' Check database tables
        '
        CheckDatabaseTable("tblNetCamDevices")
        CheckDatabaseTable("tblNetCamTypes")
        CheckDatabaseTable("tblNetCamEvents")
        CheckDatabaseTable("tblNetCamControls")

        '
        ' Define attachment type
        '
        AttachmentTypes.Clear()
        AttachmentTypes.Add("0", "Do not include an attachment")
        AttachmentTypes.Add("1", "Attach the latest snapshot of the event")
        AttachmentTypes.Add("2", "Attach the compressed file of the event")
        AttachmentTypes.Add("4", "Attach the MP4 video of the event")

        '
        ' Ensure the snapshot directory exist
        '
        If Directory.Exists(gSnapshotDirectory) = False Then
          Directory.CreateDirectory(gSnapshotDirectory)
        End If

        '
        ' Refresh the list of NetCams
        '
        RefreshNetCamList()

        '
        ' Remove orphaned directories
        '
        Dim RootDirInfo As New IO.DirectoryInfo(gSnapshotDirectory)

        For Each DirInfo As DirectoryInfo In RootDirInfo.GetDirectories
          If NetCams.ContainsKey(DirInfo.Name) = False Then
            '
            ' Delete the HomeSeer device
            '
            Dim dv_ref As Long = hs.DeviceExistsAddress(DirInfo.Name, False)
            If dv_ref > 0 Then
              Dim bResult As Boolean = hs.DeleteDevice(dv_ref)
              If bResult = True Then
                WriteMessage(String.Format("Deleted {0} orphaned HomeSeer device.", DirInfo.Name), MessageType.Warning)
              End If
            End If

            '
            ' Remove the camera snapshot directory
            '
            DeleteNetCamSnapshotDir(DirInfo.Name)
          End If
        Next

        '
        ' Start Plug-in Threads
        '
        If gFoscamAutoDiscovery = True Then
          '
          ' Start the DiscoveryThread thread
          '
          DiscoveryThread = New Thread(New ThreadStart(AddressOf DiscoveryBeacon))
          DiscoveryThread.Name = "Discovery"
          DiscoveryThread.Start()

          WriteMessage(String.Format("{0} Thread Started", DiscoveryThread.Name), MessageType.Debug)

          '
          ' Start the SendDiscoveryThread thread
          '
          SendDiscoveryThread = New Thread(New ThreadStart(AddressOf SendDiscoveryBeacon))
          SendDiscoveryThread.Name = "SendDiscovery"
          SendDiscoveryThread.Start()

          WriteMessage(String.Format("{0} Thread Started", SendDiscoveryThread.Name), MessageType.Debug)
        End If

        '
        ' Start the Snapshot File Maintenance Thread
        '
        SnapshotFileMaintenanceThread = New Thread(New ThreadStart(AddressOf SnapshotFileMaintenance))
        SnapshotFileMaintenanceThread.Name = "SnapshotFileMaintenance"
        SnapshotFileMaintenanceThread.Start()

        WriteMessage(String.Format("{0} Thread Started", SnapshotFileMaintenanceThread.Name), MessageType.Debug)

        '
        ' Start the Snapshot File Maintenance Thread
        '
        RefreshSnapshotThread = New Thread(New ThreadStart(AddressOf RefreshSnapshotsThread))
        RefreshSnapshotThread.Name = "RefreshSnapshotThread"
        RefreshSnapshotThread.Start()

        WriteMessage(String.Format("{0} Thread Started", RefreshSnapshotThread.Name), MessageType.Debug)

      End If

      '
      ' Register the Configuration Web Page
      '
      WebConfigPage = New hspi_webconfig(IFACE_NAME)
      WebConfigPage.hspiref = Me
      RegisterWebPage(WebConfigPage.PageName, LINK_TEXT, LINK_PAGE_TITLE, instance)

      '
      ' Register the ASXP Web Page
      '
      'RegisterASXPWebPage(LINK_URL, LINK_TEXT, LINK_PAGE_TITLE, instance)

      '
      ' Register the Help File
      '
      RegisterHelpPage(LINK_HELP, IFACE_NAME & " Help File", IFACE_NAME)

      '
      ' Register for events from homeseer
      '
      'callback.RegisterEventCB(Enums.HSEvent.VALUE_CHANGE, IFACE_NAME, "")

      '
      ' Write the startup message
      '
      logMessage = String.Format("{0} version {1} initialization complete.", IFACE_NAME, HSPI.Version())
      Call WriteMessage(logMessage, MessageType.Informational)

      '
      ' Indicate the plug-in has been initialized
      '
      gHSInitialized = True

    Catch pEx As Exception
      '
      ' Process the program exception
      '
      ProcessError(pEx, "InitIO")

      gHSInitialized = False
    End Try

    Return InterfaceStatus

  End Function

  ''' <summary>
  ''' When HomeSeer shuts down or a plug-in is disabled from the interfaces page this function is then called. 
  ''' You should terminate any threads that you started, close any COM ports or TCP connections and release memory. 
  ''' After you return from this function the plugin EXE will terminate and must be allowed to terminate cleanly.
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub ShutdownIO() Implements HomeSeerAPI.IPlugInAPI.ShutdownIO

    Try
      '
      ' Write the debug message
      '
      WriteMessage("Entered ShutdownIO() subroutine.", MessageType.Debug)

      Try
        hs.SaveEventsDevices()
      Catch pEx As Exception
        WriteMessage("Could not save devices", MessageType.Error)
      End Try

      If DiscoveryThread.IsAlive Then
        DiscoveryThread.Abort()
      End If

      If SendDiscoveryThread.IsAlive Then
        SendDiscoveryThread.Abort()
      End If

      If SnapshotFileMaintenanceThread.IsAlive Then
        SnapshotFileMaintenanceThread.Abort()
      End If

      If RefreshSnapshotThread.IsAlive Then
        RefreshSnapshotThread.Abort()
      End If

      If instance = "" Then
        bShutDown = True
      End If

    Catch pEx As Exception
      '
      ' Process the program exception
      '
      ProcessError(pEx, "ShutdownIO")
    Finally
      gHSInitialized = False
    End Try

  End Sub

  ''' <summary>
  ''' There may be times when you need to offer a custom function that is not part of the plugin API. 
  ''' The following API functions allow users to call your plugin from scripts and web pages by calling the functions by name.
  ''' </summary>
  ''' <param name="proc"></param>
  ''' <param name="parms"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function PluginFunction(ByVal proc As String, ByVal parms() As Object) As Object Implements IPlugInAPI.PluginFunction

    Try

      Dim ty As Type = Me.GetType
      Dim mi As MethodInfo = ty.GetMethod(proc)
      If mi Is Nothing Then
        WriteMessage(String.Format("Method {0} does not exist in {1}.", proc, IFACE_NAME), MessageType.Error)
        Return Nothing
      End If
      Return (mi.Invoke(Me, parms))

    Catch pEx As Exception
      '
      ' Process the program exception
      '
      ProcessError(pEx, "PluginFunction")
    End Try

    Return Nothing

  End Function

  ''' <summary>
  ''' There may be times when you need to offer a custom function that is not part of the plugin API. 
  ''' The following API functions allow users to call your plugin from scripts and web pages by calling the functions by name.
  ''' </summary>
  ''' <param name="proc"></param>
  ''' <param name="parms"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function PluginPropertyGet(ByVal proc As String, parms() As Object) As Object Implements IPlugInAPI.PluginPropertyGet

    Try
      Dim ty As Type = Me.GetType
      Dim mi As PropertyInfo = ty.GetProperty(proc)
      If mi Is Nothing Then
        WriteMessage(String.Format("Property {0} does not exist in {1}.", proc, IFACE_NAME), MessageType.Error)
        Return Nothing
      End If
      Return mi.GetValue(Me, parms)
    Catch pEx As Exception
      '
      ' Process the program exception
      '
      ProcessError(pEx, "PluginPropertyGet")
    End Try

    Return Nothing

  End Function

  ''' <summary>
  ''' There may be times when you need to offer a custom function that is not part of the plugin API. 
  ''' The following API functions allow users to call your plugin from scripts and web pages by calling the functions by name.
  ''' </summary>
  ''' <param name="proc"></param>
  ''' <param name="value"></param>
  ''' <remarks></remarks>
  Public Sub PluginPropertySet(ByVal proc As String, value As Object) Implements IPlugInAPI.PluginPropertySet

    Try

      Dim ty As Type = Me.GetType
      Dim mi As PropertyInfo = ty.GetProperty(proc)
      If mi Is Nothing Then
        WriteMessage(String.Format("Property {0} does not exist in {1}.", proc, IFACE_NAME), MessageType.Error)
      End If
      mi.SetValue(Me, value, Nothing)

    Catch pEx As Exception
      '
      ' Process the program exception
      '
      ProcessError(pEx, "PluginPropertySet")
    End Try

  End Sub

  ''' <summary>
  ''' This procedure will be called in your plug-in by HomeSeer whenever the user uses the search function of HomeSeer, and your plug-in is loaded and initialized.  
  ''' Unlike ActionReferencesDevice and TriggerReferencesDevice, this search is not being specific to a device, it is meant to find a match anywhere in the resources managed by your plug-in.  
  ''' This could include any textual field or object name that is utilized by the plug-in.
  ''' </summary>
  ''' <param name="SearchString"></param>
  ''' <param name="RegEx"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function Search(ByVal SearchString As String,
                         ByVal RegEx As Boolean) As HomeSeerAPI.SearchReturn() Implements IPlugInAPI.Search

  End Function

#End Region

#Region "HSPI - Devices"

  ''' <summary>
  ''' Return TRUE if your plug-in allows for configuration of your devices via the device utility page. 
  ''' This will allow you to generate some HTML controls that will be displayed to the user for modifying the device
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function SupportsConfigDevice() As Boolean Implements HomeSeerAPI.IPlugInAPI.SupportsConfigDevice
    Return True
  End Function

  ''' <summary>
  ''' If your plug-in manages all devices in the system, you can return TRUE from this function. Your configuration page will be available for all devices.
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function SupportsConfigDeviceAll() As Boolean Implements HomeSeerAPI.IPlugInAPI.SupportsConfigDeviceAll
    Return False
  End Function

  ''' <summary>
  ''' If SupportsConfigDevice returns TRUE, this function will be called when the device properties are displayed for your device. 
  ''' The device properties is displayed from the Device Utility page. This page displays a tab for each plug-in that controls the device. 
  ''' Normally, only one plug-in will be associated with a single device. 
  ''' If there is any configuration that needs to be set on the device, you can return any HTML that you would like displayed. 
  ''' Normally this would be any jquery controls that allow customization of the device. The returned HTML is just an HTML fragment and not a complete page.
  ''' </summary>
  ''' <param name="ref"></param>
  ''' <param name="user"></param>
  ''' <param name="userRights"></param>
  ''' <param name="newDevice"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ConfigDevice(ref As Integer, user As String, userRights As Integer, newDevice As Boolean) As String Implements HomeSeerAPI.IPlugInAPI.ConfigDevice

    Dim stb As New StringBuilder
    Dim jqButton As New clsJQuery.jqButton("btnSave", "Save", "DeviceUtility", True)

    Try
      '
      ' Write the debug message
      '
      WriteMessage("Entered ConfigDevice() function.", MessageType.Debug)

      Dim dv As Scheduler.Classes.DeviceClass = hs.GetDeviceByRef(ref)
      If dv Is Nothing Then Return "Error"

      Dim dv_type As String = dv.Device_Type_String(Nothing)
      Dim dv_addr As String = dv.Address(Nothing)

      If NetCams.ContainsKey(dv_addr) = False Then Return "NetCam Device not found."

      Dim NetCamDevice As NetCamDevice = NetCams(dv_addr)

      stb.AppendLine("<form id='frmConfigDevice' name='SampleTab' method='Post'>")
      stb.AppendLine("<table cellspacing='0' width='100%'>")
      stb.AppendLine(" <tr>")
      stb.AppendFormat("  <th class='tableheader' style='text-align:left;' colspan='2'>{0}</th>", "NetCam Properties")
      stb.AppendLine(" </tr>")
      stb.AppendLine(" <tr>")
      stb.AppendFormat("  <td class='tablecell' style='width:150px'>{0}</td>", "Device Name:")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>", NetCamDevice.Name)
      stb.AppendLine(" </tr>")
      stb.AppendLine(" <tr>")
      stb.AppendFormat("  <td class='tablecell' style='width:150px'>{0}</td>", "IP Address:")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>", NetCamDevice.Address)
      stb.AppendLine(" </tr>")
      stb.AppendLine(" <tr>")
      stb.AppendFormat("  <td class='tablecell' style='width:150px'>{0}</td>", "Port Number:")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>", NetCamDevice.Port)
      stb.AppendLine(" </tr>")
      stb.AppendLine(" <tr>")

      'Dim slid As New clsJQuery.jqSlider("sliderid", 0, 100, 25, clsJQuery.jqSlider.jqSliderOrientation.horizontal, 90, "DeviceUtility", True)

      'stb.AppendLine(" <tr>")
      'stb.AppendFormat("  <td class='tablecell'>{0}</td>", "Sensor Color:")
      'stb.AppendFormat("  <td class='tablecell'>{0}</td>", slid.build)
      'stb.AppendLine(" </tr>")

      stb.AppendLine("</table>")
      'stb.AppendFormat("<p>{0}</p>", slid.build)
      stb.AppendLine("</form>")

      'stb.AppendLine("<canvas id='test_canvas' width='640px' height='480px' style='border:1px solid #d3d3d3'/>")
      'stb.AppendLine("<script>")
      'stb.AppendLine("var ctx = document.getElementById('test_canvas').getContext('2d');")
      'stb.AppendLine("var img = new Image();")
      'stb.AppendLine("window.setInterval('refreshCanvas()', 1000);")
      'stb.AppendLine("function refreshCanvas(){")
      'stb.AppendLine("img.src='http://192.168.2.12/images/hspi_ultranetcam3/snapshots/NetCam001/last_snapshot.jpg?key=' + Math.random();")
      'stb.AppendLine("ctx.drawImage(img, 0, 0);")
      'stb.AppendLine("};")
      'stb.AppendLine("</script>")

    Catch pEx As Exception

    End Try

    'Dim jqButton As New clsJQuery.jqButton("button", "Press", "deviceutility", True)
    'stb.Append(jqButton.Build)

    stb.Append(clsPageBuilder.DivStart("div_config_device", ""))
    stb.Append(clsPageBuilder.DivEnd)

    Return stb.ToString

  End Function

  ''' <summary>
  ''' This function is called when a user posts information from your plugin tab on the device utility page.
  ''' </summary>
  ''' <param name="ref"></param>
  ''' <param name="data"></param>
  ''' <param name="user"></param>
  ''' <param name="userRights"></param>
  ''' <returns>
  '''  DoneAndSave = 1            Any changes to the config are saved and the page is closed and the user it returned to the device utility page
  '''  DoneAndCancel = 2          Changes are not saved and the user is returned to the device utility page
  '''  DoneAndCancelAndStay = 3   No action is taken, the user remains on the plugin tab
  '''  CallbackOnce = 4           Your plugin ConfigDevice is called so your tab can be refereshed, the user stays on the plugin tab
  '''  CallbackTimer = 5          Your plugin ConfigDevice is called and a page timer is called so ConfigDevicePost is called back every 2 seconds
  ''' </returns>
  ''' <remarks></remarks>
  Function ConfigDevicePost(ByVal ref As Integer, ByVal data As String, ByVal user As String, ByVal userRights As Integer) As Enums.ConfigDevicePostReturn Implements IPlugInAPI.ConfigDevicePost

    Try

      '
      ' Write the debug message
      '
      WriteMessage("Entered ConfigDevicePost() function.", MessageType.Debug)

      '
      ' Check if device exists and get a reference to it
      '
      Dim dv As Scheduler.Classes.DeviceClass = hs.GetDeviceByRef(ref)
      If dv Is Nothing Then Return Enums.ConfigDevicePostReturn.DoneAndCancelAndStay

    Catch pEx As Exception

    End Try

    Return Enums.ConfigDevicePostReturn.DoneAndCancelAndStay
  End Function

  ''' <summary>
  ''' Return TRUE if the plugin supports the ability to add devices through the Add Device link on the device utility page. 
  ''' If TRUE a tab appears on the add device page that allows the user to configure specific options for the new device.
  ''' When ConfigDevice is called the newDevice parameter will be True if this is the first time the device config screen is being displayed and a new device is being created. 
  ''' See ConfigDevicePost  for more information.
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks>
  ''' </remarks>
  Public Function SupportsAddDevice() As Boolean Implements HomeSeerAPI.IPlugInAPI.SupportsAddDevice
    Return False
  End Function

  ''' <summary>
  ''' If a device is owned by your plug-in (interface property set to the name of the plug-in) and the device's status_support property is set to True, 
  ''' then this procedure will be called in your plug-in when the device's status is being polled, such as when the user clicks "Poll Devices" on the device status page.
  ''' Normally your plugin will automatically keep the status of its devices updated. 
  ''' There may be situations where automatically updating devices is not possible or CPU intensive. 
  ''' In these cases the plug-in may not keep the devices updated. HomeSeer may then call this function to force an update of a specific device. 
  ''' This request is normally done when a user displays the status page, or a script runs and needs to be sure it has the latest status.
  ''' </summary>
  ''' <param name="dvref"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function PollDevice(ByVal dvref As Integer) As IPlugInAPI.PollResultInfo Implements HomeSeerAPI.IPlugInAPI.PollDevice

    '
    ' Write the debug message
    '
    WriteMessage("Entered PollDevice() function.", MessageType.Debug)

  End Function

#End Region

#Region "HSPI - Triggers"

  ''' <summary>
  ''' Return True if the given trigger can also be used as a condition, for the given trigger number.
  ''' </summary>
  ''' <param name="TriggerNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property HasConditions(ByVal TriggerNumber As Integer) As Boolean Implements HomeSeerAPI.IPlugInAPI.HasConditions
    Get
      Return hspi_plugin.HasConditions(TriggerNumber)
    End Get
  End Property

  ''' <summary>
  ''' Return True if your plugin contains any triggers, else return false  ''' 
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property HasTriggers() As Boolean Implements HomeSeerAPI.IPlugInAPI.HasTriggers
    Get
      Return hspi_plugin.HasTriggers
    End Get
  End Property

  ''' <summary>
  ''' Not documented in API
  ''' </summary>
  ''' <param name="TrigInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerTrue(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As Boolean Implements HomeSeerAPI.IPlugInAPI.TriggerTrue
    Return hspi_plugin.TriggerTrue(TrigInfo)
  End Function

  ''' <summary>
  ''' Return the number of triggers that your plugin supports.
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property TriggerCount As Integer Implements HomeSeerAPI.IPlugInAPI.TriggerCount
    Get
      Return hspi_plugin.TriggerCount
    End Get
  End Property

  ''' <summary>
  ''' Return the number of sub triggers your plugin supports.
  ''' </summary>
  ''' <param name="TriggerNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property SubTriggerCount(ByVal TriggerNumber As Integer) As Integer Implements HomeSeerAPI.IPlugInAPI.SubTriggerCount
    Get
      Return hspi_plugin.SubTriggerCount(TriggerNumber)
    End Get
  End Property

  ''' <summary>
  ''' Return the name of the given trigger based on the trigger number passed.
  ''' </summary>
  ''' <param name="TriggerNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property TriggerName(ByVal TriggerNumber As Integer) As String Implements HomeSeerAPI.IPlugInAPI.TriggerName
    Get
      Return hspi_plugin.TriggerName(TriggerNumber)
    End Get
  End Property

  ''' <summary>
  ''' Return the text name of the sub trigger given its trigger number and sub trigger number
  ''' </summary>
  ''' <param name="TriggerNumber"></param>
  ''' <param name="SubTriggerNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property SubTriggerName(ByVal TriggerNumber As Integer, ByVal SubTriggerNumber As Integer) As String Implements HomeSeerAPI.IPlugInAPI.SubTriggerName
    Get
      Return hspi_plugin.SubTriggerName(TriggerNumber, SubTriggerNumber)
    End Get
  End Property

  ''' <summary>
  ''' Given a strTrigActInfo object detect if this this trigger is configured properly, if so, return True, else False.
  ''' </summary>
  ''' <param name="TrigInfo"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property TriggerConfigured(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As Boolean Implements HomeSeerAPI.IPlugInAPI.TriggerConfigured
    Get
      Return hspi_plugin.TriggerConfigured(TrigInfo)
    End Get
  End Property

  ''' <summary>
  ''' Return HTML controls for a given trigger.
  ''' </summary>
  ''' <param name="sUnique"></param>
  ''' <param name="TrigInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerBuildUI(ByVal sUnique As String, ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As String Implements HomeSeerAPI.IPlugInAPI.TriggerBuildUI
    Return hspi_plugin.TriggerBuildUI(sUnique, TrigInfo)
  End Function

  ''' <summary>
  ''' Process a post from the events web page when a user modifies any of the controls related to a plugin trigger. 
  ''' After processing the user selctions, create and return a strMultiReturn object. 
  ''' </summary>
  ''' <param name="PostData"></param>
  ''' <param name="TrigInfoIn"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerProcessPostUI(ByVal PostData As System.Collections.Specialized.NameValueCollection, _
                                          ByVal TrigInfoIn As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As HomeSeerAPI.IPlugInAPI.strMultiReturn Implements HomeSeerAPI.IPlugInAPI.TriggerProcessPostUI
    Return hspi_plugin.TriggerProcessPostUI(PostData, TrigInfoIn)
  End Function

  ''' <summary>
  ''' After the trigger has been configured, this function is called in your plugin to display the configured trigger. 
  ''' Return text that describes the given trigger.
  ''' </summary>
  ''' <param name="TrigInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerFormatUI(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As String Implements HomeSeerAPI.IPlugInAPI.TriggerFormatUI
    Return hspi_plugin.TriggerFormatUI(TrigInfo)
  End Function

  ''' <summary>
  ''' HomeSeer will set this to TRUE if the trigger is being used as a CONDITION.  
  ''' Check this value in BuildUI and other procedures to change how the trigger is rendered if it is being used as a condition or a trigger.
  ''' Indicates (when True) that the Trigger is in Condition mode - it is for triggers that can also operate as a condition
  '''  or for allowing Conditions to appear when a condition is being added to an event.
  ''' </summary>
  ''' <param name="TrigInfo"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Property Condition(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As Boolean Implements HomeSeerAPI.IPlugInAPI.Condition
    Set(ByVal value As Boolean)
      hspi_plugin.Condition(TrigInfo) = value
    End Set
    Get
      Return hspi_plugin.Condition(TrigInfo)
    End Get
  End Property

  ''' <summary>
  ''' Return True if the given device is referenced by the given trigger.
  ''' </summary>
  ''' <param name="TrigInfo"></param>
  ''' <param name="dvRef"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerReferencesDevice(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo, _
                                                                           ByVal dvRef As Integer) As Boolean _
                                                                           Implements HomeSeerAPI.IPlugInAPI.TriggerReferencesDevice
    Return False
  End Function

#End Region

#Region "HSPI - Actions"

  ''' <summary>
  ''' Return the number of actions the plugin supports.
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ActionCount() As Integer Implements HomeSeerAPI.IPlugInAPI.ActionCount
    Return hspi_plugin.ActionCount
  End Function

  ''' <summary>
  ''' When an event is triggered, this function is called to carry out the selected action. Use the ActInfo parameter to determine what action needs to be executed then execute this action.
  ''' Return TRUE if the action was executed successfully, else FALSE if there was an error.
  ''' </summary>
  ''' <param name="ActInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function HandleAction(ByVal ActInfo As IPlugInAPI.strTrigActInfo) As Boolean Implements HomeSeerAPI.IPlugInAPI.HandleAction
    Return hspi_plugin.HandleAction(ActInfo)
  End Function

  ''' <summary>
  ''' Missing in the API Docs
  ''' </summary>
  ''' <param name="ActInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ActionFormatUI(ByVal ActInfo As IPlugInAPI.strTrigActInfo) As String Implements HomeSeerAPI.IPlugInAPI.ActionFormatUI
    Return hspi_plugin.ActionFormatUI(ActInfo)
  End Function

  Public Function ActionProcessPostUI(ByVal PostData As Collections.Specialized.NameValueCollection, _
                                      ByVal TrigInfoIN As IPlugInAPI.strTrigActInfo) As IPlugInAPI.strMultiReturn Implements HomeSeerAPI.IPlugInAPI.ActionProcessPostUI

    Return hspi_plugin.ActionProcessPostUI(PostData, TrigInfoIN)

  End Function

  ''' <summary>
  ''' When a user edits your event actions in the HomeSeer events, this function is called to process the selections.
  ''' </summary>
  ''' <param name="ActInfo"></param>
  ''' <param name="dvRef"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ActionReferencesDevice(ByVal ActInfo As IPlugInAPI.strTrigActInfo, _
                                         ByVal dvRef As Integer) As Boolean _
                                         Implements HomeSeerAPI.IPlugInAPI.ActionReferencesDevice

  End Function

  ''' <summary>
  ''' This function is called from the HomeSeer event page when an event is in edit mode. Your plug-in needs to return HTML controls so the user can make action selections. 
  ''' Normally this is one of the HomeSeer jquery controls such as a clsJquery.jqueryCheckbox.
  ''' </summary>
  ''' <param name="sUnique"></param>
  ''' <param name="ActInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ActionBuildUI(ByVal sUnique As String, ByVal ActInfo As IPlugInAPI.strTrigActInfo) As String Implements HomeSeerAPI.IPlugInAPI.ActionBuildUI
    Return hspi_plugin.ActionBuildUI(sUnique, ActInfo)
  End Function

  ''' <summary>
  ''' Return TRUE if the given action is configured properly. 
  ''' There may be times when a user can select invalid selections for the action and in this case you would return FALSE so HomeSeer will not allow the action to be saved.
  ''' </summary>
  ''' <param name="ActInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ActionConfigured(ByVal ActInfo As IPlugInAPI.strTrigActInfo) As Boolean Implements HomeSeerAPI.IPlugInAPI.ActionConfigured
    Return hspi_plugin.ActionConfigured(ActInfo)
  End Function

  ''' <summary>
  ''' Return the name of the action given an action number. The name of the action will be displayed in the HomeSeer events actions list.
  ''' </summary>
  ''' <param name="ActionNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property ActionName(ByVal ActionNumber As Integer) As String Implements HomeSeerAPI.IPlugInAPI.ActionName
    Get
      Return hspi_plugin.ActionName(ActionNumber)
    End Get
  End Property

  ''' <summary>
  ''' The HomeSeer events page has an option to set the editing mode to "Advanced Mode". 
  ''' This is typically used to enable options that may only be of interest to advanced users or programmers. The Set in this function is called when advanced mode is enabled. 
  ''' Your plug-in can also enable this mode if an advanced selection was saved and needs to be displayed.
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Property ActionAdvancedMode As Boolean Implements HomeSeerAPI.IPlugInAPI.ActionAdvancedMode
    Set(ByVal value As Boolean)
      mvarActionAdvanced = value
    End Set
    Get
      Return mvarActionAdvanced
    End Get
  End Property
  Private mvarActionAdvanced As Boolean

#End Region

#Region "HSPI - WebPage"

  ''' <summary>
  ''' When your plug-in web page has form elements on it, and the form is submitted, this procedure is called to handle the HTTP "Put" request.  
  ''' There must be one PagePut procedure in each plug-in object or class that is registered as a web page in HomeSeer.
  ''' </summary>
  ''' <param name="pageName"></param>
  ''' <param name="user"></param>
  ''' <param name="userRights"></param>
  ''' <param name="queryString"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetPagePlugin(ByVal pageName As String, ByVal user As String, ByVal userRights As Integer, ByVal queryString As String) As String Implements HomeSeerAPI.IPlugInAPI.GetPagePlugin

    Try
      '
      ' Build and return the actual page
      '
      WebPage = SelectPage(pageName)
      Return WebPage.GetPagePlugin(pageName, user, userRights, queryString, instance)

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "GetPagePlugin")
      Return ""
    End Try

  End Function

  ''' <summary>
  ''' Determine what page we need to display
  ''' </summary>
  ''' <param name="pageName"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function SelectPage(ByVal pageName As String) As Object

    WriteMessage("Entered SelectPage() Function", MessageType.Debug)
    WriteMessage("pageName=" & pageName, MessageType.Debug)

    SelectPage = Nothing
    Try

      Select Case pageName
        Case WebConfigPage.PageName
          SelectPage = WebConfigPage
        Case Else
          SelectPage = WebConfigPage
      End Select

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "SelectPage")
    End Try

  End Function

  ''' <summary>
  ''' When a user clicks on any controls on one of your web pages, this function is then called with the post data. 
  ''' You can then parse the data and process as needed.
  ''' </summary>
  ''' <param name="pageName"></param>
  ''' <param name="data"></param>
  ''' <param name="user"></param>
  ''' <param name="userRights"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function PostBackProc(ByVal pageName As String, ByVal data As String, ByVal user As String, ByVal userRights As Integer) As String Implements HomeSeerAPI.IPlugInAPI.PostBackProc
    Try
      '
      ' Build and return the actual page
      '
      WebPage = SelectPage(pageName)
      Return WebPage.postBackProc(pageName, data, user, userRights)

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "PostBackProc")
    End Try
  End Function

  ''' <summary>
  ''' This function is called by HomeSeer from the form or class object that a web page was registered with using RegisterConfigLink.  
  ''' You must have a GenPage procedure per web page that you register with HomeSeer.  
  ''' This page is called when the user requests the web page with an HTTP Get command, which is the default operation when the browser requests a page.
  ''' </summary>
  ''' <param name="link"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GenPage(ByVal link As String) As String Implements HomeSeerAPI.IPlugInAPI.GenPage

    Dim sb As New StringBuilder()

    Try
      '
      ' Generate the HTML re-direct
      '
      sb.Append("HTTP/1.1 301 Moved Permanently" & vbCrLf)
      sb.AppendFormat("Location: {0}" & vbCrLf, LINK_TARGET)
      sb.Append(vbCrLf)

      Return sb.ToString

    Catch pEx As Exception
      '
      ' Process the error
      '
      Return ""
    End Try

  End Function

  ''' <summary>
  ''' When your plug-in web page has form elements on it, and the form is submitted, this procedure is called to handle the HTTP "Put" request.  
  ''' There must be one PagePut procedure in each plug-in object or class that is registered as a web page in HomeSeer
  ''' </summary>
  ''' <param name="data"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function PagePut(ByVal data As String) As String Implements HomeSeerAPI.IPlugInAPI.PagePut
    Return ""
  End Function

#End Region

#Region "HSPI - Speak Proxy"

  ''' <summary>
  ''' If your plug-in is registered as a Speak proxy plug-in, then when HomeSeer is asked to speak something, it will pass the speak information to your plug-in using this procedure.  
  ''' When your plug-in is ready to do the actual speaking, it should call SpeakProxy, and pass the information that it got from this procedure to SpeakProxy.  
  ''' It may be necessary or a feature of your plug-in to modify the text being spoken or the host/instance list provided in the host parameter - this is acceptable.
  ''' </summary>
  ''' <param name="device"></param>
  ''' <param name="txt"></param>
  ''' <param name="w"></param>
  ''' <param name="host"></param>
  ''' <remarks></remarks>
  Public Sub SpeakIn(device As Integer, txt As String, w As Boolean, host As String) Implements HomeSeerAPI.IPlugInAPI.SpeakIn

  End Sub

#End Region

#Region "HSPI - Callbacks"

  Public Sub HSEvent(ByVal EventType As Enums.HSEvent, ByVal parms() As Object) Implements HomeSeerAPI.IPlugInAPI.HSEvent

    Console.WriteLine("HSEvent: " & EventType.ToString)
    Select Case EventType
      Case Enums.HSEvent.VALUE_CHANGE
        '
        ' Do stuff
        '
    End Select

  End Sub

  ''' <summary>
  ''' Return True if the given device is referenced in the given action.
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function RaisesGenericCallbacks() As Boolean Implements HomeSeerAPI.IPlugInAPI.RaisesGenericCallbacks
    Return False
  End Function

#End Region

#Region "HSPI - Plug-in Version Information"

  ''' <summary>
  ''' Returns the full version number of the HomeSeer plug-in
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Shared Function Version() As String

    Dim strVersion As String = ""

    Try
      '
      ' Write the debug message
      '
      WriteMessage("Entered Version() function.", MessageType.Debug)

      '
      ' Build the plug-in's version
      '
      strVersion = GetVersionInfo(GetExecutingAssembly.Location).ProductMajorPart() & "." _
                  & GetVersionInfo(GetExecutingAssembly.Location).ProductMinorPart() & "." _
                  & GetVersionInfo(GetExecutingAssembly.Location).ProductBuildPart() & "." _
                  & GetVersionInfo(GetExecutingAssembly.Location).ProductPrivatePart()

    Catch pEx As Exception
      '
      ' Process error condtion
      '
      strVersion = "??.??.??"
    End Try

    Return strVersion

  End Function

#End Region

#Region "HSPI - Public Sub/Functions"

#Region "HSPI - NetCam Controls"

  ''' <summary>
  ''' NetCam Control URL
  ''' </summary>
  ''' <param name="control_id"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetNetCamControlURL(ByVal control_id As Integer) As String

    Return hspi_plugin.GetNetCamControlURL(control_id)

  End Function

  ''' <summary>
  ''' Gets the NetCam Types from the underlying database
  ''' </summary>
  ''' <param name="netcam_type"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetNetCamControlsFromDB(ByVal netcam_type As Integer) As DataTable

    Return hspi_plugin.GetNetCamControlsFromDB(netcam_type)

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
                                       ByVal bRefreshNetCamList As Boolean) As Boolean

    Return hspi_plugin.InsertNetCamControls(netcam_type, control_name, control_url, bRefreshNetCamList)

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
                                       ByVal bRefreshNetCamList As Boolean) As Boolean

    Return hspi_plugin.UpdateNetCamControls(control_id, control_name, control_url, bRefreshNetCamList)

  End Function

  ''' <summary>
  ''' Removes existing NetCam Controls stored in the database
  ''' </summary>
  ''' <param name="control_id"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function DeleteNetCamControls(ByVal control_id As Integer) As Boolean

    Return hspi_plugin.DeleteNetCamControls(control_id)

  End Function

  ''' <summary>
  ''' Processes Upload File
  ''' </summary>
  ''' <param name="netcam_type"></param>
  ''' <param name="MyFileStream"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ImportNetCamControls(ByVal netcam_type As Integer, ByRef MyFileStream As System.IO.Stream) As Boolean

    Return hspi_plugin.ImportNetCamControls(netcam_type, MyFileStream)

  End Function

  ''' <summary>
  ''' Truncates NetCam Controls
  ''' </summary>
  ''' <param name="netcam_type"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TruncateNetCamControls(ByVal netcam_type As Integer) As Boolean

    Return hspi_plugin.TruncateNetCamControls(netcam_type)

  End Function

#End Region

#Region "HSPI - Snapshots"

  ''' <summary>
  ''' Takes a single snapshot
  ''' </summary>
  ''' <param name="netcam_id"></param>
  ''' <remarks></remarks>
  Public Sub NetCamSnapshot(ByVal netcam_id As String)

    hspi_plugin.NetCamSnapshot(netcam_id)

  End Sub

  ''' <summary>
  ''' Get the camera picture filenames
  ''' </summary>
  ''' <param name="strNetCamName"></param>
  ''' <param name="strEventId"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetCameraSnapshotInfo(ByVal strNetCamName As String, Optional ByVal strEventId As String = "") As DataTable

    Return hspi_plugin.GetCameraSnapshotInfo(strNetCamName, strEventId)

  End Function

  ''' <summary>
  ''' Purges snapshots from filesystem
  ''' </summary>
  ''' <param name="strNetCamName"></param>
  ''' <param name="strEventId"></param>
  ''' <remarks></remarks>
  Public Sub PurgeSnapshotEvent(ByVal strNetCamName As String, ByVal strEventId As String)

    hspi_plugin.PurgeSnapshotEvent(strNetCamName, strEventId)

  End Sub

  ''' <summary>
  ''' Refreshes the latest snapshots
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub RefreshLatestSnapshots()

    hspi_plugin.RefreshLatestSnapshots()

  End Sub

  ''' <summary>
  ''' Get the camera snapshot minimum width
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetCameraSnapshotMinWidth() As Integer

    Return hspi_plugin.GetCameraSnapshotMinWidth()

  End Function

  ''' <summary>
  ''' Get the camera viewer filenames
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetCameraSnapshotViewer() As DataTable

    Return hspi_plugin.GetCameraSnapshotViewer()

  End Function

#End Region

#Region "HSPI - NetCam Events"

  ''' <summary>
  ''' GetNetCamEventSummary
  ''' </summary>
  ''' <param name="strNetCamName"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetNetCamEventSummary(ByVal strNetCamName As String) As SortedList

    Return hspi_plugin.GetNetCamEventSummary(strNetCamName)

  End Function

#End Region

#Region "HSPI - NetCam Types"

  ''' <summary>
  ''' Gets the NetCam Types from the underlying database
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetNetCamTypesFromDB() As DataTable

    Return hspi_plugin.GetNetCamTypesFromDB()

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
                                   ByVal videostream_path As String) As Boolean

    Return hspi_plugin.InsertNetCamType(netcam_vendor, netcam_model, snapshot_path, videostream_path)

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

    Return hspi_plugin.UpdateNetCamType(netcam_type, netcam_vendor, netcam_model, snapshot_path, videostream_path)

  End Function

  ''' <summary>
  ''' Removes existing NetCam Type stored in the database
  ''' </summary>
  ''' <param name="netcam_type"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function DeleteNetCamType(ByVal netcam_type As Integer) As Boolean

    Return hspi_plugin.DeleteNetCamType(netcam_type)

  End Function

#End Region

#Region "HSPI - NetCam Devices"

  ''' <summary>
  ''' Refreshes the NetCam List
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub RefreshNetCamList()

    hspi_plugin.RefreshNetCamList()

  End Sub

  ''' <summary>
  ''' Gets the NetCam Devices from the underlying database
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetNetCamDevicesFromDB() As DataTable

    Return hspi_plugin.GetNetCamDevicesFromDB()

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
                                     ByVal auth_pass As String) As Boolean

    Return hspi_plugin.InsertNetCamDevice(netcam_name, netcam_address, netcam_port, netcam_type, auth_user, auth_pass)

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

    Return hspi_plugin.UpdateNetCamDevice(netcam_id, netcam_name, netcam_address, netcam_port, netcam_type, auth_user, auth_pass)

  End Function

  ''' <summary>
  ''' Removes existing NetCam Profile stored in the database
  ''' </summary>
  ''' <param name="netcam_id"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function DeleteNetCamDevice(ByVal netcam_id As Integer) As Boolean

    Return hspi_plugin.DeleteNetCamDevice(netcam_id)

  End Function

#End Region

#Region "HSPI - Misc"

  ''' <summary>
  ''' Alarm Trigger
  ''' </summary>
  ''' <param name="strRemoteIP"></param>
  ''' <remarks></remarks>
  Public Sub AlarmTrigger(ByVal strRemoteIP As String)

    hspi_plugin.AlarmTrigger(strRemoteIP)

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

    Return hspi_plugin.ExecuteSQL(strSQL, iRecordCount, iPageSize, iPageCount, iPageCur)

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

    Return hspi_plugin.GetSetting(strSection, strKey, strValueDefault)

  End Function

  ''' <summary>
  '''  Saves plug-in settings to INI file
  ''' </summary>
  ''' <param name="strSection"></param>
  ''' <param name="strKey"></param>
  ''' <param name="strValue"></param>
  ''' <remarks></remarks>
  Public Sub SaveSetting(ByVal strSection As String, _
                         ByVal strKey As String, _
                         ByVal strValue As String)

    hspi_plugin.SaveSetting(strSection, strKey, strValue)

  End Sub

#End Region

#End Region

#Region "HSPI - Public Properties"

  ''' <summary>
  ''' Returns the FFmpeg Install Status
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetFFmpegStatus() As String

    Return hspi_plugin.GetFFmpegStatus()

  End Function

  ''' <summary>
  ''' Returns number of NetCams
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetNetCamCount() As Integer

    Return hspi_plugin.GetNetCamCount()

  End Function


  ''' <summary>
  ''' Returns number of snapshots in directory
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetSnapshotCount() As Integer

    Return hspi_plugin.GetSnapshotCount()

  End Function

  ''' <summary>
  ''' Property to control log level
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Property LoggingLevel() As Integer
    Get
      Return gLogLevel
    End Get
    Set(ByVal value As Integer)
      gLogLevel = value
    End Set
  End Property

#End Region

#Region "HSPI - Web Authorization"

  ''' <summary>
  ''' Returns the list of users authorized to access the web page
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function WEBUserRolesAuthorized() As Integer

    Return hspi_main.WEBUserRolesAuthorized

  End Function

  '----------------------------------------------------------------------
  'Purpose: Determine if logged in user is authorized to view the web page
  'Inputs:  LoggedInUser as String
  'Outputs: Boolean (True indicates user is authorized)
  '----------------------------------------------------------------------
  ''' <summary>
  ''' Determine if logged in user is authorized to view the web page
  ''' </summary>
  ''' <param name="LoggedInUser"></param>
  ''' <param name="USER_ROLES_AUTHORIZED"></param>
  ''' <returns>Boolean (True indicates user is authorized)</returns>
  ''' <remarks></remarks>
  Public Function WEBUserIsAuthorized(ByVal LoggedInUser As String, _
                                      ByVal USER_ROLES_AUTHORIZED As Integer) As Boolean

    Return hspi_main.WEBUserIsAuthorized(LoggedInUser, USER_ROLES_AUTHORIZED)

  End Function

#End Region

End Class
