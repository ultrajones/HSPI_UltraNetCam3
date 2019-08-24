Imports HomeSeerAPI
Imports System.Data.Common

Module hspi_devices

  Dim bCreateRootDevice = True

  ''' <summary>
  ''' Create the HomeSeer Root Device
  ''' </summary>
  ''' <param name="strRootId"></param>
  ''' <param name="strRootName"></param>
  ''' <param name="dv_ref_child"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function CreateRootDevice(ByVal strRootId As String, _
                                   ByVal strRootName As String, _
                                   ByVal dv_ref_child As Integer) As Integer

    Dim dv As Scheduler.Classes.DeviceClass

    Dim dv_ref As Integer = 0
    Dim dv_misc As Integer = 0

    Dim dv_name As String = ""
    Dim dv_type As String = ""
    Dim dv_addr As String = ""

    Dim DeviceShowValues As Boolean = False

    Try
      '
      ' Set the local variables
      '
      If strRootId = "Plugin" Then
        dv_name = "UltraNetCam3 Plugin"
        dv_addr = String.Format("{0}-Root", strRootName.Replace(" ", "-"))
        dv_type = dv_name
      Else
        dv_name = strRootName
        dv_addr = String.Format("{0}-Root", strRootId, strRootName.Replace(" ", "-"))
        dv_type = strRootName
      End If

      dv = LocateDeviceByAddr(dv_addr)
      Dim bDeviceExists As Boolean = Not dv Is Nothing

      If bDeviceExists = True Then
        '
        ' Lets use the existing device
        '
        dv_addr = dv.Address(hs)
        dv_ref = dv.Ref(hs)

        Call WriteMessage(String.Format("Updating existing HomeSeer {0} root device.", dv_name), MessageType.Debug)

      Else
        '
        ' Create A HomeSeer Device
        '
        dv_ref = hs.NewDeviceRef(dv_name)
        If dv_ref > 0 Then
          dv = hs.GetDeviceByRef(dv_ref)
        End If

        Call WriteMessage(String.Format("Creating new HomeSeer {0} root device.", dv_name), MessageType.Debug)

      End If

      '
      ' Define the HomeSeer device
      '
      dv.Address(hs) = dv_addr
      dv.Interface(hs) = IFACE_NAME
      dv.InterfaceInstance(hs) = Instance

      '
      ' Update location properties on new devices only
      '
      If bDeviceExists = False Then
        dv.Location(hs) = IFACE_NAME & " Plugin"
        dv.Location2(hs) = IIf(strRootId = "Plugin", "Plug-ins", dv_type)
      End If

      '
      ' The following simply shows up in the device properties but has no other use
      '
      dv.Device_Type_String(hs) = dv_type

      '
      ' Set the DeviceTypeInfo
      '
      Dim DT As New DeviceTypeInfo
      DT.Device_API = DeviceTypeInfo.eDeviceAPI.Plug_In
      DT.Device_Type = DeviceTypeInfo.eDeviceType_Plugin.Root
      dv.DeviceType_Set(hs) = DT

      '
      ' Make this a parent root device
      '
      dv.Relationship(hs) = Enums.eRelationship.Parent_Root
      dv.AssociatedDevice_Add(hs, dv_ref_child)

      Dim image As String = "device_root.png"

      Dim VSPair As VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
      VSPair.PairType = VSVGPairType.SingleValue
      VSPair.Value = 0
      VSPair.Status = "Root"
      VSPair.Render = Enums.CAPIControlType.Values
      hs.DeviceVSP_AddPair(dv_ref, VSPair)

      Dim VGPair As VGPair = New VGPair()
      VGPair.PairType = VSVGPairType.SingleValue
      VGPair.Set_Value = 0
      VGPair.Graphic = String.Format("{0}{1}", gImageDir, image)
      hs.DeviceVGP_AddPair(dv_ref, VGPair)

      '
      ' Update the Device Misc Bits
      '
      If DeviceShowValues = True Then
        dv.MISC_Set(hs, Enums.dvMISC.SHOW_VALUES)
      End If

      If bDeviceExists = False Then
        dv.MISC_Set(hs, Enums.dvMISC.NO_LOG)
      End If

      dv.Status_Support(hs) = False

      hs.SaveEventsDevices()

    Catch pEx As Exception

    End Try

    Return dv_ref

  End Function

  ''' <summary>
  ''' Subroutine to create HomeSeer device
  ''' </summary>
  ''' <param name="NewNetCamDevice"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function CreateNetCamDevice(ByRef NewNetCamDevice As Hashtable) As String

    Dim dv As Scheduler.Classes.DeviceClass
    Dim dv_ref As Integer = 0
    Dim dv_misc As Integer = 0

    Dim dv_name As String = ""
    Dim dv_type As String = ""
    Dim dv_addr As String = ""

    Dim DevicePairs As New ArrayList

    Try

      If NewNetCamDevice.Count = 0 Then
        Throw New Exception("CreateNetCamDevice():  Unable to create NetCam HomeSeer device because the parameters are invalid.")
      End If

      '
      ' Set the local variables
      '
      Dim iNetCamTypeId As String = Int16.Parse(NewNetCamDevice("netcam_type"))
      Dim iNetCamId As String = Int16.Parse(NewNetCamDevice("netcam_id"))
      Dim strNetCamId As String = String.Format("NetCam{0}", iNetCamId.ToString.PadLeft(3, "0"))

      dv_addr = String.Format("NetCam{0}", iNetCamId.ToString.PadLeft(3, "0"))
      dv_name = NewNetCamDevice("netcam_name")
      dv_type = "NetCam"

      dv = LocateDeviceByAddr(dv_addr)
      Dim bDeviceExists As Boolean = Not dv Is Nothing

      If bDeviceExists = True Then
        '
        ' Lets use the existing device
        '
        dv_addr = dv.Address(hs)
        dv_ref = dv.Ref(hs)

        Call WriteMessage(String.Format("Updating existing HomeSeer {0} device.", dv_name), MessageType.Debug)

      Else
        '
        ' Create A HomeSeer Device
        '
        dv_ref = hs.NewDeviceRef(dv_name)
        If dv_ref > 0 Then
          dv = hs.GetDeviceByRef(dv_ref)
        End If
        dv_addr = strNetCamId

        Call WriteMessage(String.Format("Creating new HomeSeer {0} device.", dv_name), MessageType.Debug)

      End If

      '
      ' Make this device a child of the root
      '
      If dv.Relationship(hs) <> Enums.eRelationship.Child Then

        If bCreateRootDevice = True Then
          dv.AssociatedDevice_ClearAll(hs)
          Dim dvp_ref As Integer = CreateRootDevice("Plugin", IFACE_NAME, dv_ref)
          If dvp_ref > 0 Then
            dv.AssociatedDevice_Add(hs, dvp_ref)
          End If
          dv.Relationship(hs) = Enums.eRelationship.Child
        End If

        hs.SaveEventsDevices()
      End If

      '
      ' Exit if our device exists
      ' commented out because it prevents controls from being updated
      'If bDeviceExists = True Then Return dv_addr

      '
      ' Define the HomeSeer device
      '
      dv.Address(hs) = dv_addr
      dv.Interface(hs) = IFACE_NAME
      dv.InterfaceInstance(hs) = Instance

      '
      ' Set the image
      '
      Dim strSnapshotFilename As String = String.Format("{0}/{1}/last_snapshot.jpg", "images/hspi_ultranetcam3/snapshots", dv_addr)
      Dim strThumbnailFilename As String = String.Format("{0}/{1}/last_thumbnail.jpg", "images/hspi_ultranetcam3/snapshots", dv_addr)
      dv.ImageLarge(hs) = strSnapshotFilename
      dv.Image(hs) = strThumbnailFilename

      '
      ' Update location properties on new devices only
      '
      If bDeviceExists = False Then
        dv.Location(hs) = IFACE_NAME & " Plugin"
        dv.Location2(hs) = strNetCamId
      End If

      '
      ' The following simply shows up in the device properties but has no other use
      '
      dv.Device_Type_String(hs) = dv_type

      '
      ' Set the DeviceTypeInfo
      '
      Dim DT As New DeviceTypeInfo
      DT.Device_API = DeviceTypeInfo.eDeviceAPI.Plug_In
      DT.Device_Type = DeviceTypeInfo.eDeviceType_Plugin.Root
      dv.DeviceType_Set(hs) = DT

      '
      ' Define the default image
      '
      Dim strNetCamImage As String = "netcam.png"

      DevicePairs.Clear()
      DevicePairs.Add(New hspi_device_pairs(-2, "", strNetCamImage, HomeSeerAPI.ePairStatusControl.Control))
      DevicePairs.Add(New hspi_device_pairs(-1, "Take Snapshot", strNetCamImage, HomeSeerAPI.ePairStatusControl.Control))
      DevicePairs.Add(New hspi_device_pairs(0, dv_type, strNetCamImage, HomeSeerAPI.ePairStatusControl.Status))

      Try

        Dim strSQL As String = String.Format("SELECT control_id, control_name FROM tblNetCamControls WHERE netcam_type={0}", iNetCamTypeId)

        '
        ' Execute the data reader
        '
        Dim MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

        MyDbCommand.Connection = DBConnectionMain
        MyDbCommand.CommandType = CommandType.Text
        MyDbCommand.CommandText = strSQL

        SyncLock SyncLockMain
          Dim dtrResults As IDataReader = MyDbCommand.ExecuteReader()

          '
          ' Process the resutls
          '
          While dtrResults.Read()
            Dim id As String = dtrResults("control_id")
            Dim control_name As String = dtrResults("control_name")

            DevicePairs.Add(New hspi_device_pairs(id, control_name, strNetCamImage, HomeSeerAPI.ePairStatusControl.Control))
          End While

          dtrResults.Close()
        End SyncLock

        MyDbCommand.Dispose()

      Catch pEx As Exception

      End Try

      hs.DeviceVSP_ClearAll(dv_ref, True)
      hs.DeviceVGP_ClearAll(dv_ref, True)
      hs.SaveEventsDevices()

      '
      ' Add the Status Graphic Pairs
      '
      For Each Pair As hspi_device_pairs In DevicePairs

        Dim VSPair As VSPair = New VSPair(Pair.Type)
        VSPair.PairType = VSVGPairType.SingleValue
        VSPair.Value = Pair.Value
        VSPair.Status = Pair.Status
        VSPair.Render = Enums.CAPIControlType.Values
        hs.DeviceVSP_AddPair(dv_ref, VSPair)

        Dim VGPair As VGPair = New VGPair()
        VGPair.PairType = VSVGPairType.SingleValue
        VGPair.Set_Value = Pair.Value
        VGPair.Graphic = String.Format("{0}{1}", gImageDir, Pair.Image)
        hs.DeviceVGP_AddPair(dv_ref, VGPair)

      Next

      dv.MISC_Set(hs, Enums.dvMISC.SHOW_VALUES)

      dv.MISC_Set(hs, Enums.dvMISC.NO_LOG)
      dv.Status_Support(hs) = False

      hs.SaveEventsDevices()

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "CreateNetCamDevice()")
    End Try

    Return dv_addr

  End Function

  ''' <summary>
  ''' Locates device by device code
  ''' </summary>
  ''' <param name="strDeviceAddr"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function LocateDeviceByAddr(ByVal strDeviceAddr As String) As Object

    Dim objDevice As Object
    Dim dev_ref As Long = 0

    Try

      dev_ref = hs.DeviceExistsAddress(strDeviceAddr, False)
      objDevice = hs.GetDeviceByRef(dev_ref)
      If Not objDevice Is Nothing Then
        Return objDevice
      End If

    Catch pEx As Exception
      '
      ' Process the program exception
      '
      Call ProcessError(pEx, "LocateDeviceByAddr")
    End Try
    Return Nothing ' No device found

  End Function

  ''' <summary>
  ''' Sets the HomeSeer string and device values
  ''' </summary>
  ''' <param name="dv_addr"></param>
  ''' <param name="dv_value"></param>
  ''' <remarks></remarks>
  Public Sub SetDeviceValue(ByVal dv_addr As String, _
                            ByVal dv_value As String)

    Try

      WriteMessage(String.Format("{0}->{1}", dv_addr, dv_value), MessageType.Debug)

      Dim dv_ref As Integer = hs.DeviceExistsAddress(dv_addr, False)
      Dim bDeviceExists As Boolean = dv_ref <> -1

      WriteMessage(String.Format("Device address {0} was found.", dv_addr), MessageType.Debug)

      If bDeviceExists = True Then

        If IsNumeric(dv_value) Then

          Dim dblDeviceValue As Double = Double.Parse(hs.DeviceValueEx(dv_ref))
          Dim dblSensorValue As Double = Double.Parse(dv_value)

          If dblDeviceValue <> dblSensorValue Then
            hs.SetDeviceValueByRef(dv_ref, dblSensorValue, True)
          End If

        End If

      Else
        WriteMessage(String.Format("Device address {0} cannot be found.", dv_addr), MessageType.Warning)
      End If

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "SetDeviceValue()")

    End Try

  End Sub

  ''' <summary>
  ''' Sets the HomeSeer device string
  ''' </summary>
  ''' <param name="dv_addr"></param>
  ''' <param name="dv_string"></param>
  ''' <remarks></remarks>
  Public Sub SetDeviceString(ByVal dv_addr As String, _
                             ByVal dv_string As String)

    Try

      WriteMessage(String.Format("{0}->{1}", dv_addr, dv_string), MessageType.Debug)

      Dim dv_ref As Integer = hs.DeviceExistsAddress(dv_addr, False)
      Dim bDeviceExists As Boolean = dv_ref <> -1

      WriteMessage(String.Format("Device address {0} was found.", dv_addr), MessageType.Debug)

      If bDeviceExists = True Then

        hs.SetDeviceString(dv_ref, dv_string, True)

      Else
        WriteMessage(String.Format("Device address {0} cannot be found.", dv_addr), MessageType.Warning)
      End If

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "SetDeviceString()")

    End Try

  End Sub

End Module
