Imports System.Text
Imports System.Web
Imports Scheduler
Imports System.Collections.Specialized
Imports System.Web.Script.Serialization
Imports System.Text.RegularExpressions

Public Class hspi_webconfig
  Inherits clsPageBuilder

  Public hspiref As HSPI

  Dim TimerEnabled As Boolean

  ''' <summary>
  ''' Initializes new webconfig
  ''' </summary>
  ''' <param name="pagename"></param>
  ''' <remarks></remarks>
  Public Sub New(ByVal pagename As String)
    MyBase.New(pagename)
  End Sub

#Region "Page Building"

  ''' <summary>
  ''' Web pages that use the clsPageBuilder class and registered with hs.RegisterLink and hs.RegisterConfigLink will then be called through this function. 
  ''' A complete page needs to be created and returned.
  ''' </summary>
  ''' <param name="pageName"></param>
  ''' <param name="user"></param>
  ''' <param name="userRights"></param>
  ''' <param name="queryString"></param>
  ''' <param name="instance"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetPagePlugin(ByVal pageName As String, ByVal user As String, ByVal userRights As Integer, ByVal queryString As String, instance As String) As String

    Try

      Dim stb As New StringBuilder
      '
      ' Called from the start of your page to reset all internal data structures in the clsPageBuilder class, such as menus.
      '
      Me.reset()

      '
      ' Determine if user is authorized to access the web page
      '
      Dim LoggedInUser As String = hs.WEBLoggedInUser()
      Dim USER_ROLES_AUTHORIZED As Integer = WEBUserRolesAuthorized()

      '
      ' Process Alarm Trigger
      '
      If WEBUserIsAuthorized(LoggedInUser, USER_ROLES_AUTHORIZED) = True Then
        Dim strIPAddress = hs.GetIPAddress
        Dim strRemoteIP As String = hs.GetLastRemoteIP

        Dim parts As Collections.Specialized.NameValueCollection = Nothing
        If (queryString <> "") Then
          parts = HttpUtility.ParseQueryString(queryString)
          If parts("trigger").ToLower = "alarm" Then
            hspi_plugin.AlarmTrigger(strRemoteIP)
            stb.AppendLine("OK")
          Else
            stb.AppendFormat("Error:  Please use http://{0}/UltraNetCam3?trigger=alarm for an {1} alarm trigger.", strIPAddress, IFACE_NAME)
          End If
          Return stb.ToString
        End If
      End If

      '
      ' Add jQuery and CSS to web page
      '
      Dim Header As New StringBuilder
      Header.AppendLine("<link type=""text/css"" href=""/hspi_ultranetcam3/css/jquery.dataTables.min.css"" rel=""stylesheet"" />")
      Header.AppendLine("<link type=""text/css"" href=""/hspi_ultranetcam3/css/dataTables.tableTools.css"" rel=""stylesheet"" />")
      Header.AppendLine("<link type=""text/css"" href=""/hspi_ultranetcam3/css/dataTables.editor.min.css"" rel=""stylesheet"" />")
      Header.AppendLine("<link type=""text/css"" href=""/hspi_ultranetcam3/css/jquery.dataTables_themeroller.css"" rel=""stylesheet"" />")
      Header.AppendLine("<link type=""text/css"" rel=""stylesheet"" href=""/hspi_ultranetcam3/css/lightbox.css"" />")

      Header.AppendLine("<script type=""text/javascript"" src=""/hspi_ultranetcam3/js/jquery.dataTables.min.js""></script>")
      Header.AppendLine("<script type=""text/javascript"" src=""/hspi_ultranetcam3/js/dataTables.tableTools.min.js""></script>")
      Header.AppendLine("<script type=""text/javascript"" src=""/hspi_ultranetcam3/js/dataTables.editor.min.js""></script>")
      Header.AppendLine("<script type=""text/javascript"" src=""/hspi_ultranetcam3/js/hspi_ultranetcam3_device_types.js""></script>")
      Header.AppendLine("<script type=""text/javascript"" src=""/hspi_ultranetcam3/js/hspi_ultranetcam3_controls.js""></script>")
      Header.AppendLine("<script type=""text/javascript"" src=""/hspi_ultranetcam3/js/hspi_ultranetcam3_devices.js""></script>")
      Header.AppendLine("<script type=""text/javascript"" src=""/hspi_ultranetcam3/js/lightbox.min.js""></script>")

      Me.AddHeader(Header.ToString)

      Dim pageTile As String = String.Format("{0} {1}", pageName, instance).TrimEnd
      stb.Append(hs.GetPageHeader(pageName, pageTile, "", "", False, False))

      '
      ' Start the page plug-in document division
      '
      stb.Append(clsPageBuilder.DivStart("pluginpage", ""))

      '
      ' A message area for error messages from jquery ajax postback (optional, only needed if using AJAX calls to get data)
      '
      stb.Append(clsPageBuilder.DivStart("divErrorMessage", "class='errormessage'"))
      stb.Append(clsPageBuilder.DivEnd)

      '
      ' Setup page timer
      '
      Me.RefreshIntervalMilliSeconds = 1000 * 3
      stb.Append(Me.AddAjaxHandlerPost("id=timer", pageName))

      If WEBUserIsAuthorized(LoggedInUser, USER_ROLES_AUTHORIZED) = False Then
        '
        ' Current user not authorized
        '
        stb.Append(WebUserNotUnauthorized(LoggedInUser))
      Else
        '
        ' Specific page starts here
        '
        stb.Append(BuildContent)
      End If

      '
      ' End the page plug-in document division
      '
      stb.Append(clsPageBuilder.DivEnd)

      '
      ' Add the body html to the page
      '
      Me.AddBody(stb.ToString)

      '
      ' Return the full page
      '
      Return Me.BuildPage()

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "GetPagePlugin")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Builds the HTML content
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildContent() As String

    Try

      Dim stb As New StringBuilder

      stb.AppendLine("<table border='0' cellpadding='0' cellspacing='0' width='1000'>")
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td width='1000' align='center' style='color:#FF0000; font-size:14pt; height:30px;'><strong><div id='divMessage'>&nbsp;</div></strong></td>")
      stb.AppendLine(" </tr>")
      stb.AppendLine(" <tr>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>", BuildTabs())
      stb.AppendLine(" </tr>")
      stb.AppendLine("</table>")

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildContent")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Builds the jQuery Tabss
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function BuildTabs() As String

    Try

      Dim stb As New StringBuilder
      Dim tabs As clsJQuery.jqTabs = New clsJQuery.jqTabs("oTabs", Me.PageName)
      Dim tab As New clsJQuery.Tab

      tabs.postOnTabClick = True

      tab.tabTitle = "Status"
      tab.tabDIVID = "tabStatus"
      tab.tabContent = "<div id='divStatus'>" & BuildTabStatus() & "</div>"
      tabs.tabs.Add(tab)

      tab = New clsJQuery.Tab
      tab.tabTitle = "Options"
      tab.tabDIVID = "tabOptions"
      tab.tabContent = "<div id='divOptions'></div>"
      tabs.tabs.Add(tab)

      tab = New clsJQuery.Tab
      tab.tabTitle = "NetCam Types"
      tab.tabDIVID = "tabNetCamTypes"
      tab.tabContent = "<div id='divNetCamTypes'></div>"
      tabs.tabs.Add(tab)

      tab = New clsJQuery.Tab
      tab.tabTitle = "NetCam Controls"
      tab.tabDIVID = "tabNetCamControls"
      tab.tabContent = "<div id='divNetCamControls'></div>"
      tabs.tabs.Add(tab)

      tab = New clsJQuery.Tab
      tab.tabTitle = "NetCam Devices"
      tab.tabDIVID = "tabNetCamDevices"
      tab.tabContent = "<div id='divNetCamDevices'></div>"
      tabs.tabs.Add(tab)

      tab = New clsJQuery.Tab
      tab.tabTitle = "Event Snapshots"
      tab.tabDIVID = "tabEventSnapshots"
      tab.tabContent = "<div id='divEventSnapshots'></div>"
      tabs.tabs.Add(tab)

      tab = New clsJQuery.Tab
      tab.tabTitle = "Latest Snapshots"
      tab.tabDIVID = "tabSnapshots"
      tab.tabContent = "<div id='divCameras'></div>"
      tabs.tabs.Add(tab)

      Return tabs.Build

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildTabs")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Status Tab
  ''' </summary>
  ''' <param name="Rebuilding"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildTabStatus(Optional ByVal Rebuilding As Boolean = False) As String

    Try

      Dim stb As New StringBuilder

      stb.AppendLine(clsPageBuilder.FormStart("frmStatus", "frmStatus", "Post"))

      stb.AppendLine("<div>")
      stb.AppendLine("<table>")

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td>")
      stb.AppendLine("   <fieldset>")
      stb.AppendLine("    <legend> Plug-In Status </legend>")
      stb.AppendLine("    <table style=""width: 100%"">")
      stb.AppendLine("     <tr>")
      stb.AppendLine("      <td style=""width: 20%""><strong>Name:</strong></td>")
      stb.AppendFormat("    <td style=""text-align: right"">{0}</td>", IFACE_NAME)
      stb.AppendLine("     </tr>")
      stb.AppendLine("     <tr>")
      stb.AppendLine("      <td style=""width: 20%""><strong>Status:</strong></td>")
      stb.AppendFormat("    <td style=""text-align: right"">{0}</td>", "OK")
      stb.AppendLine("     </tr>")
      stb.AppendLine("     <tr>")
      stb.AppendLine("      <td style=""width: 20%""><strong>Version:</strong></td>")
      stb.AppendFormat("    <td style=""text-align: right"">{0}</td>", HSPI.Version)
      stb.AppendLine("     </tr>")
      stb.AppendLine("    </table>")
      stb.AppendLine("   </fieldset>")
      stb.AppendLine("  </td>")
      stb.AppendLine(" </tr>")

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td>")
      stb.AppendLine("   <fieldset>")
      stb.AppendLine("    <legend> Network Camera Status </legend>")
      stb.AppendLine("    <table style=""width: 100%"">")
      stb.AppendLine("     <tr>")
      stb.AppendLine("      <td style=""width: 20%; white-space:nowrap""><strong>Network Cameras:</strong></td>")
      stb.AppendFormat("    <td style=""text-align: right"">{0}</td>", GetNetCamCount())
      stb.AppendLine("     </tr>")
      stb.AppendLine("     <tr>")
      stb.AppendLine("      <td style=""width: 20%; white-space:nowrap""><strong>Event Snapshots:</strong></td>")
      stb.AppendFormat("    <td style=""text-align: right"">{0}</td>", Convert.ToInt32(GetSnapshotCount()).ToString("N0"))
      stb.AppendLine("     </tr>")
      stb.AppendLine("    </table>")
      stb.AppendLine("   </fieldset>")
      stb.AppendLine("  </td>")
      stb.AppendLine(" </tr>")

      stb.AppendLine("</table>")
      stb.AppendLine("</div>")

      stb.AppendLine(clsPageBuilder.FormEnd())

      If Rebuilding Then Me.divToUpdate.Add("divStatus", stb.ToString)

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildTabStatus")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Options Tab
  ''' </summary>
  ''' <param name="Rebuilding"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildTabOptions(Optional ByVal Rebuilding As Boolean = False) As String

    Try

      Dim stb As New StringBuilder

      stb.Append(clsPageBuilder.FormStart("frmOptions", "frmOptions", "Post"))

      stb.AppendLine("<table cellspacing='0' width='100%'>")

      '
      ' NetCam Options
      '
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='2'>Network Camera Options</td>")
      stb.AppendLine(" </tr>")

      '
      ' NetCam Options (Maximum Snapshots Events Per Camera)
      '
      Dim selSnapshotEventMax As New clsJQuery.jqDropList("selSnapshotEventMax", Me.PageName, False)
      selSnapshotEventMax.id = "selSnapshotEventMax"
      selSnapshotEventMax.toolTip = "The maximum number of snapshots to keep per network camera."

      Dim txtSnapshotEventMax As String = GetSetting("Options", "SnapshotEventMax", "50")
      For index As Integer = 5 To 20 Step 5
        Dim value As String = index.ToString
        Dim desc As String = String.Format("{0} Events Per Camera", index.ToString)
        selSnapshotEventMax.AddItem(desc, value, index.ToString = txtSnapshotEventMax)
      Next
      For index As Integer = 25 To 1000 Step 25
        Dim value As String = index.ToString
        Dim desc As String = String.Format("{0} Events Per Camera", index.ToString)
        selSnapshotEventMax.AddItem(desc, value, index.ToString = txtSnapshotEventMax)
      Next

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell' style='width: 20%;'>Maximum&nbsp;Snapshots&nbsp;Events&nbsp;Per&nbsp;Camera</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", selSnapshotEventMax.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' NetCam Options (Snapshots Per Page)
      '
      Dim selSnapshotsPerPage As New clsJQuery.jqDropList("selSnapshotsPerPage", Me.PageName, False)
      selSnapshotsPerPage.id = "selSnapshotsPerPage"
      selSnapshotsPerPage.toolTip = "The default number of snapshots to display per page."

      Dim txtSnapshotsPerPage As String = GetSetting("Options", "SnapshotsPerPage", "25")
      For index As Integer = 25 To 200 Step 25
        Dim value As String = index.ToString
        Dim desc As String = String.Format("{0} Snapshots Per Page", index.ToString)
        selSnapshotsPerPage.AddItem(desc, value, index.ToString = txtSnapshotsPerPage)
      Next

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell' style='width: 20%;'>Snapshots&nbsp;Per&nbsp;Page</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", selSnapshotsPerPage.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' NetCam Options (Snapshots Max Width)
      '
      Dim selSnapshotsMaxWidth As New clsJQuery.jqDropList("selSnapshotsMaxWidth", Me.PageName, False)
      selSnapshotsMaxWidth.id = "selSnapshotsMaxWidth"
      selSnapshotsMaxWidth.toolTip = "The default number of snapshots to display per page."

      Dim txtSnapshotsMaxWidth As String = GetSetting("Options", "SnapshotsMaxWidth", "Auto")
      selSnapshotsMaxWidth.AddItem("Auto", "Auto", "Auto" = txtSnapshotsMaxWidth)
      selSnapshotsMaxWidth.AddItem("640 px", "640px", "640px" = txtSnapshotsMaxWidth)
      selSnapshotsMaxWidth.AddItem("320 px", "320px", "320px" = txtSnapshotsMaxWidth)
      selSnapshotsMaxWidth.AddItem("160 px", "160px", "160px" = txtSnapshotsMaxWidth)

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell' style='width: 20%;'>Snapshot&nbsp;Max&nbsp;Width</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", selSnapshotsMaxWidth.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      Dim selRefreshInterval As New clsJQuery.jqDropList("selRefreshInterval", Me.PageName, False)
      selRefreshInterval.id = "selRefreshInterval"
      selRefreshInterval.toolTip = "Specify how often the plug-in should refresh the snapshots."

      Dim txtRefreshInterval As String = GetSetting("Options", "SnapshotRefreshInterval", gSnapshotRefreshInterval.ToString)
      selRefreshInterval.AddItem("Disabled", "0", "0" = txtRefreshInterval)
      For index As Integer = 2 To 300
        Dim value As String = index.ToString
        Dim desc As String = String.Format("{0} Seconds", index.ToString)
        selRefreshInterval.AddItem(desc, value, index.ToString = txtRefreshInterval)
      Next

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'>Snapshot&nbsp;Refresh&nbsp;Interval</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", selRefreshInterval.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' E-Mail Notification Options
      '
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='2'>E-Mail Notification Options</td>")
      stb.AppendLine(" </tr>")

      Dim strEmailFromDefault As String = hs.GetINISetting("Settings", "gSMTPFrom", "")
      Dim strEmailRcptTo As String = hs.GetINISetting("Settings", "gSMTPTo", "")

      '
      ' E-Mail Notification Options (Send Email)
      '
      Dim selSendEmail As New clsJQuery.jqDropList("selSendEmail", Me.PageName, False)
      selSendEmail.id = "selSendEmail"
      selSendEmail.toolTip = "Enable sending snapshot event e-mail notifications using the plug-in."

      Dim bSendEmail As Boolean = CBool(GetSetting("EmailNotification", "EmailEnabled", False))
      Dim txtSendEmail As String = IIf(bSendEmail = True, "1", "0")

      selSendEmail.AddItem("No", "0", txtSendEmail = "0")
      selSendEmail.AddItem("Yes", "1", txtSendEmail = "1")

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'>Email Notification</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", selSendEmail.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' E-Mail Notification Options (Email To)
      '
      Dim txtEmailRcptTo As String = GetSetting("EmailNotification", "EmailRcptTo", strEmailRcptTo)
      Dim tbEmailRcptTo As New clsJQuery.jqTextBox("txtEmailRcptTo", "text", txtEmailRcptTo, PageName, 60, False)
      tbEmailRcptTo.id = "txtEmailRcptTo"
      tbEmailRcptTo.promptText = "The default e-mail notification recipient address."
      tbEmailRcptTo.toolTip = tbEmailRcptTo.promptText
      tbEmailRcptTo.enabled = bSendEmail

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell' style=""width: 20%"">Email To</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", tbEmailRcptTo.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' E-Mail Notification Options (Email From)
      '
      Dim txtEmailFrom As String = GetSetting("EmailNotification", "EmailFrom", strEmailFromDefault)
      Dim tbEmailFrom As New clsJQuery.jqTextBox("txtEmailFrom", "text", txtEmailFrom, PageName, 60, False)
      tbEmailFrom.id = "txtEmailFrom"
      tbEmailFrom.promptText = "The default e-mail notification sender address."
      tbEmailFrom.toolTip = tbEmailFrom.promptText
      tbEmailFrom.enabled = bSendEmail

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell' style=""width: 20%"">Email From</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", tbEmailFrom.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' E-Mail Notification Options (Email Subject)
      '
      Dim txtEmailSubject As String = GetSetting("EmailNotification", "EmailSubject", EMAIL_SUBJECT)
      Dim tbEmailSubject As New clsJQuery.jqTextBox("txtEmailSubject", "text", txtEmailSubject, PageName, 60, False)
      tbEmailSubject.id = "txtEmailSubject"
      tbEmailSubject.promptText = "The default e-mail notification subject."
      tbEmailSubject.toolTip = tbEmailSubject.promptText
      tbEmailSubject.enabled = bSendEmail

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell' style=""width: 20%"">Email Subject</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", tbEmailSubject.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' E-Mail Notification Options (Email Body Template)
      '
      Dim jqButton1 As New clsJQuery.jqButton("btnSaveEmailBody", "Save", Me.PageName, True)
      Dim chkResetEmailBody As New clsJQuery.jqCheckBox("chkResetEmailBody", "&nbsp;Reset To Default", Me.PageName, True, False)
      chkResetEmailBody.checked = False
      chkResetEmailBody.enabled = bSendEmail

      Dim txtEmailBodyDisabled As String = IIf(bSendEmail = True, "", "disabled")
      Dim txtEmailBody As String = GetSetting("EmailNotification", "EmailBody", EMAIL_BODY_TEMPLATE)
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell' style=""width: 20%"">Email Body Template</td>")
      stb.AppendFormat("  <td class='tablecell'><textarea {0} rows='6' cols='60' name='txtEmailBody'>{1}</textarea>{2}{3}</td>{4}", txtEmailBodyDisabled, _
                                                                                                                                    txtEmailBody.Trim.Replace("~", vbCrLf), _
                                                                                                                                    jqButton1.Build(), _
                                                                                                                                    chkResetEmailBody.Build, _
                                                                                                                                    vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' Archive Options
      '
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='2'>Archive Options</td>")
      stb.AppendLine(" </tr>")

      '
      ' Archive Options (Archive Events)
      '
      Dim bArchiveEnabled As Boolean = CBool(GetSetting("Archive", "ArchiveEnabled", False))
      Dim txtArchiveEnabled As String = IIf(bArchiveEnabled = True, "1", "0")

      Dim selArchiveEnabled As New clsJQuery.jqDropList("selArchiveEnabled", Me.PageName, False)
      selArchiveEnabled.id = "selArchiveEnabled"
      selArchiveEnabled.toolTip = "Enable archiving snapshot events to archive directory."

      selArchiveEnabled.AddItem("No", "0", txtArchiveEnabled = "0")
      selArchiveEnabled.AddItem("Yes", "1", txtArchiveEnabled = "1")

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'>Archive&nbsp;Events</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", selArchiveEnabled.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' Archive Options (Archive Limit)
      '
      Dim selArchiveLimit As New clsJQuery.jqDropList("selArchiveLimit", Me.PageName, False)
      selArchiveLimit.id = "selArchiveLimit"
      selArchiveLimit.toolTip = "The maximum size of all archived snapshot events to be stored in the archive directory."
      selArchiveLimit.enabled = bArchiveEnabled

      Dim txtArchiveLimit As String = GetSetting("Archive", "ArchiveLimit", "1000")
      selArchiveLimit.AddItem("Unlimited", "0", txtArchiveLimit = "0")
      For index As Integer = 100 To 5000 Step 100
        Dim value As String = index.ToString
        Dim desc As String = String.Format("{0} MB", index.ToString)
        selArchiveLimit.AddItem(desc, value, index.ToString = txtArchiveLimit)
      Next

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'>Archive&nbsp;Limit</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", selArchiveLimit.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' Archive Options (Archive Files)
      '
      Dim selArchiveFiles As New clsJQuery.jqDropList("selArchiveFiles", Me.PageName, False)
      selArchiveFiles.id = "selArchiveFiles"
      selArchiveFiles.toolTip = "Specify what event snapshots should be archived."
      selArchiveFiles.enabled = bArchiveEnabled

      Dim txtArchiveFiles As String = GetSetting("Archive", "ArchiveFiles", "1")
      selArchiveFiles.AddItem("Compressed File", "1", txtArchiveFiles = "1")
      selArchiveFiles.AddItem("Video File", "2", txtArchiveFiles = "2")
      selArchiveFiles.AddItem("Compressed File + Video File", "3", txtArchiveFiles = "3")

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'>Archive&nbsp;Files</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", selArchiveFiles.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' Archive Options (Archive Directory)
      '
      Dim txtArchiveDirectory As String = GetSetting("Archive", "ArchiveDir", "")
      Dim tbArchiveDirectory As New clsJQuery.jqTextBox("txtArchiveDirectory", "text", txtArchiveDirectory, PageName, 60, False)
      tbArchiveDirectory.id = "txtArchiveDirectory"
      tbArchiveDirectory.promptText = "The absolute path to the archive directory (e.g.  C:\LocalDir or \\HostnameOrIPAddress\SharedFolder)."
      tbArchiveDirectory.toolTip = tbArchiveDirectory.promptText
      tbArchiveDirectory.enabled = bArchiveEnabled

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell' style=""width: 20%"">Archive&nbsp;Directory</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", tbArchiveDirectory.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' FTP Archive Options
      '
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='2'>FTP Archive Options</td>")
      stb.AppendLine(" </tr>")

      '
      ' FTP Archive Options (Archive Events)
      '
      Dim selFTPArchiveEnabled As New clsJQuery.jqDropList("selFTPArchiveEnabled", Me.PageName, False)
      selFTPArchiveEnabled.id = "selFTPArchiveEnabled"
      selFTPArchiveEnabled.toolTip = "Enable archiving snapshot events to FTP archive."

      Dim bFTPArchive As Boolean = CBool(GetSetting("FTPArchive", "ArchiveEnabled", False))
      Dim txtFTPArchiveEnabled As String = IIf(bFTPArchive = True, "1", "0")

      selFTPArchiveEnabled.AddItem("No", "0", txtFTPArchiveEnabled = "0")
      selFTPArchiveEnabled.AddItem("Yes", "1", txtFTPArchiveEnabled = "1")

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'>Archive&nbsp;Events</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", selFTPArchiveEnabled.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' FTP Archive Options (FTP Archive Limit)
      '
      Dim selFTPArchiveLimit As New clsJQuery.jqDropList("selFTPArchiveLimit", Me.PageName, False)
      selFTPArchiveLimit.id = "selFTPArchiveLimit"
      selFTPArchiveLimit.toolTip = "The maximum size of all archived snapshot events to be stored in the archive directory."
      selFTPArchiveLimit.enabled = bFTPArchive

      Dim txtFTPArchiveLimit As String = GetSetting("FTPArchive", "ArchiveLimit", "50")
      selFTPArchiveLimit.AddItem("Unlimited", "0", txtArchiveLimit = "0")
      For index As Integer = 5 To 1000 Step 5
        Dim value As String = index.ToString
        Dim desc As String = String.Format("{0} MB", index.ToString)
        selFTPArchiveLimit.AddItem(desc, value, index.ToString = txtFTPArchiveLimit)
      Next

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'>Archive&nbsp;Limit</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", selFTPArchiveLimit.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' FTP Archive Options (FTP Archive Files)
      '
      Dim selFTPArchiveFiles As New clsJQuery.jqDropList("selFTPArchiveFiles", Me.PageName, False)
      selFTPArchiveFiles.id = "selFTPArchiveFiles"
      selFTPArchiveFiles.toolTip = "Specify what should be copied to the FTP archive."
      selFTPArchiveFiles.enabled = bFTPArchive

      Dim txtFTPArchiveFiles As String = GetSetting("FTPArchive", "ArchiveFiles", "1")
      selFTPArchiveFiles.AddItem("Compressed File", "1", txtFTPArchiveFiles = "1")
      selFTPArchiveFiles.AddItem("Video File", "2", txtFTPArchiveFiles = "2")
      selFTPArchiveFiles.AddItem("Compressed File + Video File", "3", txtFTPArchiveFiles = "3")

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'>Archive&nbsp;Files</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", selFTPArchiveFiles.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' FTP Archive Options (FTP Archive Directory)
      '
      Dim txtFTPArchiveDir As String = GetSetting("FTPArchive", "ArchiveDir", "/ultranetcam")
      Dim tbFTPArchiveDir As New clsJQuery.jqTextBox("txtFTPArchiveDir", "text", txtFTPArchiveDir, PageName, 60, False)
      tbFTPArchiveDir.id = "txtFTPArchiveDir"
      tbFTPArchiveDir.promptText = "The directory on the FTP archive to store the snapshot events."
      tbFTPArchiveDir.toolTip = tbFTPArchiveDir.promptText
      tbFTPArchiveDir.enabled = bFTPArchive

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell' style=""width: 20%"">Archive&nbsp;Directory</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", tbFTPArchiveDir.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' FTP Archive Options (FTP Server)
      '
      Dim txtFTPServerName As String = GetSetting("FTPClient", "ServerName", "")
      Dim tbFTPServerName As New clsJQuery.jqTextBox("txtFTPServerName", "text", txtFTPServerName, PageName, 60, False)
      tbFTPServerName.id = "txtFTPServerName"
      tbFTPServerName.promptText = "Enter the fully qualified domain name or IP address of the FTP server."
      tbFTPServerName.toolTip = tbFTPServerName.promptText
      tbFTPServerName.enabled = bFTPArchive

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell' style=""width: 20%"">FTP Server Name</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", tbFTPServerName.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' FTP Archive Options (FTP User Name)
      '
      Dim txtFTPUserName As String = GetSetting("FTPClient", "UserName", "")
      Dim tbFTPUserName As New clsJQuery.jqTextBox("txtFTPUserName", "text", txtFTPUserName, PageName, 60, False)
      tbFTPUserName.id = "txtFTPUserName"
      tbFTPUserName.promptText = "Enter your FTP user name."
      tbFTPUserName.toolTip = tbFTPUserName.promptText
      tbFTPUserName.enabled = bFTPArchive

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell' style=""width: 20%"">FTP User Name</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", tbFTPUserName.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' FTP Archive Options (FTP Password)
      '
      Dim txtFTPUserPass As String = GetSetting("FTPClient", "UserPass", "")
      Dim tbFTPUserPass As New clsJQuery.jqTextBox("txtFTPUserPass", "text", "", PageName, 60, False)
      tbFTPUserPass.id = "txtFTPUserPass"
      tbFTPUserPass.promptText = "Enter your FTP user password (the password field will be emtpy after a page refresh)."
      tbFTPUserPass.toolTip = tbFTPUserPass.promptText
      tbFTPUserPass.enabled = bFTPArchive

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell' style=""width: 20%"">FTP User Password</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", tbFTPUserPass.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' Snapshot Event to Video Options
      '
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='2'>Snapshot Event to Video Options</td>")
      stb.AppendLine(" </tr>")

      '
      ' Snapshot Event to Video Options (FFmpeg Status)
      '
      Dim selFFmpegStatus As New clsJQuery.jqDropList("selFFmpegStatus", Me.PageName, False)
      selFFmpegStatus.id = "selFFmpegStatus"
      selFFmpegStatus.toolTip = "The status indicating if the FFmpeg 3rd party program is installed.  This program is required to convert snapshots to MP4 video."

      Dim strFFmpegStatus As String = GetFFmpegStatus()
      selFFmpegStatus.AddItem(strFFmpegStatus, strFFmpegStatus, True)

      Dim downloadURL As String = String.Format("<a target='_blank' href='{0}'>{1}</a>", "http://ffmpeg.org/download.html", "FFmpeg Download (See FFmpeg Windows Builds)")
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'>FFmpeg&nbsp;Status</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}{1}</td>{2}", selFFmpegStatus.Build, downloadURL, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' Snapshot Event to Video Options (Convert Snapshot Event to Video)
      '
      Dim selVideoSnapshotCount As New clsJQuery.jqDropList("selVideoSnapshotCount", Me.PageName, False)
      selVideoSnapshotCount.id = "selVideoSnapshotCount"
      selVideoSnapshotCount.toolTip = "Snapshots will automatically be converted to an MP4 video if the number of snapshots exceeds this value."

      Dim txtVideoSnapshotCounts As String = GetSetting("ConvertToVideo", "EventSnapshotCount", gEventSnapshotCount)
      For index As Integer = 10 To 120 Step 10
        Dim desc As String = String.Format("If event contains {0} snapshots or greater", index.ToString)
        Dim value As String = index.ToString
        selVideoSnapshotCount.AddItem(desc, value, index.ToString = txtVideoSnapshotCounts)
      Next

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'>Convert&nbsp;Snapshot&nbsp;Event&nbsp;to&nbsp;Video</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", selVideoSnapshotCount.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' Web Page Access
      '
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='2'>Web Page Access</td>")
      stb.AppendLine(" </tr>")

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'>Authorized User Roles</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", BuildWebPageAccessCheckBoxes, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' Application Options
      '
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='2'>Application Options</td>")
      stb.AppendLine(" </tr>")

      '
      ' Application Options (Foscam Auto Discovery)
      '
      Dim selFoscamAutoDiscovery As New clsJQuery.jqDropList("selFoscamAutoDiscovery", Me.PageName, False)
      selFoscamAutoDiscovery.id = "selFoscamAutoDiscovery"
      selFoscamAutoDiscovery.toolTip = "Enable Foscam Auto Discovery."

      Dim bFoscamAutoDiscovery As Boolean = CBool(GetSetting("Options", "FoscamAutoDiscovery", False))
      Dim txtFoscamAutoDiscovery As String = IIf(bFoscamAutoDiscovery = True, "1", "0")

      selFoscamAutoDiscovery.AddItem("No", "0", txtFoscamAutoDiscovery = "0")
      selFoscamAutoDiscovery.AddItem("Yes", "1", txtFoscamAutoDiscovery = "1")
      selFoscamAutoDiscovery.autoPostBack = True

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'>Foscam Auto Discovery</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", selFoscamAutoDiscovery.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' Application Logging Level
      '
      Dim selLogLevel As New clsJQuery.jqDropList("selLogLevel", Me.PageName, False)
      selLogLevel.id = "selLogLevel"
      selLogLevel.toolTip = "Specifies the plug-in logging level."

      Dim itemValues As Array = System.Enum.GetValues(GetType(LogLevel))
      Dim itemNames As Array = System.Enum.GetNames(GetType(LogLevel))

      For i As Integer = 0 To itemNames.Length - 1
        Dim itemSelected As Boolean = IIf(gLogLevel = itemValues(i), True, False)
        selLogLevel.AddItem(itemNames(i), itemValues(i), itemSelected)
      Next
      selLogLevel.autoPostBack = True

      stb.AppendLine(" <tr>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", "Logging&nbsp;Level", vbCrLf)
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", selLogLevel.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      stb.AppendLine("</table")

      stb.Append(clsPageBuilder.FormEnd())

      If Rebuilding Then Me.divToUpdate.Add("divOptions", stb.ToString)

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildTabOptions")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Device Types Tab
  ''' </summary>
  ''' <param name="Rebuilding"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildTabNetCamTypes(Optional ByVal Rebuilding As Boolean = False) As String

    Try

      Dim stb As New StringBuilder

      stb.Append(clsPageBuilder.FormStart("frmNetCamTypes", "frmNetCamTypes", "Post"))

      stb.AppendLine("<div id='divNetCamTypesTable'>")
      stb.AppendLine(BuildNetCamTypesTable())
      stb.AppendLine("</div>")

      stb.Append(clsPageBuilder.FormEnd())

      If Rebuilding Then Me.divToUpdate.Add("divNetCamTypes", stb.ToString)

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildTabNetCamTypes")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Device Types table
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildNetCamTypesTable(Optional ByVal Rebuilding As Boolean = False) As String

    Try

      Dim stb As New StringBuilder

      stb.AppendLine("<table id='table_netcam_types' class='display compact' style='width:100%'>")
      stb.AppendFormat("<caption class='tableheader'>{0}</caption>{1}", "Network Camera Types", vbCrLf)
      stb.AppendLine(" <thead>")
      stb.AppendLine("  <tr>")
      stb.AppendLine("   <th>Vendor Name</th>")
      stb.AppendLine("   <th>Model Type</th>")
      stb.AppendLine("   <th>Snapshot Path</th>")
      stb.AppendLine("   <th>Video-stream Path</th>")
      stb.AppendLine("   <th>Action</th>")
      stb.AppendLine("  </tr>")
      stb.AppendLine(" </thead>")
      stb.AppendLine("<tbody>")

      Using MyDataTable As DataTable = GetNetCamTypesFromDB()

        Dim iRowIndex As Integer = 1
        If MyDataTable.Columns.Contains("netcam_type") Then

          For Each row As DataRow In MyDataTable.Rows
            Dim netcam_type As Integer = row("netcam_type")
            Dim netcam_vendor As String = row("netcam_vendor")
            Dim netcam_model As String = row("netcam_model")
            Dim snapshot_path As String = row("snapshot_path")
            Dim videostream_path As String = row("videostream_path")

            stb.AppendFormat("  <tr id='{0}'>", netcam_type.ToString)
            stb.AppendFormat("   <td>{0}</td>{1}", netcam_vendor, vbCrLf)
            stb.AppendFormat("   <td>{0}</td>{1}", netcam_model, vbCrLf)
            stb.AppendFormat("   <td>{0}</td>{1}", HttpUtility.HtmlEncode(snapshot_path), vbCrLf)
            stb.AppendFormat("   <td>{0}</td>{1}", HttpUtility.HtmlEncode(videostream_path), vbCrLf)
            stb.AppendFormat("   <td>{0}</td>{1}", "", vbCrLf)
            stb.AppendLine("  </tr>")

          Next
        End If

      End Using

      stb.AppendLine("</tbody>")
      stb.AppendLine(" <tfoot>")
      stb.AppendLine("  <tr>")
      stb.AppendLine("   <th>Vendor Name</th>")
      stb.AppendLine("   <th>Model Type</th>")
      stb.AppendLine("   <th>Snapshot Path</th>")
      stb.AppendLine("   <th>Video-stream Path</th>")
      stb.AppendLine("   <th>Action</th>")
      stb.AppendLine("  </tr>")
      stb.AppendLine(" </tfoot>")
      stb.AppendLine("</table>")

      Dim strInfo As String = "Check your network camera documentation for the correct snapshot URL path."
      Dim strHint As String = "Use the replacment variables like $user and $pass instead of embeding the credentials in the snapshot path."

      stb.AppendLine(" <div>&nbsp;</div>")
      stb.AppendLine(" <p>")
      stb.AppendFormat("<img alt='Info' src='/images/hspi_ultranetcam3/ico_info.gif' width='16' height='16' border='0' />&nbsp;{0}<br/>", strInfo)
      stb.AppendFormat("<img alt='Hint' src='/images/hspi_ultranetcam3/ico_hint.gif' width='16' height='16' border='0' />&nbsp;{0}<br/>", strHint)
      stb.AppendLine(" </p>")

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildNetCamTypesTable")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Device Tab
  ''' </summary>
  ''' <param name="Rebuilding"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildTabNetCamControls(ByVal netcam_type As Integer, Optional ByVal Rebuilding As Boolean = False) As String

    Try

      Dim stb As New StringBuilder

      stb.Append(clsPageBuilder.FormStart("frmNetCamControls", "frmNetCamControls", "Post"))

      Dim selNetCamType As New clsJQuery.jqDropList("selNetCamType", Me.PageName, False)
      selNetCamType.id = "selNetCamType"
      selNetCamType.toolTip = "Select the Network Camera type."

      Using MyDataTable As DataTable = GetNetCamTypesFromDB()

        If MyDataTable.Rows.Count > 0 Then

          If netcam_type = 0 Then
            selNetCamType.AddItem("[Select NetCam Type]", "0", netcam_type = 0)
          End If

          For Each r As DataRow In MyDataTable.Rows
            Dim value As String = r("netcam_type")  ' This is the Id
            Dim desc As String = r("netcam_name")
            selNetCamType.AddItem(desc, value, netcam_type = value)
          Next

        End If

      End Using

      stb.AppendFormat("<p>Network&nbsp;Camera&nbsp;Type:&nbsp;{0}</p>", selNetCamType.Build())

      If netcam_type = 0 Then

        Dim strInfo As String = "Select the Network Camera Type from the dropdown list to display or edit the Network Camera controls."
        Dim strHint As String = "For more information, please check the UltraNetCam3 HSPI User's Guide."

        stb.AppendLine(" <div>&nbsp;</div>")
        stb.AppendLine(" <p>")
        stb.AppendFormat("<img alt='Info' src='/images/hspi_ultranetcam3/ico_info.gif' width='16' height='16' border='0' />&nbsp;{0}<br/>", strInfo)
        stb.AppendFormat("<img alt='Hint' src='/images/hspi_ultranetcam3/ico_hint.gif' width='16' height='16' border='0' />&nbsp;{0}<br/>", strHint)
        stb.AppendLine(" </p>")

      Else

        stb.AppendLine("<div id='divNetCamControlsTable'>")
        stb.AppendLine(BuildNetCamControlsTable(netcam_type))
        stb.AppendLine("</div>")

      End If

      stb.Append(clsPageBuilder.FormEnd())

      If Rebuilding Then Me.divToUpdate.Add("divNetCamControls", stb.ToString)

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildTabNetCamControls")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Controls table
  ''' </summary>
  ''' <param name="netcam_type"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildNetCamControlsTable(netcam_type As Integer) As String

    Try

      Dim stb As New StringBuilder

      stb.AppendLine("<table id='table_netcam_controls' class='display compact' style='width:100%'>")
      stb.AppendFormat("<caption class='tableheader'>{0}</caption>{1}", "Network Camera Controls", vbCrLf)
      stb.AppendLine(" <thead>")
      stb.AppendLine("  <tr>")
      stb.AppendLine("   <th>Control Name</th>")
      stb.AppendLine("   <th>Control URL</th>")
      stb.AppendLine("   <th>Action</th>")
      stb.AppendLine("  </tr>")
      stb.AppendLine(" </thead>")
      stb.AppendLine("<tbody>")

      Using MyDataTable As DataTable = GetNetCamControlsFromDB(netcam_type)

        If MyDataTable.Columns.Contains("control_id") Then

          For Each row As DataRow In MyDataTable.Rows
            Dim control_id As Integer = row("control_id")
            Dim control_name As String = row("control_name")
            Dim control_url As String = row("control_url")

            stb.AppendFormat("  <tr id='{0}'>", control_id.ToString)
            stb.AppendFormat("   <td>{0}</td>{1}", control_name, vbCrLf)
            stb.AppendFormat("   <td>{0}</td>{1}", control_url, vbCrLf)
            stb.AppendFormat("   <td>{0}</td>{1}", "", vbCrLf)
            stb.AppendLine("  </tr>")
          Next

        End If

      End Using

      stb.AppendLine("</tbody>")
      stb.AppendLine(" <tfoot>")
      stb.AppendLine("  <tr>")
      stb.AppendLine("   <th>Control Name</th>")
      stb.AppendLine("   <th>Control URL</th>")
      stb.AppendLine("   <th>Action</th>")
      stb.AppendLine("  </tr>")
      stb.AppendLine(" </tfoot>")
      stb.AppendLine("</table>")

      Dim strInfo As String = "Edit the NetCam Controls using the action links."
      Dim strHint As String = "Any errors that occur during editing will be written to the HomeSeer log."

      stb.AppendLine(" <div>&nbsp;</div>")
      stb.AppendLine(" <p>")
      stb.AppendFormat("<img alt='Info' src='/images/hspi_ultranetcam3/ico_info.gif' width='16' height='16' border='0' />&nbsp;{0}<br/>", strInfo)
      stb.AppendFormat("<img alt='Hint' src='/images/hspi_ultranetcam3/ico_hint.gif' width='16' height='16' border='0' />&nbsp;{0}<br/>", strHint)
      stb.AppendLine(" </p>")

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildNetCamControlsTable")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Device Tab
  ''' </summary>
  ''' <param name="Rebuilding"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildTabNetCamDevices(Optional ByVal Rebuilding As Boolean = False) As String

    Try

      Dim stb As New StringBuilder

      stb.Append(clsPageBuilder.FormStart("frmNetCamDevices", "frmNetCamDevices", "Post"))

      stb.AppendLine("<div id='divNetCamDevicesTable'>")
      stb.AppendLine(BuildNetCamDevicesTable())
      stb.AppendLine("</div>")

      stb.Append(clsPageBuilder.FormEnd())

      If Rebuilding Then Me.divToUpdate.Add("divNetCamDevices", stb.ToString)

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildTabNetCamDevices")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Devices table
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildNetCamDevicesTable(Optional ByVal Rebuilding As Boolean = False) As String

    Try

      Dim stb As New StringBuilder

      stb.AppendLine("<table id='table_netcam_devices' class='display compact' style='width:100%'>")
      stb.AppendFormat("<caption class='tableheader'>{0}</caption>{1}", "Network Cameras", vbCrLf)
      stb.AppendLine(" <thead>")
      stb.AppendLine("  <tr>")
      stb.AppendLine("   <th>Name</th>")
      stb.AppendLine("   <th>IP Address</th>")
      stb.AppendLine("   <th>Port</th>")
      stb.AppendLine("   <th>Model</th>")
      stb.AppendLine("   <th>User Id</th>")
      stb.AppendLine("   <th>User Password</th>")
      stb.AppendLine("   <th>Action</th>")
      stb.AppendLine("  </tr>")
      stb.AppendLine(" </thead>")
      stb.AppendLine("<tbody>")

      Using MyDataTable As DataTable = GetNetCamDevicesFromDB()

        If MyDataTable.Columns.Contains("netcam_id") Then

          For Each row As DataRow In MyDataTable.Rows
            Dim netcam_id As Integer = row("netcam_id")
            Dim netcam_name As String = row("netcam_name")
            Dim netcam_address As String = row("netcam_address")
            Dim netcam_port As String = row("netcam_port")
            Dim netcam_type As String = row("netcam_type")
            Dim auth_user As String = row("auth_user")
            Dim auth_pass As String = row("auth_pass")

            stb.AppendFormat("  <tr id='{0}'>", netcam_id.ToString)
            stb.AppendFormat("   <td>{0}</td>{1}", netcam_name, vbCrLf)
            stb.AppendFormat("   <td>{0}</td>{1}", netcam_address, vbCrLf)
            stb.AppendFormat("   <td>{0}</td>{1}", netcam_port, vbCrLf)
            stb.AppendFormat("   <td>{0}</td>{1}", netcam_type, vbCrLf)
            stb.AppendFormat("   <td>{0}</td>{1}", auth_user, vbCrLf)
            stb.AppendFormat("   <td>{0}</td>{1}", "", vbCrLf)
            stb.AppendFormat("   <td>{0}</td>{1}", "", vbCrLf)
            stb.AppendLine("  </tr>")

          Next
        End If

      End Using

      stb.AppendLine("</tbody>")
      stb.AppendLine(" <tfoot>")
      stb.AppendLine("  <tr>")
      stb.AppendLine("   <th>Name</th>")
      stb.AppendLine("   <th>IP Address</th>")
      stb.AppendLine("   <th>Port</th>")
      stb.AppendLine("   <th>Model</th>")
      stb.AppendLine("   <th>User Id</th>")
      stb.AppendLine("   <th>User Password</th>")
      stb.AppendLine("   <th>Action</th>")
      stb.AppendLine("  </tr>")
      stb.AppendLine(" </tfoot>")
      stb.AppendLine("</table>")

      Dim strInfo As String = "Edit the network camera devices using the action links."
      Dim strHint As String = "You will need to manually add Network Cameras that do not support auto discovery."

      stb.AppendLine(" <div>&nbsp;</div>")
      stb.AppendLine(" <p>")
      stb.AppendFormat("<img alt='Info' src='/images/hspi_ultranetcam3/ico_info.gif' width='16' height='16' border='0' />&nbsp;{0}<br/>", strInfo)
      stb.AppendFormat("<img alt='Hint' src='/images/hspi_ultranetcam3/ico_hint.gif' width='16' height='16' border='0' />&nbsp;{0}<br/>", strHint)
      stb.AppendLine(" </p>")

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildNetCamDevicesTable")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Snapshot Events Tab
  ''' </summary>
  ''' <param name="Rebuilding"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildTabEventSnapshots(ByVal netcam_id As String, ByVal event_id As String, Optional ByVal Rebuilding As Boolean = False) As String

    Try

      Dim stb As New StringBuilder

      stb.Append(clsPageBuilder.FormStart("frmEventSnapshots", "frmEventSnapshots", "Post"))

      Dim selNetworkCamera As New clsJQuery.jqDropList("selNetworkCamera", Me.PageName, True)
      selNetworkCamera.id = "selNetworkCamera"
      selNetworkCamera.toolTip = "Select the Network Camera"
      selNetworkCamera.autoPostBack = True

      Using NetCamDevices As DataTable = hspi_plugin.GetNetCamDevicesFromDB()

        If netcam_id.Length = 0 Then
          selNetworkCamera.AddItem("[Select Network Camera]", "", "" = netcam_id)
        End If

        For Each row As DataRow In NetCamDevices.Rows

          Dim value As String = String.Format("NetCam{0}", row("netcam_id").ToString.PadLeft(3, "0"))
          Dim desc As String = row("netcam_name")

          selNetworkCamera.AddItem(desc, value, value = netcam_id)
        Next

      End Using

      Dim selNetCamEvent As New clsJQuery.jqDropList("selNetCamEvent", Me.PageName, True)
      selNetCamEvent.id = "selNetCamEvent"
      selNetCamEvent.toolTip = "Select the Network Camera Snapshot Event."
      selNetCamEvent.autoPostBack = True
      selNetCamEvent.enabled = False

      '
      ' Build Event Snapshot Dropdown
      '
      selNetCamEvent.AddItem("[Select Event Snapshot]", "", event_id = "")

      Dim NetCamEvents As SortedList = hspi_plugin.GetNetCamEventSummary(netcam_id)
      If NetCamEvents.Count > 0 Then
        selNetCamEvent.ClearItems()
        selNetCamEvent.AddItem("[Select Event Snapshot]", "", event_id = "")
        selNetCamEvent.enabled = True

        For Each strEventId As String In NetCamEvents.Keys
          Dim value As String = strEventId
          Dim desc As String = String.Format("{0} - [{1} Snapshots]", value, NetCamEvents(strEventId))

          selNetCamEvent.AddItem(desc, value, value = event_id)
        Next
      Else
        selNetCamEvent.ClearItems()
        selNetCamEvent.AddItem("[No Event Snapshots Found]", "", event_id = "")
      End If

      Dim btnDeleteEventSnapshots As New clsJQuery.jqButton("btnDeleteEventSnapshots", "&nbsp;Delete Selected Snapshot Event", Me.PageName, True)
      btnDeleteEventSnapshots.enabled = event_id.Length > 0 And NetCamEvents.Count > 0

      '
      ' General Options
      '
      stb.AppendLine("<table cellspacing='0' width='100%'>")
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader'>Network Camera Event Snapshots</td>")
      stb.AppendLine(" </tr>")
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'>")
      stb.AppendFormat("  Network Camera: {0}&nbsp;", selNetworkCamera.Build)
      stb.AppendFormat("  Snapshot Event:{0}&nbsp;", selNetCamEvent.Build)
      stb.AppendFormat("  {0}&nbsp;", btnDeleteEventSnapshots.Build)
      stb.AppendLine("  </td>")
      stb.AppendLine(" </tr>")

      If event_id.Length > 0 Then

        Using CameraSnapshotInfo As DataTable = hspi_plugin.GetCameraSnapshotInfo(netcam_id, event_id)

          If CameraSnapshotInfo.Rows.Count > 0 Then

            stb.AppendLine(" <tr>")
            stb.AppendLine("  <td class='tablecell'>")

            For Each row As DataRow In CameraSnapshotInfo.Rows
              Dim snapshot_filename As String = row("snapshot_filename")
              Dim thumbnail_filename As String = row("thumbnail_filename")
              Dim creation_date As String = row("creation_date")
              Dim snapshot_age As String = row("snapshot_age")

              stb.AppendFormat("  <div class=""tablecell"" style=""float:left; border: solid black 1px; margin:2px; text-align:center; width:{0};"">", "auto")
              stb.AppendFormat("   <a id=""lnk_{0}"" href=""{1}"" title=""{2}"" data-lightbox=""lightbox[1]"">", event_id, snapshot_filename, netcam_id)
              stb.AppendFormat("    <img id=""img_{0}"" rel=""lightbox[1]"" style=""width:100%"" src='{1}' />", event_id, thumbnail_filename)
              stb.AppendLine("   </a>")
              stb.AppendFormat("   <div>{0}</div><div>{1}</div><div>{2}</div>", row("event_id"), creation_date, snapshot_age)
              stb.AppendLine("  </div>")

            Next

            stb.AppendLine("  </td>")
            stb.AppendLine(" </tr>")

          End If

        End Using

      End If

      stb.AppendLine("</table>")

      stb.Append(clsPageBuilder.FormEnd())

      Dim strInfo As String = "Select the Network Camera, then select the Event Snapshot from the dropdown lists."
      Dim strHint As String = "Event Snapshots are created using a HomeSeer event.&nbsp; " & _
                              "For more information, please check the UltraNetCam3 HSPI User's Guide."

      stb.AppendLine(" <div>&nbsp;</div>")
      stb.AppendLine(" <p>")
      stb.AppendFormat("<img alt='Info' src='/images/hspi_ultranetcam3/ico_info.gif' width='16' height='16' border='0' />&nbsp;{0}<br/>", strInfo)
      stb.AppendFormat("<img alt='Hint' src='/images/hspi_ultranetcam3/ico_hint.gif' width='16' height='16' border='0' />&nbsp;{0}<br/>", strHint)
      stb.AppendLine(" </p>")

      If Rebuilding Then Me.divToUpdate.Add("tabEventSnapshots", stb.ToString)

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildTabEventSnapshots")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Options Tab
  ''' </summary>
  ''' <param name="Rebuilding"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildTabCameras(ByVal SnapshotWidth As String, Optional ByVal Rebuilding As Boolean = False) As String

    Try

      Dim stb As New StringBuilder

      stb.Append(clsPageBuilder.FormStart("frmCameras", "frmCameras", "Post"))

      Dim selSnapshotsWidth As New clsJQuery.jqDropList("selSnapshotsWidth", Me.PageName, False)
      If SnapshotWidth.Length = 0 Then SnapshotWidth = gSnapshotsMaxWidth

      selSnapshotsWidth.id = "selSnapshotsWidth"
      selSnapshotsWidth.toolTip = "Specifies the maximum snapshot width."
      selSnapshotsWidth.AddItem("Auto", "Auto", IIf(SnapshotWidth = "Auto", True, False))
      selSnapshotsWidth.AddItem("160 px", "160px", IIf(SnapshotWidth = "160px", True, False))
      selSnapshotsWidth.AddItem("320 px", "320px", IIf(SnapshotWidth = "320px", True, False))
      selSnapshotsWidth.AddItem("640 px", "640px", IIf(SnapshotWidth = "640px", True, False))
      selSnapshotsWidth.autoPostBack = True

      '
      ' General Options
      '
      stb.AppendLine("<table cellspacing='0' width='100%'>")
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='2'>Latest Network Camera Snapshots</td>")
      stb.AppendLine(" </tr>")
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'><div id='lastRefresh'/></td>")
      stb.AppendLine("  <td class='tablecell' align='right'>Snapshot Width: " & selSnapshotsWidth.Build & "</td>")
      stb.AppendLine(" </tr>")
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell' colspan='2'>")

      Dim MyDataTable As DataTable = hspi_plugin.GetCameraSnapshotViewer()
      If MyDataTable.Columns.Contains("netcam_id") Then

        For Each row As DataRow In MyDataTable.Rows
          Dim netcam_name As String = row("netcam_name")
          Dim netcam_id As String = row("netcam_id")
          Dim snapshot_age As String = row("snapshot_age")

          stb.AppendFormat("  <div class=""tablecell"" style=""float:left; border: solid black 1px; margin:2px; text-align:center; width:{0};"">", SnapshotWidth)
          stb.AppendFormat("   <a id=""lnk_{0}"" href=""#"" title=""{1}"" data-lightbox=""lightbox[0]"">", netcam_id, netcam_name)
          stb.AppendFormat("    <img id=""img_{0}"" rel=""lightbox[0]"" style=""width:100%"" />", netcam_id)
          stb.AppendLine("   </a>")
          stb.AppendFormat("   <div>{0}</div><div>{1}</div>", netcam_name, snapshot_age)
          stb.AppendLine("  </div>")

        Next
      End If

      stb.AppendLine("  </td>")
      stb.AppendLine(" </tr>")
      stb.AppendLine("</table>")

      stb.Append(clsPageBuilder.FormEnd())

      '
      ' Update the Refresh Interval
      '
      Dim iRefreshInterval As Integer = 4000
      Dim strRefreshInterval As String = GetSetting("Options", "SnapshotRefreshInterval", gSnapshotRefreshInterval.ToString)
      If IsNumeric(strRefreshInterval) = True Then
        iRefreshInterval *= Integer.Parse(strRefreshInterval)
      End If
      If iRefreshInterval < 4000 Then iRefreshInterval = 4000

      stb.AppendLine("<script>")
      stb.AppendLine("function refreshSnapshots() {")
      stb.AppendLine("  var ticks = new Date().getTime();")
      If MyDataTable.Columns.Contains("netcam_id") Then
        For Each row As DataRow In MyDataTable.Rows
          Dim netcam_id As String = row("netcam_id")
          Dim strSnapshotFilename As String = row("snapshot_filename")

          stb.AppendLine("    $('#img_" & netcam_id & "').attr('src', '" & strSnapshotFilename & "?ticks=' + ticks);")
          stb.AppendLine("    $('#lnk_" & netcam_id & "').attr('href', '" & strSnapshotFilename & "?ticks=' + ticks);")
        Next
      End If
      stb.AppendLine("    $('#lastRefresh').html( new Date() + '' );")
      stb.AppendLine("};")

      stb.AppendLine("$(function() { refreshSnapshots(); setInterval(function() { refreshSnapshots(); }, " & iRefreshInterval.ToString & ");});")
      stb.AppendLine("</script>")

      If Rebuilding Then Me.divToUpdate.Add("divCameras", stb.ToString)

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildTabCameras")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Web Page Access Checkbox List
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function BuildWebPageAccessCheckBoxes()

    Try

      Dim stb As New StringBuilder

      Dim USER_ROLES_AUTHORIZED As Integer = WEBUserRolesAuthorized()

      Dim cb1 As New clsJQuery.jqCheckBox("chkWebPageAccess_Guest", "Guest", Me.PageName, True, True)
      Dim cb2 As New clsJQuery.jqCheckBox("chkWebPageAccess_Admin", "Admin", Me.PageName, True, True)
      Dim cb3 As New clsJQuery.jqCheckBox("chkWebPageAccess_Normal", "Normal", Me.PageName, True, True)
      Dim cb4 As New clsJQuery.jqCheckBox("chkWebPageAccess_Local", "Local", Me.PageName, True, True)

      cb1.id = "WebPageAccess_Guest"
      cb1.checked = CBool(USER_ROLES_AUTHORIZED And USER_GUEST)

      cb2.id = "WebPageAccess_Admin"
      cb2.checked = CBool(USER_ROLES_AUTHORIZED And USER_ADMIN)
      cb2.enabled = False

      cb3.id = "WebPageAccess_Normal"
      cb3.checked = CBool(USER_ROLES_AUTHORIZED And USER_NORMAL)

      cb4.id = "WebPageAccess_Local"
      cb4.checked = CBool(USER_ROLES_AUTHORIZED And USER_LOCAL)

      stb.Append(clsPageBuilder.FormStart("frmWebPageAccess", "frmWebPageAccess", "Post"))

      stb.Append(cb1.Build())
      stb.Append(cb2.Build())
      stb.Append(cb3.Build())
      stb.Append(cb4.Build())

      stb.Append(clsPageBuilder.FormEnd())

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildWebPageAccessCheckBoxes")
      Return "error - " & Err.Description
    End Try

  End Function

#End Region

#Region "Page Processing"

  ''' <summary>
  ''' Post a message to this web page
  ''' </summary>
  ''' <param name="sMessage"></param>
  ''' <remarks></remarks>
  Sub PostMessage(ByVal sMessage As String)

    Try

      Me.divToUpdate.Add("divMessage", sMessage)

      Me.pageCommands.Add("starttimer", "")

      TimerEnabled = True

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "PostMessage")
    End Try

  End Sub

  ''' <summary>
  ''' When a user clicks on any controls on one of your web pages, this function is then called with the post data. You can then parse the data and process as needed.
  ''' </summary>
  ''' <param name="page">The name of the page as registered with hs.RegisterLink or hs.RegisterConfigLink</param>
  ''' <param name="data">The post data</param>
  ''' <param name="user">The name of logged in user</param>
  ''' <param name="userRights">The rights of logged in user</param>
  ''' <returns>Any serialized data that needs to be passed back to the web page, generated by the clsPageBuilder class</returns>
  ''' <remarks></remarks>
  Public Overrides Function postBackProc(page As String, data As String, user As String, userRights As Integer) As String

    Try

      WriteMessage("Entered postBackProc() function.", MessageType.Debug)

      Dim postData As NameValueCollection = HttpUtility.ParseQueryString(data)

      '
      ' Write debug to console
      '
      If gLogLevel >= MessageType.Debug Then
        For Each keyName As String In postData.AllKeys
          Console.WriteLine(String.Format("{0}={1}", keyName, postData(keyName)))
        Next
      End If

      '
      ' Process actions
      '
      Select Case postData("editor_action")
        Case "netcamtypes-create"
          Dim netcam_vendor As String = postData("data[netcam_vendor]").Trim
          Dim netcam_model As String = postData("data[netcam_model]").Trim
          Dim snapshot_path As String = postData("data[snapshot_path]").Trim.Replace("&amp;", "&")
          Dim videostream_path As String = postData("data[videostream_path]").Trim.Replace("&amp;", "&")

          If netcam_vendor.Length = 0 Then
            Return DatatableFieldError("netcam_vendor", "The vendor name field is blank.  This is a required field.")
          ElseIf netcam_model.Length = 0 Then
            Return DatatableFieldError("netcam_model", "The model type field is blank.  This is a required field.")
          ElseIf snapshot_path.Length = 0 Then
            Return DatatableFieldError("snapshot_path", "The snapshot path field is blank.  This is a required field.")
          ElseIf Regex.IsMatch(snapshot_path, "^\/") = False Then
            Return DatatableFieldError("snapshot_path", "The snapshot path must start with a forward slash character.")
          End If

          Dim netcam_type As Integer = hspi_plugin.InsertNetCamType(netcam_vendor, netcam_model, snapshot_path, videostream_path)
          If netcam_type = 0 Then
            Return DatatableError("Unable to add new network camera type due to an unexpected error.")
          Else
            Return DatatableRowNetCamType(netcam_type, netcam_vendor, netcam_model, snapshot_path, videostream_path)
          End If

        Case "netcamtypes-edit"
          Dim netcam_type As String = postData("id")
          Dim netcam_vendor As String = postData("data[netcam_vendor]").Trim
          Dim netcam_model As String = postData("data[netcam_model]").Trim
          Dim snapshot_path As String = postData("data[snapshot_path]").Trim.Replace("&amp;", "&")
          Dim videostream_path As String = postData("data[videostream_path]").Trim.Replace("&amp;", "&")

          If netcam_vendor.Length = 0 Then
            Return DatatableFieldError("netcam_vendor", "The vendor name field is blank.  This is a required field.")
          ElseIf netcam_model.Length = 0 Then
            Return DatatableFieldError("netcam_model", "The model type field is blank.  This is a required field.")
          ElseIf snapshot_path.Length = 0 Then
            Return DatatableFieldError("snapshot_path", "The snapshot path field is blank.  This is a required field.")
          ElseIf Regex.IsMatch(snapshot_path, "^\/") = False Then
            Return DatatableFieldError("snapshot_path", "The snapshot path must start with a forward slash character.")
          End If

          Dim bSuccess As Boolean = hspi_plugin.UpdateNetCamType(netcam_type, netcam_vendor, netcam_model, snapshot_path, videostream_path)
          If bSuccess = False Then
            Return DatatableError("Unable to modify new network camera type due to an unexpected error.")
          Else
            Return DatatableRowNetCamType(netcam_type, netcam_vendor, netcam_model, snapshot_path, videostream_path)
          End If

        Case "netcamtypes-remove"
          Dim netcam_type As String = postData("id[]")
          Dim bSuccess As Boolean = hspi_plugin.DeleteNetCamType(netcam_type)

          If bSuccess = False Then
            Return DatatableError("Unable to delete the network camera type due to an unexpected error.")
          Else
            BuildTabNetCamTypes(True)
            Me.pageCommands.Add("executefunction", "reDrawNetCamTypes()")
            Return "{ }"
          End If

        Case "netcamcontrols-create"
          Dim netcam_type As String = postData("netcam_type")
          Dim control_name As String = postData("data[control_name]").Trim
          Dim control_url As String = postData("data[control_url]").Trim.Replace("&amp;", "&")

          If control_name.Length = 0 Then
            Return DatatableFieldError("control_name", "The control name field is blank.  This is a required field.")
          ElseIf control_url.Length = 0 Then
            Return DatatableFieldError("control_url", "The control URL field is blank.  This is a required field.")
          ElseIf Regex.IsMatch(control_url, "^\/") = False Then
            Return DatatableFieldError("control_url", "The control URL path must start with a forward slash character.")
          End If

          Dim control_id As Integer = hspi_plugin.InsertNetCamControls(netcam_type, control_name, control_url)
          If control_id = 0 Then
            Return DatatableError("Unable to add new network camera control due to an unexpected error.")
          Else
            Return DatatableRowNetCamControl(control_id, control_name, control_url)
          End If

        Case "netcamcontrols-edit"
          Dim control_id As String = postData("id")
          Dim control_name As String = postData("data[control_name]").Trim
          Dim control_url As String = postData("data[control_url]").Trim.Replace("&amp;", "&")

          If control_name.Length = 0 Then
            Return DatatableFieldError("control_name", "The control name field is blank.  This is a required field.")
          ElseIf control_url.Length = 0 Then
            Return DatatableFieldError("control_url", "The control URL field is blank.  This is a required field.")
          ElseIf Regex.IsMatch(control_url, "^\/") = False Then
            Return DatatableFieldError("control_url", "The control URL path must start with a forward slash character.")
          End If

          Dim bSuccess As Boolean = hspi_plugin.UpdateNetCamControls(control_id, control_name, control_url)
          If bSuccess = False Then
            Return DatatableError("Unable to modify the network camera control due to an unexpected error.")
          Else
            Return DatatableRowNetCamControl(control_id, control_name, control_url)
          End If

        Case "netcamcontrols-remove"
          Dim control_id As String = postData("id[]")
          Dim bSuccess As Boolean = hspi_plugin.DeleteNetCamControls(control_id)

          If bSuccess = False Then
            Return DatatableError("Unable to delete the network camera control due to an unexpected error.")
          Else
            BuildTabNetCamTypes(True)
            Me.pageCommands.Add("executefunction", "reDrawNetCamControls()")
            Return "{ }"
          End If

        Case "netcamdevices-create"
          Dim netcam_name As String = postData("data[netcam_name]").Trim
          Dim netcam_address As String = postData("data[netcam_address]").Trim
          Dim netcam_port As String = postData("data[netcam_port]").Trim
          Dim netcam_type As String = postData("data[netcam_type]").Trim
          Dim auth_user As String = postData("data[auth_user]").Trim
          Dim auth_pass As String = postData("data[auth_pass]").Trim

          If netcam_name.Length = 0 Then
            Return DatatableFieldError("netcam_name", "The network camera name field is blank.  This is a required field.")
          ElseIf netcam_address.Length = 0 Then
            Return DatatableFieldError("netcam_address", "The network camera IP address field is blank.  This is a required field.")
            'ElseIf Regex.IsMatch(netcam_address, "^([a-zA-Z0-9]|[a-zA-Z0-9][a-zA-Z0-9\-]{0,61}[a-zA-Z0-9])(\.([a-zA-Z0-9]|[a-zA-Z0-9][a-zA-Z0-9\-]{0,61}[a-zA-Z0-9]))+$") = False Then
            'Return DatatableFieldError("netcam_address", "NetCam address value must contain a hostname or IP address.")
          ElseIf netcam_port.Length = 0 Then
            Return DatatableFieldError("netcam_port", "The network camera TCP port field is blank.  This is a required field.")
          ElseIf Regex.IsMatch(netcam_port, "^\d+$") = False Then
            Return DatatableFieldError("netcam_port", "The network camera TCP port must be numeric.")
          ElseIf netcam_type.Length = 0 Then
            Return DatatableFieldError("netcam_type", "The network camera type field is blank.  This is a required field.")
          ElseIf auth_user.Length = 0 Then
            Return DatatableFieldError("auth_user", "The network camera type field is blank.  This is a required field.")
          End If

          Dim netcam_id As Integer = hspi_plugin.InsertNetCamDevice(netcam_name, netcam_address, netcam_port, netcam_type, auth_user, auth_pass)
          If netcam_id = 0 Then
            Return DatatableError("Unable to add new network camera due to an unexpected error.")
          Else
            Return DatatableRowNetCamDevice(netcam_id, netcam_name, netcam_address, netcam_port, netcam_type, auth_user, auth_pass)
          End If

        Case "netcamdevices-edit"
          Dim netcam_id As String = postData("id").Trim
          Dim netcam_name As String = postData("data[netcam_name]").Trim
          Dim netcam_address As String = postData("data[netcam_address]").Trim
          Dim netcam_port As String = postData("data[netcam_port]").Trim
          Dim netcam_type As String = postData("data[netcam_type]").Trim
          Dim auth_user As String = postData("data[auth_user]").Trim
          Dim auth_pass As String = postData("data[auth_pass]").Trim

          If netcam_name.Length = 0 Then
            Return DatatableFieldError("netcam_name", "The network camera name field is blank.  This is a required field.")
          ElseIf netcam_address.Length = 0 Then
            Return DatatableFieldError("netcam_address", "The network camera IP address field is blank.  This is a required field.")
            'ElseIf Regex.IsMatch(netcam_address, "^([a-zA-Z0-9]|[a-zA-Z0-9][a-zA-Z0-9\-]{0,61}[a-zA-Z0-9])(\.([a-zA-Z0-9]|[a-zA-Z0-9][a-zA-Z0-9\-]{0,61}[a-zA-Z0-9]))+$") = False Then
            'Return DatatableFieldError("netcam_address", "NetCam address value must contain a hostname or IP address.")
          ElseIf netcam_port.Length = 0 Then
            Return DatatableFieldError("netcam_port", "The network camera TCP port field is blank.  This is a required field.")
          ElseIf Regex.IsMatch(netcam_port, "^\d+$") = False Then
            Return DatatableFieldError("netcam_port", "The network camera TCP port must be numeric.")
          ElseIf netcam_type.Length = 0 Then
            Return DatatableFieldError("netcam_type", "The network camera type field is blank.  This is a required field.")
          ElseIf auth_user.Length = 0 Then
            Return DatatableFieldError("auth_user", "The network camera type field is blank.  This is a required field.")
          End If

          Dim bSuccess As Boolean = hspi_plugin.UpdateNetCamDevice(netcam_id, netcam_name, netcam_address, netcam_port, netcam_type, auth_user, auth_pass)
          If bSuccess = False Then
            Return DatatableError("Unable to add new network camera due to an unexpected error.")
          Else
            Return DatatableRowNetCamDevice(netcam_id, netcam_name, netcam_address, netcam_port, netcam_type, auth_user, auth_pass)
          End If

        Case "netcamdevices-remove"
          Dim netcam_id As String = postData("id[]")
          Dim bSuccess As Boolean = hspi_plugin.DeleteNetCamDevice(netcam_id)

          If bSuccess = False Then
            Return DatatableError("Unable to delete the network camera due to an unexpected error.")
          Else
            BuildTabNetCamTypes(True)
            Me.pageCommands.Add("executefunction", "reDrawNetCamDevices()")
            Return "{ }"
          End If

      End Select

      '
      ' Process the post data
      '
      Select Case postData("id")
        Case "tabStatus"
          BuildTabStatus(True)

        Case "tabOptions"
          BuildTabOptions(True)

        Case "tabNetCamTypes"
          BuildTabNetCamTypes(True)
          Me.pageCommands.Add("executefunction", "reDrawNetCamTypes()")

        Case "tabNetCamControls"
          BuildTabNetCamControls(0, True)
          Me.pageCommands.Add("executefunction", "reDrawNetCamControls()")

        Case "selNetCamType"
          Dim strValue As String = postData(postData("id"))
          Dim netcam_type As Integer = 0
          If IsNumeric(strValue) Then netcam_type = Integer.Parse(strValue)
          BuildTabNetCamControls(netcam_type, True)
          Me.pageCommands.Add("executefunction", "reDrawNetCamControls()")

        Case "tabNetCamDevices"
          BuildTabNetCamDevices(True)
          Me.pageCommands.Add("executefunction", "reDrawNetCamDevices()")

        Case "tabEventSnapshots", "selNetworkCamera"
          Dim netcam_id As String = postData("selNetworkCamera") & ""
          Dim event_id As String = String.Empty

          BuildTabEventSnapshots(netcam_id, event_id, True)

        Case "selNetCamEvent"
          Dim netcam_id As String = postData("selNetworkCamera") & ""
          Dim event_id As String = postData("selNetCamEvent") & ""

          BuildTabEventSnapshots(netcam_id, event_id, True)

        Case "btnDeleteEventSnapshots"
          Dim netcam_id As String = postData("selNetworkCamera") & ""
          Dim event_id As String = postData("selNetCamEvent") & ""

          hspi_plugin.PurgeSnapshotEvent(netcam_id, event_id)
          BuildTabEventSnapshots(netcam_id, "", True)

          PostMessage("The selected event snapshots were deleted.")

        Case "tabSnapshots"
          BuildTabCameras(gSnapshotsMaxWidth, True)

        Case "selSnapshotsWidth"
          Dim SnapShotWidth As String = postData(postData("id"))
          BuildTabCameras(SnapShotWidth, True)

        Case "selDeviceTypes"
          Dim JavaScriptSerializer As New JavaScriptSerializer()

          Using MyDataTable As DataTable = GetNetCamTypesFromDB()
            Dim ipOpts As New List(Of ipOpt)
            For Each r As DataRow In MyDataTable.Rows
              Dim ipOpt As New ipOpt
              ipOpt.value = r("netcam_type")
              ipOpt.label = r("netcam_name")
              ipOpts.Add(ipOpt)
            Next
            Return JavaScriptSerializer.Serialize(ipOpts)
          End Using

        Case "selSnapshotEventMax"
          Dim value As String = postData(postData("id"))
          SaveSetting("Options", "SnapshotEventMax", value)

          PostMessage("The maximum snapshot event option has been updated.")

        Case "selSnapshotsPerPage"
          Dim value As String = postData(postData("id"))
          SaveSetting("Options", "SnapshotsPerPage", value)

          PostMessage("The snapshot per page option has been updated.")

        Case "selSnapshotsMaxWidth"
          Dim value As String = postData(postData("id"))
          SaveSetting("Options", "SnapshotsMaxWidth", value)

          PostMessage("The snapshot maximum width option has been updated.")

        Case "selRefreshInterval"
          Dim value As String = postData(postData("id"))
          SaveSetting("Options", "SnapshotRefreshInterval", value)

          PostMessage("The snapshot refresh interval has been updated.")

        Case "selSendEmail"
          Dim strValue As String = postData(postData("id"))
          SaveSetting("EmailNotification", "EmailEnabled", strValue)
          BuildTabOptions(True)

          PostMessage("The E-mail Notification option has been updated.")

        Case "txtEmailRcptTo"
          Dim strValue As String = postData(postData("id"))
          SaveSetting("EmailNotification", "EmailRcptTo", strValue)

          PostMessage("The E-mail Notification option has been updated.")

        Case "txtEmailFrom"
          Dim strValue As String = postData(postData("id"))
          SaveSetting("EmailNotification", "EmailFrom", strValue)

          PostMessage("The E-mail Notification option has been updated.")

        Case "txtEmailSubject"
          Dim strValue As String = postData(postData("id"))
          SaveSetting("EmailNotification", "EmailSubject", strValue)

          PostMessage("The E-mail Notification option has been updated.")

        Case "btnSaveEmailBody", "txtEmailBody"
          Dim strValue As String = postData("txtEmailBody").Trim.Replace(vbCrLf, "~")
          SaveSetting("EmailNotification", "EmailBody", strValue)

          PostMessage("The E-mail Notification option has been updated.")

        Case "chkResetEmailBody"
          SaveSetting("EmailNotification", "EmailBody", EMAIL_BODY_TEMPLATE)
          BuildTabOptions(True)

          PostMessage("The E-mail Notification option has been updated.")

        Case "selArchiveEnabled"
          Dim strValue As String = postData(postData("id"))
          SaveSetting("Archive", "ArchiveEnabled", strValue)
          BuildTabOptions(True)

          PostMessage("The Event Archive option has been updated.")

        Case "selArchiveLimit"
          Dim strValue As String = postData(postData("id"))
          SaveSetting("Archive", "ArchiveLimit", strValue)

          PostMessage("The Event Archive limit has been updated.")

        Case "selArchiveFiles"
          Dim strValue As String = postData(postData("id"))
          SaveSetting("Archive", "ArchiveFiles", strValue)

          PostMessage("The Event Archive Files option has been updated.")

        Case "txtArchiveDirectory"
          Dim strValue As String = postData(postData("id"))
          SaveSetting("Archive", "ArchiveDir", strValue)

          PostMessage("The Event Archive Directory option has been updated.")

        Case "selFTPArchiveEnabled"
          Dim strValue As String = postData(postData("id"))
          SaveSetting("FTPArchive", "ArchiveEnabled", strValue)
          BuildTabOptions(True)

          PostMessage("The FTP Archive option has been updated.")

        Case "selFTPArchiveLimit"
          Dim strValue As String = postData(postData("id"))
          SaveSetting("FTPArchive", "ArchiveLimit", strValue)

          PostMessage("The FTP Archive Limit option has been updated.")

        Case "selFTPArchiveFiles"
          Dim strValue As String = postData(postData("id"))
          SaveSetting("FTPArchive", "ArchiveFiles", strValue)

          PostMessage("The FTP Archive Limit option has been updated.")

        Case "txtFTPArchiveDir"
          Dim strValue As String = postData(postData("id"))
          SaveSetting("FTPArchive", "ArchiveDir", strValue)

          PostMessage("The FTP Archive Directory option has been updated.")

        Case "txtFTPServerName"
          Dim strValue As String = postData(postData("id"))
          SaveSetting("FTPClient", "ServerName", strValue)

          PostMessage("The FTP Server Name has been updated.")

        Case "txtFTPUserName"
          Dim strValue As String = postData(postData("id"))
          SaveSetting("FTPClient", "UserName", strValue)

          PostMessage("The FTP User Name has been updated.")

        Case "txtFTPUserPass"
          Dim strValue As String = postData(postData("id"))
          SaveSetting("FTPClient", "UserPass", strValue)

          PostMessage("The FTP User Name has been updated.")

        Case "selVideoSnapshotCount"
          Dim strValue As String = postData(postData("id"))
          SaveSetting("ConvertToVideo", "EventSnapshotCount", strValue)

          PostMessage("The Video Snapshot option has been updated.")

        Case "selFoscamAutoDiscovery"
          Dim strValue As String = postData(postData("id"))
          SaveSetting("Options", "FoscamAutoDiscovery", strValue)

          PostMessage("The Foscam Auto Discovery option has been updated.  You must restart the plug-in for the setting to take affect.")

        Case "selLogLevel"
          gLogLevel = Int32.Parse(postData("selLogLevel"))
          hs.SaveINISetting("Options", "LogLevel", gLogLevel.ToString, gINIFile)

          PostMessage("The application logging level has been updated.")

        Case "WebPageAccess_Guest"

          Dim AUTH_ROLES As Integer = WEBUserRolesAuthorized()
          If postData("chkWebPageAccess_Guest") = "checked" Then
            AUTH_ROLES = AUTH_ROLES Or USER_GUEST
          Else
            AUTH_ROLES = AUTH_ROLES Xor USER_GUEST
          End If
          hs.SaveINISetting("WEBUsers", "AuthorizedRoles", AUTH_ROLES.ToString, gINIFile)

        Case "WebPageAccess_Normal"

          Dim AUTH_ROLES As Integer = WEBUserRolesAuthorized()
          If postData("chkWebPageAccess_Normal") = "checked" Then
            AUTH_ROLES = AUTH_ROLES Or USER_NORMAL
          Else
            AUTH_ROLES = AUTH_ROLES Xor USER_NORMAL
          End If
          hs.SaveINISetting("WEBUsers", "AuthorizedRoles", AUTH_ROLES.ToString, gINIFile)

        Case "WebPageAccess_Local"

          Dim AUTH_ROLES As Integer = WEBUserRolesAuthorized()
          If postData("chkWebPageAccess_Local") = "checked" Then
            AUTH_ROLES = AUTH_ROLES Or USER_LOCAL
          Else
            AUTH_ROLES = AUTH_ROLES Xor USER_LOCAL
          End If
          hs.SaveINISetting("WEBUsers", "AuthorizedRoles", AUTH_ROLES.ToString, gINIFile)

        Case "timer" ' This stops the timer and clears the message
          If TimerEnabled Then 'this handles the initial timer post that occurs immediately upon enabling the timer.
            TimerEnabled = False
          Else
            Me.pageCommands.Add("stoptimer", "")
            Me.divToUpdate.Add("divMessage", "&nbsp;")
          End If

      End Select

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "postBackProc")
    End Try

    Return MyBase.postBackProc(page, data, user, userRights)

  End Function

  Private Class ipOpt
    Public label As String = String.Empty
    Public value As String = String.Empty
  End Class

  ''' <summary>
  ''' Returns the Datatable Error JSON
  ''' </summary>
  ''' <param name="errorString"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function DatatableError(ByVal errorString As String) As String

    Try
      Return String.Format("{{ ""error"": ""{0}"" }}", errorString)
    Catch pEx As Exception
      Return String.Format("{{ ""error"": ""{0}"" }}", pEx.Message)
    End Try

  End Function

  ''' <summary>
  ''' Returns the Datatable Field Error JSON
  ''' </summary>
  ''' <param name="fieldName"></param>
  ''' <param name="fieldError"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function DatatableFieldError(fieldName As String, fieldError As String) As String

    Try
      Return String.Format("{{ ""fieldErrors"": [ {{""name"": ""{0}"",""status"": ""{1}""}} ] }}", fieldName, fieldError)
    Catch pEx As Exception
      Return String.Format("{{ ""fieldErrors"": [ {{""name"": ""{0}"",""status"": ""{1}""}} ] }}", fieldName, pEx.Message)
    End Try

  End Function

  ''' <summary>
  ''' Returns the Datatable Row JSON
  ''' </summary>
  ''' <param name="netcam_type"></param>
  ''' <param name="netcam_vendor"></param>
  ''' <param name="netcam_model"></param>
  ''' <param name="snapshot_path"></param>
  ''' <param name="videostream_path"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function DatatableRowNetCamType(ByVal netcam_type As String, _
                                          ByVal netcam_vendor As String, _
                                          ByVal netcam_model As String, _
                                          ByVal snapshot_path As String, _
                                          ByVal videostream_path As String) As String

    Try

      Dim sb As New StringBuilder
      sb.AppendLine("{")
      sb.AppendLine(" ""row"": { ")

      sb.AppendFormat(" ""{0}"": {1}, ", "DT_RowId", netcam_type)
      sb.AppendFormat(" ""{0}"": ""{1}"", ", "netcam_vendor", netcam_vendor)
      sb.AppendFormat(" ""{0}"": ""{1}"", ", "netcam_model", netcam_model)
      sb.AppendFormat(" ""{0}"": ""{1}"", ", "snapshot_path", snapshot_path)
      sb.AppendFormat(" ""{0}"": ""{1}"" ", "videostream_path", videostream_path)

      sb.AppendLine(" }")
      sb.AppendLine("}")

      Return sb.ToString

    Catch pEx As Exception
      Return "{ }"
    End Try

  End Function

  ''' <summary>
  ''' Returns the Datatable Row JSON
  ''' </summary>
  ''' <param name="control_id"></param>
  ''' <param name="control_name"></param>
  ''' <param name="control_url"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function DatatableRowNetCamControl(ByVal control_id As String, _
                                             ByVal control_name As String, _
                                             ByVal control_url As String) As String

    Try

      '(control_id, netcam_type, control_name, control_url

      Dim sb As New StringBuilder
      sb.AppendLine("{")
      sb.AppendLine(" ""row"": { ")

      sb.AppendFormat(" ""{0}"": {1}, ", "DT_RowId", control_id)
      sb.AppendFormat(" ""{0}"": ""{1}"", ", "control_name", control_name)
      sb.AppendFormat(" ""{0}"": ""{1}"" ", "control_url", control_url)

      sb.AppendLine(" }")
      sb.AppendLine("}")

      Return sb.ToString

    Catch pEx As Exception
      Return "{ }"
    End Try

  End Function

  ''' <summary>
  ''' Returns the Datatable Row JSON
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
  Private Function DatatableRowNetCamDevice(ByVal netcam_id As String, _
                                             ByVal netcam_name As String, _
                                             ByVal netcam_address As String,
                                             ByVal netcam_port As String, _
                                             ByVal netcam_type As String, _
                                             ByVal auth_user As String, _
                                             ByVal auth_pass As String) As String

    Try

      Dim sb As New StringBuilder
      sb.AppendLine("{")
      sb.AppendLine(" ""row"": { ")

      sb.AppendFormat(" ""{0}"": {1}, ", "DT_RowId", netcam_id)
      sb.AppendFormat(" ""{0}"": ""{1}"", ", "netcam_name", netcam_name)
      sb.AppendFormat(" ""{0}"": ""{1}"", ", "netcam_address", netcam_address)
      sb.AppendFormat(" ""{0}"": ""{1}"", ", "netcam_port", netcam_port)
      sb.AppendFormat(" ""{0}"": ""{1}"", ", "netcam_type", netcam_type)
      sb.AppendFormat(" ""{0}"": ""{1}"", ", "auth_user", auth_user)
      sb.AppendFormat(" ""{0}"": ""{1}"" ", "auth_pass", auth_pass)

      sb.AppendLine(" }")
      sb.AppendLine("}")

      Return sb.ToString

    Catch pEx As Exception
      Return "{ }"
    End Try

  End Function


#End Region

#Region "HSPI - Web Authorization"

  ''' <summary>
  ''' Returns the HTML Not Authorized web page
  ''' </summary>
  ''' <param name="LoggedInUser"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function WebUserNotUnauthorized(LoggedInUser As String) As String

    Try

      Dim sb As New StringBuilder

      sb.AppendLine("<table border='0' cellpadding='2' cellspacing='2' width='575px'>")
      sb.AppendLine("  <tr>")
      sb.AppendLine("   <td nowrap>")
      sb.AppendLine("     <h4>The Web Page You Were Trying To Access Is Restricted To Authorized Users ONLY</h4>")
      sb.AppendLine("   </td>")
      sb.AppendLine("  </tr>")
      sb.AppendLine("  <tr>")
      sb.AppendLine("   <td>")
      sb.AppendLine("     <p>This page is displayed if the credentials passed to the web server do not match the ")
      sb.AppendLine("      credentials required to access this web page.</p>")
      sb.AppendFormat("     <p>If you know the <b>{0}</b> user should have access,", LoggedInUser)
      sb.AppendFormat("      then ask your <b>HomeSeer Administrator</b> to check the <b>{0}</b> plug-in options", IFACE_NAME)
      sb.AppendFormat("      page to make sure the roles assigned to the <b>{0}</b> user allow access to this", LoggedInUser)
      sb.AppendLine("        web page.</p>")
      sb.AppendLine("  </td>")
      sb.AppendLine(" </tr>")
      sb.AppendLine(" </table>")

      Return sb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "WebUserNotUnauthorized")
      Return "error - " & Err.Description
    End Try

  End Function

#End Region

End Class