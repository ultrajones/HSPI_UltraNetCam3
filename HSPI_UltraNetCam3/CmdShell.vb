' Dim MyCommand As New CmdShell
' Dim Results() As String = MyCommand.DoCommand("ping", "www.homeseer.com -n 1", "c:\windows\system32")
'
' For Each Line As String In Results 
'   Call hs.WriteLog("Debug", Line.ToString) 
' Next  
'
Public Class CmdShell

  '-------------------------------------------------------------------------------  
  ' Purpose:   Calls command line utility and returns results as String() 
  ' Input:     strProgram as String, strArgs as String, strWorkingDirectory as String  
  ' Output:    String() as String 
  '-------------------------------------------------------------------------------  
  Function DoCommand(ByVal strProgram As String, _
                     ByVal strArgs As String, _
                     ByVal strWorkingDirectory As String, _
                     Optional ByVal iWaitSeconds As Integer = 10) As Hashtable

    Dim Output As New Hashtable

    Output("StdOut") = ""
    Output("StdErr") = ""

    Try

      Using myproc As New Process

        With myproc
          .StartInfo.FileName = strProgram
          .StartInfo.Arguments = strArgs
          .StartInfo.RedirectStandardOutput = True
          .StartInfo.RedirectStandardError = True
          .StartInfo.UseShellExecute = False
          .StartInfo.CreateNoWindow = True
          .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
          .StartInfo.WorkingDirectory = strWorkingDirectory
          .Start()
        End With

        'Dim StdOut As StreamReader = myproc.StandardOutput
        'Dim strStdOut As String = StdOut.ReadToEnd()
        'Output("StdOut") = Regex.Replace(strStdOut, "\n+", " ")

        'Dim StdErr As StreamReader = myproc.StandardError
        'Dim strStdErr As String = StdErr.ReadToEnd()
        'Output("StdErr") = Regex.Replace(strStdErr, "\n+", " ")

        myproc.WaitForExit(1000 * iWaitSeconds)

      End Using

    Catch pEx As Exception
      '  
      ' Fatal error occured, so return the error within the string output  
      '  
      Output("StdErr") = pEx.Message
    End Try

    '  
    ' Parse the dig output  
    '  
    Return Output

  End Function

End Class
