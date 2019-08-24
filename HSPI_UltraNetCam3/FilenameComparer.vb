Public Class FilenameComparer
  Implements System.Collections.IComparer

  Public Function Compare(ByVal info1 As Object, ByVal info2 As Object) As Integer Implements System.Collections.IComparer.Compare
    Dim FileInfo1 As System.IO.FileInfo = DirectCast(info1, System.IO.FileInfo)
    Dim FileInfo2 As System.IO.FileInfo = DirectCast(info2, System.IO.FileInfo)

    Dim File1 As String = FileInfo1.FullName
    Dim File2 As String = FileInfo2.FullName

    If String.Compare(File1, File2) > 0 Then Return 0
    If String.Compare(File1, File2) < 0 Then Return -1
    Return 0
  End Function

End Class