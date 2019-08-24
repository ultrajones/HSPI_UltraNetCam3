Public Class DateComparer
  Implements System.Collections.IComparer

  Public Function Compare(ByVal info1 As Object, ByVal info2 As Object) As Integer Implements System.Collections.IComparer.Compare
    Dim FileInfo1 As System.IO.FileInfo = DirectCast(info1, System.IO.FileInfo)
    Dim FileInfo2 As System.IO.FileInfo = DirectCast(info2, System.IO.FileInfo)

    Dim Date1 As DateTime = FileInfo1.CreationTime
    Dim Date2 As DateTime = FileInfo2.CreationTime

    If Date1 > Date2 Then Return 0
    If Date1 < Date2 Then Return -1
    Return 0
  End Function

End Class