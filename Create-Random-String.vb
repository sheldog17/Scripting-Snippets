Public Sub GiveMeARandomString()
  
  Dim BuildStr As New StringBuilder()
  BuildStr.Append(RandomString(4, True))
  BuildStr.Append(RandomInt(1000, 9999))
  BuildStr.Append(RandomString(4, False))
   
  Dim RandomString as String = ScrambleIt(BuildStr)
  
End Sub

Private Function RandomNumber(min As Integer, max As Integer) As Integer
   Dim random As New Random()
   Return random.Next(min, max)
End Function

Private Function RandomString(size As Integer, lowerCase As Boolean) As String
   Dim builder As New StringBuilder()
   Dim random As New Random()
   Dim ch As Char
   Dim i As Integer
   For i = 0 To size - 1
      ch = Convert.ToChar(Convert.ToInt32((26 * random.NextDouble() + 65)))
      builder.Append(ch)
   Next
   i If lowerCase Then
      Return builder.ToString().ToLower()
   End If
   Return builder.ToString()
End Function

Public Function ScrambleIt(ByVal phrase As String) As String
    Static rand As New Random()
    Return New String(phrase.ToLower.ToCharArray.OrderBy(Function(r) rand.Next).ToArray)
End Function 
