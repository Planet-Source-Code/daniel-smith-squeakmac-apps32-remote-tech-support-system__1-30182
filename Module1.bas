Attribute VB_Name = "Module1"
Public Declare Function Sleep Lib "kernel32" (ms As Long)
 
Public MonCount As Integer

Public Function FileExist(ByVal FileName As String) As Boolean
'Determines if a file exists
On Error Resume Next
If Dir(FileName, vbSystem + vbHidden) = "" Then
    FileExist = False
Else
    FileExist = True
End If
End Function

Public Function EvalData(sIncoming As String, ParseOption As Integer) As String
Dim i As Integer
  
Select Case ParseOption
    Case 1
        For i = 1 To Len(sIncoming)
            If Mid(sIncoming, i, 1) = "," Then
                EvalData = Left(sIncoming, i - 1)
                Exit Function
            End If
        Next
    Case 2
        For i = 1 To Len(sIncoming)
            If Mid(sIncoming, i, 1) = "," Then
              EvalData = Right(sIncoming, Len(sIncoming) - i)
              Exit Function
            End If
        Next
End Select
End Function
