Attribute VB_Name = "Module1"
Public Function MakeLogFile(ByVal ErrorOccurredPosition As String) As String
Dim sFileName As String
Dim iFileNo As Integer
    If Err.Number <> 0 Then
        sFileName = App.Path & "\ErrorLogFile.txt"
        iFileNo = FreeFile
            Open sFileName For Append As #iFileNo
            Print #iFileNo, "Err Number :" & Err.Number
            Print #iFileNo, "Err Description :" & Err.Description
            Print #iFileNo, "Date :" & Date
            Print #iFileNo, "Time :" & Time
            Print #iFileNo, "ErrorOccurredPosition :" & ErrorOccurredPosition
            Close #iFileNo
    End If
End Function
