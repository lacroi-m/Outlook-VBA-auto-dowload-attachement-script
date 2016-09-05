
Public Sub saveAttachtoDisk(itm As Outlook.MailItem)

Dim objAtt As Outlook.Attachment
Dim saveFolder As String
Dim dateFormat
    dateFormat = Format(Now, "yyyy-mm-dd H-mm")

saveFolder = "C:\Users\ze\Documents\Attachements"

     For Each objAtt In itm.Attachments
        If StrComp(Left(objAtt.FileName, 16), "attachement.dat", vbTextCompare) = 0 Then
          objAtt.SaveAsFile saveFolder & "\" & dateFormat & objAtt.DisplayName
          Set objAtt = Nothing
        End If
    Next
End Sub
