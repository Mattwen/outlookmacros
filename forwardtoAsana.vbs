
' Opens a template from the network share and makes a new CAD Request
Sub OpenCADTemplate()
    Set temp = Application.CreateItemFromTemplate( _
    "P:\CAD Resources\Outlook Template\CADTEAMTemplateBlank.oft")
    temp.Display
    Set temp = Nothing
End Sub

' Forwards the current target email and sends to CADTeam
Public Sub ForwardwithCC()
    Dim myRecipient As Outlook.Recipient
    Dim objMsg, oMail As MailItem
    Dim str As String
    
    ' For a reply or reply all, replace forward  with Reply or ReplyAll
    Set objMsg = ActiveExplorer.Selection.Item(1).Forward
    
    ' Str for project requirements for CAD Request
    str = _
    "<BODY style=font-size:11pt; <p>Description:<br>" _
    & "Project Name:<br>" _
    & "Project #:<br>" _
    & "Task/Subtask:<br>" _
    & "Est.Time:<br>" _
    & "Due:<br></p></BODY>"
    
    Set myRecipient = objMsg.Recipients.Add("cadteam@rushingco.com")
    objMsg.HTMLBody = str & objMsg.HTMLBody
     
     
    If TypeName(ActiveExplorer.Selection.Item(1)) = "MailItem" Then
     Set oMail = ActiveExplorer.Selection.Item(1)
     
      
     On Error Resume Next
     
      objMsg.Display
     
    Else
     
    End If
     
    Set objMsg = Nothing
End Sub

Public Sub ForwardITRequest()
    Dim myRecipient As Outlook.Recipient
    Dim objMsg, oMail As MailItem
    Dim str As String
    
    ' For a reply or reply all, replace forward  with Reply or ReplyAll
    Set objMsg = ActiveExplorer.Selection.Item(1).Forward
    
    
    Set myRecipient = objMsg.Recipients.Add("AsanaIT@rushingco.com")
     
    If TypeName(ActiveExplorer.Selection.Item(1)) = "MailItem" Then
     Set oMail = ActiveExplorer.Selection.Item(1)
     
      
     On Error Resume Next
     
      objMsg.Display
     
    Else
     
    End If
     
    Set objMsg = Nothing
End Sub
