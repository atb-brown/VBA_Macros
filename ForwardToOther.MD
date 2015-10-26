### This macro assumes that the email was mistakenly sent to the user, but the forwardStamp can be edited as the targetPerson and targetEmail can


```
Sub Forward_OtherAustinBrown()

    ' Current application
    Dim outlookObject As Outlook.Application
    Dim nameSpaceObject As Outlook.NameSpace
    ' Editing the message
    Dim documentObject As Word.Document
    Dim selectionObject As Word.Selection
    Dim inspectorObject As Outlook.Inspector
    ' Forwarding
    Dim forwardMail As Outlook.MailItem
    Dim signatureMail As Object
    ' Strings used for forwarding the email
    Dim forwardStamp As String
    Dim targetEmail As String
    Dim signature As String
    
    ' target recipient variables
    targetPerson = "FirstName LastName"
    targetEmail = "firstname.lastname@domain.com"
    
    ' message being added to the forwarded email
    forwardStamp = "Hello " & targetPerson & ", I think this email was meant for you." _
                    & vbCrLf & vbCrLf & vbCrLf & "Forwarded to you from:"
    
    ' Get the current application
    Set outlookObject = Application
    
    ' Make a new email for the default signature
    Set signatureMail = CreateObject("Outlook.Application").CreateItem(0)
    
    Set nameSpaceObject = outlookObject.Session
    
    If Not outlookObject.ActiveInspector Is Nothing Then
        If outlookObject.ActiveInspector.EditorType = olEditorWord Then
            ' Setup the forward email
            Set forwardMail = outlookObject.ActiveInspector.CurrentItem.Forward
            Set inspectorObject = forwardMail.GetInspector
            
            ' Get the signature from signatureMail
            With signatureMail
            .Display ' Open the email
            End With
                signature = signatureMail.Body
            signatureMail.Close olDiscard ' Close the email so it's no longer in the user's way
            ' Remove the blank lines above the signature
            signature = Right(signature, Len(signature) - 4)
            
            ' Get the received email as a Word document
            Set documentObject = inspectorObject.WordEditor
            Set selectionObject = documentObject.Windows(1).Selection
            
            ' Put cursor at beginning and append the forwardStamp
            selectionObject.Move wdStory, -1
            documentObject.Characters(1).InsertBefore _
                forwardStamp & vbCrLf & signature
            selectionObject.Move wdParagraph, 1
            
            ' Display the new message and add the target Recipient
            forwardMail.Display
            forwardMail.Recipients.Add targetEmail
        Else
            MsgBox "The editor isn't a WordEditor type."
        End If
        ' Reset objects
        Set outlookObject = Nothing
        Set nameSpaceObject = Nothing
        Set documentObject = Nothing
        Set inspectorObject = Nothing
        Set selectionObject = Nothing
        Set forwardMail = Nothing
    Else
        MsgBox "There is no active inspector."
    End If
    
End Sub

```
