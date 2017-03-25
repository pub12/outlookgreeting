'Copy and paste this code into outlook

Option Explicit
Private WithEvents oExpl As Explorer
Private WithEvents oItem As MailItem
Private bDiscardEvents As Boolean
 
Private Const strGreeting As String = "Hi "
Private Const strSignoff As String = "Thanks, Pubs."
 
 
Dim oResponse As MailItem
 
Private Sub Application_Startup()
   Set oExpl = Application.ActiveExplorer
   bDiscardEvents = False
End Sub
 
Private Sub oExpl_SelectionChange()
   On Error Resume Next
   Set oItem = oExpl.Selection.Item(1)
End Sub
 
Private Function getSenderName(senderName As String)
    'Assumes naming convention of "<Lastname>, <Firstname> <with optional middlename>"
    Dim strSplit As Variant
   
    strSplit = Split(senderName, ",")
    If UBound(strSplit) > 0 Then
        getSenderName = strSplit(1)
       
        strSplit = Split(getSenderName, " ") 'if there is optional middle name, just get first name
        If UBound(strSplit) > 0 Then
            getSenderName = strSplit(1)
        End If
    Else
        getSenderName = senderName
    End If
End Function
 
Private Sub processReply()
    Dim strText As String
    Dim olInspector As Outlook.Inspector
    Dim olDocument As Word.Document
    Dim olSelection As Word.Selection
    Dim oResponse As MailItem
 
 
    strText = strGreeting
   
    Set oResponse = oItem.Reply
    oResponse.Display
   
    Set olInspector = Application.ActiveInspector()
    Set olDocument = olInspector.WordEditor
    Set olSelection = olDocument.Application.Selection
 
    olSelection.InsertBefore strText & getSenderName(oItem.senderName) & "," & vbNewLine & vbNewLine & vbNewLine & vbNewLine & strSignoff
   
    Set oItem = Nothing
   
End Sub

Private Sub oItem_Reply(ByVal Response As Object, Cancel As Boolean)
    processReply
       
    Cancel = True
End Sub

Private Sub oItem_ReplyAll(ByVal Response As Object, Cancel As Boolean)
    processReply
    Cancel = True
End Sub
