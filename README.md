<div align="center">

## Send SMS message via Http for free


</div>

### Description

Sends an SMS message to a cell phone for free. It makes use of the ServerXMLHTTP object contained in msxml3.dll. Uses the free German Web service www.billiger-telefonieren.de. The cookie checks of the site are circumvented by doing the cookie

handling explicitely. Therefore this code should work even server-side!

Please note that the site still puts some requirement on the send message. For example messages with subjects like "test" are rejected.

And: you can't send more than a certain number of messages to the the same number.

For the most recent updates please visit my homepage.

New information (30-Oct-02): I uploaded my third generation of SMS code under "SMS via HTTP - third generation".
 
### More Info
 
1. Message text (up to 160 chars)

2. Phone number (e.g. +49171xxxxxx)

I used it with Microsoft msxml3.dll (the released version). Download it from http://msdn.microsoft.com/xml/default.asp.

Comes back with a success or a failure message depending on the HTML that the site returns.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Klemens Schmid](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/klemens-schmid.md)
**Level**          |Beginner
**User Rating**    |4.7 (47 globes from 10 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/klemens-schmid-send-sms-message-via-http-for-free__1-5746/archive/master.zip)





### Source Code

```
'Author
' mailto:klemens.schmid@gmx.de, http://www.schmidks.de
'Description
' This code fires off an SMS message to the given phone number
' It makes use of the German service "www.billiger-telefonieren.de"
' The cookie checks of the site are circumvented by doing the cookie
' handling explicitely. Therefore this code should work even server-side!
' Please note that the site still puts some requirement on the send
' message. For example messages with subjects like "test" are rejected.
' And: you can't send more than a certain number of messages to the
' the same number.
'Prerequisites
' The posting is done thru the ServerXMLHTTP object which is contained
' in the Microsoft XML object msxml3.dll. Install this from
' http://msdn.microsoft.com/xml/default.asp.
Option Explicit
Public Sub SendSMS()
Dim strText As String
Dim strPhoneNo As String
Dim strCookie As String
Dim oHttp As ServerXMLHTTP
'make use of the XMLHTTPRequest object contained in msxml.dll
Set oHttp = CreateObject("msxml2.serverXMLHTTP")
'enter your data
strText = InputBox("Text:", "Send Text via SMS", "vbsms:")
strPhoneNo = InputBox("Phone Number:", "Send Text via SMS")
'fire of an http request to request for a cookie
oHttp.open "GET", "http://www.billiger-telefonieren.de/sms/send.php3?action=accept", False
oHttp.send
strCookie = oHttp.getResponseHeader("set-cookie")
strCookie = Left$(strCookie, InStr(strCookie, ";") - 1)
'better check the feedback
Debug.Print oHttp.responseText
'do the actual send
oHttp.open "POST", "http://www.billiger-telefonieren.de/sms/send.php3", False
oHttp.setRequestHeader "Cookie", strCookie
'we need to do it a second time due to KB article Q234486.
oHttp.setRequestHeader "Cookie", strCookie
oHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
oHttp.send "action=send&number=" & strPhoneNo & "&email=&message=" & strText
Debug.Print oHttp.responseText
If InStr(oHttp.responseText, "erfolgreich eine Nachricht an die") Then
 MsgBox "Message has been sent successfully", vbInformation
Else
 MsgBox "Service refused to send the message", vbCritical
End If
End Sub
```

