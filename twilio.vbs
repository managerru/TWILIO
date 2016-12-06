Option Explicit
Dim fso, f1, ts, ts1, fso1, Soobshenie, SiteAddr1, cur_time, Account_SID, Auth_Token, Messaging_SID, NumberTo, NumberFrom, DataToSend
  Const ForReading = 1
  Const ForWriting = 2


Main

Sub Die()

End Sub



Sub Main
On Error Resume Next

'Twilio credentials
Account_SID = "DSDSDS61bcb984a8be2ed03f6ffa4cc" 
Auth_Token = "787788575ae5008eebe5a97cb66730be2"
Messaging_SID = "dfgdfgdf48e0bacf9735f29d98c01ff3e40"
NumberFrom = "+12044004069"

Set fso = CreateObject("Scripting.FileSystemObject")
Set ts = fso.OpenTextFile("twilionumbers.txt", ForReading)

SiteAddr1 = "https://api.twilio.com/2010-04-01/Accounts/"&Account_SID&"/Messages.json"
Soobshenie = "Test message"

Do While Not ts.AtEndOfStream
	NumberTo = ts.ReadLine
	Submit 
Loop

  ts.Close
  ts1.Close
  MsgBox "SMS Sending complete"
WScript.Quit()
End Sub

Sub Submit() 
    Dim XMLHTTP 
    Set XMLHTTP = CreateObject("MSXML2.XMLHTTP") 
    XMLHTTP.Open "POST", SiteAddr1, False , Account_SID, Auth_Token  
    XMLHTTP.setrequestheader "Content-type", "application/x-www-form-urlencoded"
    DataToSend = "From="&NumberFrom&"&To="&NumberTo&"&MessagingServiceSid="&Messaging_SID&"&Body="&Soobshenie
    XMLHTTP.Send DataToSend 
    cur_time = time & " --- " & date
    Set fso1 = CreateObject("Scripting.FileSystemObject")
    Set ts1 = fso1.OpenTextFile("twilioresult.txt", 8, True)
    ts1.WriteLine("--------------------------------------------------------------------------------------")
    ts1.WriteLine(cur_time)		
    ts1.WriteLine("--------------------------------------------------------------------------------------")
    ts1.WriteLine(XMLHTTP.Status)
    ts1.WriteLine(XMLHTTP.responsetext)
    ts1.Close
    Set fso1 = Nothing
    Set ts1 = Nothing
End Sub 
