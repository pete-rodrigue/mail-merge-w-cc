# mail-merge-w-cc
This small repo shows a couple ways to send mail merge emails with CC lines.


## Excel macro & a simple excel sheet.

The first way to do this is with an excel sheet and a simple macro. Please see the file in this repo called "send test emails.xslm."


```
' This is a short script that allows you to send emails in bulk to multiple recipents, with people CC'd and/or BCC'd.

' Sub is short for "subroutine" I think. This is the task we're going to run.

Sub Send_Bulk_Email()
    Dim ws As Worksheet                        ' Dim ws is VBA's way of creating an object called ws
    Set ws = ThisWorkbook.Sheets("Sheet1")     ' Set the ws object to be equal to the sheet we're using in our worksheet, called "Sheet1"
    
    Dim i As Integer
    Dim OA As Object                           ' Create objects called OA (for the outlook application) and msg (for the email message)
    Dim msg As Object
    
    Set OA = CreateObject("Outlook.application")
    
    Dim lastRow As Integer
    
    lastRow = Application.CountA(ws.Range("A:A"))  ' Find the last row in our sheet that has information
    
    ' loop through each row in the sheet and send an email!
    For i = 2 To lastRow
        Set msg = OA.createItem(0)
        msg.To = ws.Range("A" & i).Value
        msg.CC = ws.Range("B" & i).Value
        msg.BCC = ws.Range("C" & i).Value
        msg.Subject = ws.Range("D" & i).Value
        msg.Body = ws.Range("K" & i).Value
        
        msg.Send                                    ' msg.Send is the line that sends the email. If you want to manually click send,
                                                    ' replace msg.Send with msg.Display
        ws.Range("L" & i).Value = "Email Sent"      ' this puts a note in the sheet that we send the email
        Next i                                      ' then we go to the next row
    
End Sub
```
