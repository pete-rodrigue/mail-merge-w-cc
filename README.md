# mail-merge-w-cc
This small repo shows a couple ways to send mail merge emails with CC lines, attachments, and other fancy little bits.


## Excel macro & a simple excel sheet.

The first way to do this is with an excel sheet and a simple macro. Thanks to youtube user "Automate Data" for providing a thourogh walkthrough about how to do this. Here's a link to their channel: https://www.youtube.com/watch?v=2eHxTRisCVM 

Please see the file in this repo called "send test emails.xslm." You basically fill out that excel sheet with the information you need (who the emails are going to, any CC and BCC lines, the message, etc). Then hit ALT+F11 to open the macro editor, and run the macro to send the emails. Here's what the sheet should look like, roughly:

![image](https://github.com/pete-rodrigue/mail-merge-w-cc/assets/8962291/a508295e-bc3b-4d6d-bf02-6e116da5bedb)



And here is what the script should look like:

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


There are some disadvantages to doing things this way:

* You can't send people attachments.
* You can't easily send people HTML emails with pictures, etc.

But we can use Power Automate to accomplish those fancier things in mail merge. See next section.

## Sending bulk emails with CC lines, etc, using Power Automate

Thanks to youtube user "Efficiency 365 by Dr Nitin" for providing a thorough walkthrough on this. Their video is here: https://www.youtube.com/watch?v=Ij6RRDkipRI

First, you'll want to create a folder that has any attachments you want to send in it, along with an excel file called "contacts":

![image](https://github.com/pete-rodrigue/mail-merge-w-cc/assets/8962291/02d9a359-5834-4555-943f-d263bb2abeba)

See the folder in this repo called "send bulk emails w power automate". That has a template of the excel file you'll need and a dummy PDF file. Note that you'll need to make sure that the "settings" and "contacts" rows/cells in the excel workbook need to be saved as tables with specific names.

Then go to Power Automate and click "create instant cloud flow." Click "manually trigger a cloud flow." 

Now, on this page:

![image](https://github.com/pete-rodrigue/mail-merge-w-cc/assets/8962291/d9a553b2-ecae-46a0-a9ef-a905721f34ef)

Click "New step." From here, we're going to add steps to send our emails. See the Power Automate flow in this repo for details, called "SendbulkemailswithCClinesandattachments_20230811153422.zip." You can import that zip file as a flow into Power Automate. You can hover over each action/item to see where the data element is being pulled from. You may need to edit a few of those items/data elements.

If the flow works, you should be able to get emails like this:

![image](https://github.com/pete-rodrigue/mail-merge-w-cc/assets/8962291/3a3e88e5-5bc7-405a-bdf5-f5d87bfd7d13)
