<div align="center">

## Download e\-mail attachments


</div>

### Description

This code enables you to download and send e-mail, which will automatically put your

attachments into a given directory.
 
### More Info
 
Mapisession control, mapimessages control, 2 command buttons and 1 text box.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Gemma Dobbins](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/gemma-dobbins.md)
**Level**          |Unknown
**User Rating**    |4.2 (161 globes from 38 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/gemma-dobbins-download-e-mail-attachments__1-4504/archive/master.zip)





### Source Code

```
Private Sub Command1_Click()
  MAPISession1.DownLoadMail = False
  MAPISession1.SignOn
  MAPIMessages1.SessionID = MAPISession1.SessionID
  MAPIMessages1.MsgIndex = -1
  MAPIMessages1.Compose
  MAPIMessages1.Send True
  MAPISession1.SignOff
End Sub
Private Sub Command2_Click()
  MAPISession1.DownLoadMail = True
  MAPISession1.SignOn
  MAPIMessages1.FetchUnreadOnly = True
  MAPIMessages1.SessionID = MAPISession1.SessionID
  MAPIMessages1.Fetch
  On Error Resume Next
  MAPIMessages1.AttachmentPathName = MAPIMessages1.AttachmentPathName '"c:\2000\" & MAPIMessages1.AttachmentName & "" 'vartype8 '& MAPIMessages1.AttachmentName & " '"
  Text1.Text = MAPIMessages1.MsgNoteText
  FileCopy MAPIMessages1.AttachmentPathName, ("c:\2000\" & MAPIMessages1.AttachmentName & "")
  MsgBox "File " & MAPIMessages1.AttachmentName & " sucessfully downloaded to C:\2000"
  MAPISession1.SignOff
End Sub
```

