<div align="center">

## Display the File Open Dialog


</div>

### Description

Display the windows common file open dialog to select files. File filters, multi-select, starting directory. Requires no ActiveX or OXC controls on the form.
 
### More Info
 
'Place code inside any module, form or class. (Will work in sub and function declarations)

string / array = files selected


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Melxium CG](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/melxium-cg.md)
**Level**          |Beginner
**User Rating**    |4.6 (23 globes from 5 users)
**Compatibility**  |VB\.NET
**Category**       |[System Services/ Functions](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/system-services-functions__10-23.md)
**World**          |[\.Net \(C\#, VB\.net\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/net-c-vb-net.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/melxium-cg-display-the-file-open-dialog__10-103/archive/master.zip)

### API Declarations

```
SoftTech InterCorp:
Please visit our corporate home page: http://www.stintercorp.com/
```


### Source Code

```

    'Create Dialog Fuction
    Dim fOpenDialog As New OpenFileDialog()
    'Variable used for retrival of selected files
    Dim sFile As String = ""
    'Start of width statement
    With fOpenDialog
      'Set Filter for the selected files
      .Filter = "Supported Images|*.bmp;*.jpg|Bitmap Images|*.bmp|JPEG Images|*.jpg|All Files|*.*"
      'Set the title of the selected file
      .Title = "Select Image to Open"
      'Enable Multi-Select
      .Multiselect = True
      'make sure all selected files exists
      .CheckFileExists = True
      'Set initial directory to the users desktop
      .InitialDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
      'show the open dialog
      .ShowDialog()
      'get each of the selected file and display it in a msgbox
      For Each sFile In .FileNames
        'display msgbox the current filepath
        MsgBox("File: " & sFile)
      Next
      'clear the dialog from resorces
      .Dispose()
    End With
```

