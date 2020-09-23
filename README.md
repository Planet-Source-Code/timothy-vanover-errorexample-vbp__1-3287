<div align="center">

## ErrorExample\.vbp


</div>

### Description

This is an example of how to use the intrinsic Visual basic App object to log errors to a file, which may later be helpful in debugging an application. The log file has the error number, error description, form name, sub or function name, and a date-time stamp when the error occurred. This also limits the error so that multiple loggings of the same error are not recorded in iterations.
 
### More Info
 
The public sub in the moderror.bas takes 4 parameters

LogErrors err.number, err.description, me.name , "MySubNameHere"

The app.StartLogging will not work in the development enviroment. The code must be in a compiled executable for it to work.

This writes a .log file which may be opened with any text editor for examination as to the general area of an error that could crash your program.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Timothy Vanover](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/timothy-vanover.md)
**Level**          |Unknown
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Debugging and Error Handling](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/debugging-and-error-handling__1-26.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/timothy-vanover-errorexample-vbp__1-3287/archive/master.zip)

### API Declarations

```
'* this code goes in a module for public use
'* Timothy A. Vanover
'* hdhunter@home.com
Public Sub LogErrors(lngErrNumber As Long, strErrDescription As String, _
strErrModule As String, strErrProcedure As String)
'*use to prevent recording multiple errors in loop
 Static LastErrorRecorded As String
 If CStr(Err.Number) & Err.Description & strErrModule & strErrProcedure = LastErrorRecorded Then
 Exit Sub
 Else
 App.StartLogging App.Path & "\" & App.EXEName & ".log", vbLogToFile
 App.LogEvent vbCrLf & _
 "Error Description: " & strErrDescription & vbCrLf & _
 "Error Number: " & lngErrNumber & vbCrLf & _
 "Module Name: " & strErrModule & vbCrLf & _
 "Procedure Name: " & strErrProcedure & vbCrLf & _
 "Version Number: " & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & _
 "Date: " & Format$(Now, "Short Date") & vbCrLf & _
 "Time: " & Format$(Now, "Long Time") & _
 vbCrLf & vbCrLf
 LastErrorRecorded = CStr(Err.Number) & Err.Description & strErrModule & strErrProcedure
 End If
End Sub
```


### Source Code

```
'* This must be compiled into an executable for the intrinsic
'* error logging to work
'* It will not work from the development enviroment.
'* paste this code on to a form, save and compile it for the demo
Private Sub Form_Load()
'*here is an example of a sub which I raise errors in for the demo
 ErrorDemoSub
 MsgBox "Errors Recorded in Error Log File"
 Unload Me
End Sub
Private Sub ErrorDemoSub()
 Dim i As Integer
 Dim ii As Integer
 On Error GoTo MyErrorLog
 'we'll simulate an error in a loop although we only log it one time
 For i = 1 To 20
 For ii = 1 To 5
  Err.Raise i
 Next ii
 Next i
 Exit Sub
MyErrorLog:
 LogErrors Err.Number, Err.Description, Me.Name, "ErrorDemoSub"
 Err.Clear
 Resume Next
End Sub
```

