<div align="center">

## LaunchAppSynchronous


</div>

### Description

Unlike the Shell command in VB which launches an application

asynchronous, this will launch the program synchronous.

What that means is that the shell execute command will launch

an application but not wait for it to execute before processing

the next line of code. This code will launch a program then

wait until the executable has terminated before executing the

next line of code.
 
### More Info
 
INPUT: The executables full path and name.

RETURN: True upon termination if successful, false if not.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Darrell Sparti, MCSD](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/darrell-sparti-mcsd.md)
**Level**          |Unknown
**User Rating**    |4.6 (78 globes from 17 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/darrell-sparti-mcsd-launchappsynchronous__1-3716/archive/master.zip)

### API Declarations

```
Private Const INFINITE = &HFFFFFFFF
Private Const NORMAL_PRIORITY_CLASS = &H20
Private Const WAIT_TIMEOUT = &H102&
'
Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type
'
Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type
'
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'
Private Declare Function WaitForInputIdle Lib "user32" (ByVal hProcess As Long, ByVal dwMilliseconds As Long) As Long
'
Private Declare Function CreateProcessByNum Lib "kernel32" Alias "CreateProcessA" _
                        (ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes _
                        As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags _
                        As Long, lpEnvironment As Any, ByVal lpCurrentDirectory As String, lpStartupInfo As _
                        STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
'
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
'
```


### Source Code

```
Option Explicit
'
'Unlike the Shell command in VB which launches an application
'asynchronous, this will launch the program synchronous.
'What that means is that the shell execute command will launch
'an application but not wait for it to execute before processing
'the next line of code. This code will launch a program then
'wait until the executable has terminated before executing the
'next line of code. Works great for launching DOS exe's such
'as batch files, reindexing old databases, and other executables
'which must perform their task before your code continues.
'Some versions don't work in Windows NT because of the added
'security but this version does work in Windows NT.
'I realize there are more elegant and sophisticated ways to do
'the same thing but this one works fine for what I needed in a
'professional application I was working on. I must credit Dan
'Appleman's Programmer's Guide To The Win32 API for this code.
'I also strongly suggest that anyone interested in understanding
'more about these kind of techniques, read his book. In fact,
'I recommend all of Dan Appleman's books when you are ready to
'go from novice to professional programmer.
'I appreciate your comments but please do your homework first!
Public Function LaunchAppSynchronous(strExecutablePathAndName As String) As Boolean
  'Launches an executable by starting it's process
  'then waits for the execution to complete.
  'INPUT: The executables full path and name.
  'RETURN: True upon termination if successful, false if not.
  Dim lngResponse As Long
  Dim typStartUpInfo As STARTUPINFO
  Dim typProcessInfo As PROCESS_INFORMATION
  LaunchAppSynchronous = False
  With typStartUpInfo
   .cb = Len(typStartUpInfo)
   .lpReserved = vbNullString
   .lpDesktop = vbNullString
   .lpTitle = vbNullString
   .dwFlags = 0
  End With
  'Launch the application by creating a new process
  lngResponse = CreateProcessByNum(strExecutablePathAndName, vbNullString, 0, 0, True, NORMAL_PRIORITY_CLASS, ByVal 0&, vbNullString, typStartUpInfo, typProcessInfo)
  If lngResponse Then
   'Wait for the application to terminate before moving on
   Call WaitForTermination(typProcessInfo)
   LaunchAppSynchronous = True
  Else
   LaunchAppSynchronous = False
  End If
End Function
Private Sub WaitForTermination(typProcessInfo As PROCESS_INFORMATION)
  'This wait routine allows other application events
  'to be processed while waiting for the process to
  'complete.
  Dim lngResponse As Long
  'Let the process initialize
  Call WaitForInputIdle(typProcessInfo.hProcess, INFINITE)
  'We don't need the thread handle so get rid of it
  Call CloseHandle(typProcessInfo.hThread)
  'Wait for the application to end
  Do
   lngResponse = WaitForSingleObject(typProcessInfo.hProcess, 0)
   If lngResponse <> WAIT_TIMEOUT Then
     'No timeout, app is terminated
     Exit Do
   End If
   DoEvents
  Loop While True
  'Kill the last handle of the process
  Call CloseHandle(typProcessInfo.hProcess)
End Sub
```

