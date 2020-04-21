VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Keep Running"
   ClientHeight    =   615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2415
   Icon            =   "keeprun.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   615
   ScaleWidth      =   2415
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.Timer dTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   240
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   840
      Top             =   120
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Detect As String
Dim Elapse As Double
Dim Tick
Dim Launch As String
Dim Reboot As String
Dim Start As String
Dim Delay As Integer
Dim Waiting As Boolean
Dim Tock

Private Sub CheckProcess()
'On Error Resume Next
  Dim cb As Long
  Dim cbNeeded As Long
  Dim NumElements As Long
  Dim ProcessIDs() As Long
  Dim cbNeeded2 As Long
  Dim NumElements2 As Long
  Dim Modules(1 To 200) As Long
  Dim lRet As Long
  Dim ModuleName As String
  Dim nSize As Long
  Dim hProcess As Long
  Dim i As Long
         
    'Get the array containing the process id's for each process object
    cb = 8
    cbNeeded = 96
    Do While cb <= cbNeeded
        cb = cb * 2
        ReDim ProcessIDs(cb / 4) As Long
        lRet = EnumProcesses(ProcessIDs(1), cb, cbNeeded)
    Loop
         
    NumElements = cbNeeded / 4
    foundit = False
    For i = 1 To NumElements
      'Get a handle to the Process
         hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, ProcessIDs(i))
      'Got a Process handle
         If hProcess <> 0 Then
           'Get an array of the module handles for the specified process
             lRet = EnumProcessModules(hProcess, Modules(1), 200, cbNeeded2)
             
           'If the Module Array is retrieved, Get the ModuleFileName
            If lRet <> 0 Then
                ModuleName = Space(MAX_PATH)
                nSize = 500
                lRet = GetModuleFileNameExA(hProcess, Modules(1), ModuleName, nSize)
                 
                If CBool(InStr(1, (Left(ModuleName, lRet)), m_strFilter, vbTextCompare)) Then
                    'AddListItem Left(ModuleName, lRet), ProcessIDs(i)
                    If LCase(Left(ModuleName, lRet)) Like LCase(Detect) = True Then
                        foundit = True
                    End If
                    
                End If
            End If
        End If
               
        'Close the handle to the process
        lRet = CloseHandle(hProcess)
    Next
    If foundit = False And Reboot = "False" Then Shell Launch, vbNormalNoFocus
    If foundit = False And Reboot = "True" Then
        blee = ShellExecute(GetDesktopWindow, "open", "shutdown", "-r -t 00", App.Path, 0)
        Elapse = 10
    End If
    
End Sub

Private Sub dTimer_Timer()
Tock = Tock + 1
If Tock > Delay Then
    dTimer.Enabled = False
    Waiting = False
End If
End Sub

Private Sub Form_Load()
Waiting = False
On Error Resume Next
Tick = 0
Delay = 0
inifound = Dir("keeprun.ini")
If inifound = "" Then
    Open "keeprun.ini" For Output As #1
    
        Print #1, "; Keep Running v" & App.Major & "." & App.Minor & " Configuration File"
        Print #1, " "
        Print #1, "; full path of executable to be started for the first time"
        Print #1, "Start="
        Print #1, " "
        Print #1, "; wait X seconds before checking for the first time"
        Print #1, "Delay=2"
        Print #1, " "
        Print #1, "; full path of executable to be checked"
        Print #1, "Detect="
        Print #1, " "
        Print #1, "; check the running processes list every X seconds"
        Print #1, "Interval=0.5"
        Print #1, " "
        Print #1, "; full path of executable to be re-launched"
        Print #1, "Launch="
        Print #1, " "
        Print #1, "; reboot if the detected process quits"
        Print #1, "Reboot=No"

    Close #1
    MsgBox "Created the keeprun.ini file, please configure before use.", vbOKOnly, "Error"
    Unload Me
    End
Else
    Open "keeprun.ini" For Input As #1
    Do While Not EOF(1)
    Input #1, temp
    inidata = Split(temp, "=")
        If UBound(inidata) > 0 Then
            'inidata(0) contains the identifier
            'inidata(1) contains the value, plus a comment
            bestdata = Split(inidata(1), "vbtab")
            'bestdata(0) contains the right thing
            
            Select Case inidata(0)
                Case "Detect"
                    Detect = bestdata(0)
                
                Case "Interval"
                    Elapse = bestdata(0)
                
                Case "Launch"
                    Launch = bestdata(0)
                
                Case "Reboot"
                    If LCase(bestdata(0)) = "yes" Then
                        Reboot = "True"
                    Else
                        Reboot = "False"
                    End If
                
                Case "Start"
                    Start = bestdata(0)
                
                Case "Delay"
                    Delay = bestdata(0)
                    
            End Select
        End If
    Loop
    Close
End If

'check to make sure every thing is right before starting the timer
If Launch <> "" Then
    If Detect <> "" Then
        If Elapse >= 0.5 Then
            If Launch Like "*keeprun.exe*" = False Then
                If Detect Like "*keeprun.exe*" = False Then
                    If Start <> "" Then
                        Shell Start, vbNormalNoFocus
                        Tock = 0
                        dTimer.Enabled = True
                        Waiting = True
                        Do While Waiting = True
                            DoEvents
                        Loop
                        dTimer.Enabled = False
                        Waiting = False
                    End If
                    Timer1.Enabled = True
                Else
                    MsgBox "Launch= or Detect= cannot be set to keeprun.exe", vbOKOnly, "Error in INI"
                    Unload Me
                End If
            Else
                MsgBox "Launch= or Detect= cannot be set to keeprun.exe", vbOKOnly, "Error in INI"
                Unload Me
            End If
        Else
            MsgBox "Interval= is too short, it must be 0.5 seconds or greater.", vbOKOnly, "Error in INI"
            Unload Me
        End If
    Else
        MsgBox "Detect= must be set to the full path of executable to be checked.", vbOKOnly, "Error in INI"
        Unload Me
    End If
Else
    MsgBox "Launch= must be set to the full path of executable to be re-launched.", vbOKOnly, "Error in INI"
    Unload Me
End If
End Sub

Private Sub Timer1_Timer()
Tick = Tick + 0.1
If Tick > Elapse Then
    Tick = 0
    CheckProcess
End If
End Sub
