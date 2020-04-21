Attribute VB_Name = "modProcessAPI"
Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Boolean
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
  (ByVal hWnd As Long, ByVal lpOperation As String, _
  ByVal lpFile As String, ByVal lpParameters As String, _
  ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Declare Function GetDesktopWindow Lib "user32" () As Long

Public Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, _
                                                        lppe As PROCESSENTRY32) As Long
Public Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, _
                                                        lppe As PROCESSENTRY32) As Long
Public Declare Function CloseHandle Lib "Kernel32.dll" (ByVal Handle As Long) As Long
Public Declare Function OpenProcess Lib "Kernel32.dll" (ByVal dwDesiredAccessas As Long, _
                                                        ByVal bInheritHandle As Long, _
                                                        ByVal dwProcId As Long) As Long
Public Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, _
                                                        ByVal cb As Long, _
                                                        ByRef cbNeeded As Long) As Long
Public Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, _
                                                        ByVal hModule As Long, _
                                                        ByVal ModuleName As String, _
                                                        ByVal nSize As Long) As Long
Public Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, _
                                                        ByRef lphModule As Long, _
                                                        ByVal cb As Long, _
                                                        ByRef cbNeeded As Long) As Long
Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, _
                                                        ByVal th32ProcessID As Long) As Long
Public Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, _
                                                        ByVal uExitCode As Long) As Long
Public Declare Function GetLastError Lib "kernel32" () As Long
Public Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Public Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long           ' This process
    th32DefaultHeapID As Long
    th32ModuleID As Long            ' Associated exe
    cntThreads As Long
    th32ParentProcessID As Long     ' This process's parent process
    pcPriClassBase As Long          ' Base priority of process threads
    dwFlags As Long
    szExeFile As String * 260       ' MAX_PATH
End Type

Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Public Type DrewProcess
    module As String
    id As Long
End Type

Public Const PROCESS_QUERY_INFORMATION = 1024
Public Const PROCESS_VM_READ = 16
Public Const MAX_PATH = 260

'STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const SYNCHRONIZE = &H100000

Public Const PROCESS_ALL_ACCESS = &H1F0FFF
Public Const TH32CS_SNAPPROCESS = &H2&
Public Const hNull = 0

'Used to Get the Error Message
Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Const LANG_NEUTRAL = &H0
Const SUBLANG_DEFAULT = &H1

'Used to determine what OS Version
Public Const WINNT As Integer = 2
Public Const WIN98 As Integer = 1

Public Function getVersion() As Integer
  Dim udtOSInfo As OSVERSIONINFO
  Dim intRetVal As Integer
         
  'Initialize the type's buffer sizes
    With udtOSInfo
        .dwOSVersionInfoSize = 148
        .szCSDVersion = Space$(128)
    End With
    
  'Make an API Call to Retrieve the OSVersion info
    intRetVal = GetVersionExA(udtOSInfo)
  
  'Set the return value
    getVersion = udtOSInfo.dwPlatformId
End Function

Public Sub KillProcessById(p_lngProcessId As Long)
  Dim lnghProcess As Long
  Dim lngReturn As Long
    
    lnghProcess = OpenProcess(1&, -1&, p_lngProcessId)
    lngReturn = TerminateProcess(lnghProcess, 0&)
    
    If lngReturn = 0 Then
        RetrieveError
    End If
End Sub

Private Sub RetrieveError()
  Dim strBuffer As String
    
    'Create a string buffer
    strBuffer = Space(200)

    FormatMessage FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, GetLastError, LANG_NEUTRAL, strBuffer, 200, ByVal 0&
End Sub

Public Function DayOfWeek(ByVal Day As Integer, ByVal Month As Integer, ByVal Year As Integer) As String
    Dim w As Integer, wQuotient, wRemainder, int6

    If Month = 1 Then
        Month = 13
        Year = Year - 1
    ElseIf Month = 2 Then
        Month = 14
        Year = Year - 1
    End If
    
    int6 = 0.6 * (Month + 1)
    int6 = Int(int6)
    w = Day + 2 * Month + int6 + Year + Int(Year / 4) - Int(Year / 100) + Int(Year / 400) + 2
    wQuotient = Int(w / 7)
    DayOfWeek = DayString(w - (wQuotient * 7))
End Function

Public Function DayOfYear(ByVal Day As Integer, ByVal Month As Integer, ByVal LeapYear As Boolean) As Integer
    Dim i As Integer, fDay As Integer

    For i = 1 To Month - 1
        fDay = fDay + DaysInMonth(i, LeapYear)
    Next
    fDay = fDay + Day
    DayOfYear = fDay
End Function

Public Function DaysBetween(ByVal startDay As Integer, ByVal startMonth As Integer, ByVal startYear As Integer, ByVal endDay As Integer, ByVal endMonth As Integer, ByVal endYear As Integer) As Long
    Dim startIsLeap As Boolean, endIsLeap As Boolean
    Dim daysToEnd As Integer, fDays As Integer
    startIsLeap = IsLeapYear(startYear)
    endIsLeap = IsLeapYear(endYear)
    startDay = DayOfYear(startDay, startMonth, startIsLeap)
    endDay = DayOfYear(endDay, endMonth, endIsLeap)

    If startYear = endYear Then
        DaysBetween = endDay - startDay
        Exit Function
    End If
    daysToEnd = DaysInYear(startYear) - startDay

    For i = startYear + 1 To endYear - 1
        fDays = fDays + DaysInYear(i)
    Next
    fDays = fDays + daysToEnd + endDay
    DaysBetween = fDays
End Function

Public Function DaysInMonth(ByVal Month As Integer, ByVal LeapYear As Boolean) As Integer

    Select Case Month
        Case 1, 3, 5, 7, 8, 10, 12: DaysInMonth = 31
        Case 2

        If LeapYear Then
            DaysInMonth = 29
        Else
            DaysInMonth = 28
        End If
        Case 4, 6, 9, 11: DaysInMonth = 30
    End Select
End Function

Public Function DaysInYear(ByVal Year As Integer) As Integer

    If IsLeapYear(Year) Then
        DaysInYear = 366
    Else
        DaysInYear = 365
    End If
End Function

Private Function DayString(ByVal Weekday As Integer)

    Select Case Weekday
        Case 0: DayString = "Saturday"
        Case 1: DayString = "Sunday"
        Case 2: DayString = "Monday"
        Case 3: DayString = "Tuesday"
        Case 4: DayString = "Wednesday"
        Case 5: DayString = "Thursday"
        Case 6: DayString = "Friday"
    End Select
End Function

Public Function IsLeapYear(ByVal Year As Integer) As Boolean

    If Year Mod 4 = 0 Then
        IsLeapYear = True


        If Year Mod 100 = 0 And Year Mod 400 <> 0 Then
            IsLeapYear = False
        End If
    End If
End Function
