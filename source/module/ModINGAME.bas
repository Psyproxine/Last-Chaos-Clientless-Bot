Attribute VB_Name = "ModINGAME"
Option Explicit
Public Const HIGH_PRIORITY_CLASS = &H80
Public Const IDLE_PRIORITY_CLASS = &H40
Public Const NORMAL_PRIORITY_CLASS = &H20
Public Const REALTIME_PRIORITY_CLASS = &H100

Public Const PROCESS_ALL_ACCESS = &H1F0FFF

Declare Function InjectLibrary Lib "madCodeHookLib.dll" (ByVal id As Long, ByVal tstr As String) As Long

Public Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long

Public ProcessName As String
Public Declare Function OpenProcess Lib "kernel32" _
(ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
ByVal dwProcessId As Long) As Long

Private Declare Function CreateToolhelp32Snapshot Lib "kernel32.dll" _
  (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long

Private Declare Function Process32First Lib "kernel32.dll" _
  (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long

Private Declare Function Process32Next Lib "kernel32.dll" _
  (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long

Private Declare Function CloseHandle Lib "kernel32.dll" _
  (ByVal hObject As Long) As Long

Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" _
  (ByVal lpString As String) As Long

Private Type PROCESSENTRY32
  dwSize As Long
  cntUsage As Long
  th32ProcessID As Long
  th32DefaultHeapID As Long
  th32ModuleID As Long
  cntThreads As Long
  th32ParentProcessID As Long
  pcPriClassBase As Long
  dwFlags As Long
  szExeFile As String * 260
End Type

Private Const TH32CS_INHERIT = &H80000000
Private Const TH32CS_SNAPALL = &HF
Private Const TH32CS_SNAPHEAPLIST = &H1
Private Const TH32CS_SNAPPROCESS = &H2
Private Const TH32CS_SNAPTHREAD = &H4
Private Const TH32CS_SNAPMODULE = &H8

Public Function GetProcessByName(name As String) As Long
 Dim test As String
  Dim RetVal As Long
  Dim hSnap As Long
  Dim PInfo As PROCESSENTRY32

  ' Snapshot vom gesamten System erstellen
  hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
  If hSnap = -1 Then
    MsgBox "Der System-Snapshot konnte nicht erstellt werden.", _
    vbInformation, "Fehler"
    Exit Function
  End If

  PInfo.dwSize = Len(PInfo)
  RetVal = Process32First(hSnap, PInfo) ' ersten Prozess ermitteln

  Do Until RetVal = 0
    With PInfo
      .szExeFile = Trim$(Left$(.szExeFile, lstrlen(.szExeFile))) _
      ' VBNullChar abtrennen
      test = LCase(Trim(Mid$(.szExeFile, InStrRev(.szExeFile, "\") + 1)))
      If test = name Then
        GetProcessByName = .th32ProcessID
        Exit Function
      End If
    End With
    
    RetVal = Process32Next(hSnap, PInfo) ' nächsten Prozess ermitteln
    DoEvents
  Loop
  
  CloseHandle hSnap ' Snapshot zerstören
  GetProcessByName = 0
End Function

Public Function GetProcessName(name As String) As String
 Dim test As String
  Dim RetVal As Long
  Dim hSnap As Long
  Dim PInfo As PROCESSENTRY32

  ' Snapshot vom gesamten System erstellen
  hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
  If hSnap = -1 Then
    MsgBox "Der System-Snapshot konnte nicht erstellt werden.", _
    vbInformation, "Fehler"
    Exit Function
  End If

  PInfo.dwSize = Len(PInfo)
  RetVal = Process32First(hSnap, PInfo) ' ersten Prozess ermitteln

  Do Until RetVal = 0
    With PInfo
      .szExeFile = Trim$(Left$(.szExeFile, lstrlen(.szExeFile))) _
      ' VBNullChar abtrennen
      test = LCase(Trim(Mid$(.szExeFile, InStrRev(.szExeFile, "\") + 1)))
      If test = name Then
        GetProcessName = Trim(Mid$(.szExeFile, InStrRev(.szExeFile, "\") + 1))
        Exit Function
      End If
    End With
    
    RetVal = Process32Next(hSnap, PInfo) ' nächsten Prozess ermitteln
    DoEvents
  Loop
  
  CloseHandle hSnap ' Snapshot zerstören
  GetProcessName = ""
End Function

Public Function FindProcess(Process) As Long
   Dim res As Long, objProcess, objWMIService, colProcesses
   res = 0
   Set objWMIService = GetObject("winmgmts:")
   Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process")
   For Each objProcess In colProcesses
       If LCase(Process) = LCase(objProcess.Caption) Then res = res + 1
   Next
   FindProcess = res
End Function

Public Sub Check_AP()
'If Not LimitEXE Then Exit Sub
'If LCase(App.EXEName) <> "ap" Then
'    MsgBox "AP.EXE  Only!" & App.EXEName
'    frmMain.Form_Unload 0
'ElseIf FindProcess("ap.exe") > 2 Then
'    MsgBox "2EXE Only!"
'    frmMain.Form_Unload 0'
'End If
End Sub

