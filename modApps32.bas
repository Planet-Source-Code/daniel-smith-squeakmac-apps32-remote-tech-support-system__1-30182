Attribute VB_Name = "modApps32"
'Apps32 module
'Use RefreshList to put running apps into a listbox, and KillApp to end one
'by SqueakMac (squeak5@mediaone.net)

Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)
Private Const TH32CS_SNAPPROCESS As Long = 2&
Private Const MAX_PATH As Integer = 260

Private Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

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
    szExeFile As String * MAX_PATH
End Type

Public Function KillTask(ProgramPath As String) As Boolean
On Local Error GoTo Finish

Const PROCESS_ALL_ACCESS = 0
Const TH32CS_SNAPPROCESS As Long = 2&

Dim uProcess As PROCESSENTRY32
Dim rProcessFound As Long
Dim hSnapshot As Long
Dim szExename As String
Dim exitCode As Long
Dim myProcess As Long
Dim appCount As Integer
Dim i As Integer

appCount = 0
uProcess.dwSize = Len(uProcess)
hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
rProcessFound = ProcessFirst(hSnapshot, uProcess)

Do While rProcessFound
    i = InStr(1, uProcess.szExeFile, Chr(0))
    szExename = LCase$(Left$(uProcess.szExeFile, i - 1))
    If LCase$(szExename) = LCase$(ProgramPath) Then
        KillTask = True
        appCount = appCount + 1
        myProcess = OpenProcess(PROCESS_ALL_ACCESS, False, uProcess.th32ProcessID)
        AppKill = TerminateProcess(myProcess, exitCode)
        Call CloseHandle(myProcess)
    End If
    rProcessFound = ProcessNext(hSnapshot, uProcess)
Loop
Call CloseHandle(hSnapshot)

Finish:
End Function


Public Function AppsRunning(List As ListBox) As String
Dim hSnapshot As Long
Dim uProcess As PROCESSENTRY32
Dim r As Long

List.Clear
hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)

If hSnapshot = 0 Then Exit Function

uProcess.dwSize = Len(uProcess)
r = ProcessFirst(hSnapshot, uProcess)

Do While r
    List.AddItem uProcess.szExeFile
    r = ProcessNext(hSnapshot, uProcess)
Loop

Call CloseHandle(hSnapshot)

For i = 0 To List.ListCount
    AppsRunning = AppsRunning & List.List(i) & vbCrLf
Next i
End Function



