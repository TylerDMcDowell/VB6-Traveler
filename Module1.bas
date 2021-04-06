Attribute VB_Name = "Module1"
Option Explicit
Public prgdll As Class1
Private frmMap As frmComm
Public Sub Main()
    Randomize Timer
    Set prgdll = CreateObject("Project1dll.Class1")
    Set frmMap = New frmComm
    frmMap.Show vbModal
    Set prgdll = Nothing
    Set frmMap = Nothing
    End
End Sub
