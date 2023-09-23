Attribute VB_Name = "ReceiverModule"
'This module contains this program's core procedures.
Option Explicit

'Defines the Microsoft Windows API constants and functions used by this program.
Private Declare Function CallNamedPipeA Lib "Kernel32.dll" (ByVal lpNamedPipeName As String, lpInBuffer As Any, ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, lpBytesRead As Long, ByVal nTimeOut As Long) As Long

'This procedure is executed when this program started.
Public Sub Main()
On Error GoTo ErrorTrap
Dim BytesRead As Long
Dim InputBuffer() As Byte

   ReDim InputBuffer(&H0& To &H100&) As Byte
   CheckForError CallNamedPipeA("\\.\pipe\namedpipe", CLng(&H0&), CLng(&H0&), InputBuffer(0), UBound(InputBuffer()) - LBound(InputBuffer()), BytesRead, CLng(30000))

   If BytesRead = &H0& Then
      MsgBox "No message was received", vbExclamation
   Else
      ReDim Preserve InputBuffer(LBound(InputBuffer()) To BytesRead) As Byte
      MsgBox "The following message was received: " & vbCr & """" & CStr(InputBuffer()) & ".""", vbInformation
   End If
EndRoutine:
   Exit Sub
   
ErrorTrap:
   MsgBox Err.Description & vbCr & "Error code: " & CStr(Err.Number), vbExclamation
   Resume EndRoutine
End Sub


