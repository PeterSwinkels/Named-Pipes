Attribute VB_Name = "APIErrorModule"
'This module contains API error handling procedures.
Option Explicit

'Defines the Microsoft Windows API constants, functions, and structures used.
Private Const ERROR_SUCCESS As Long = 0
Private Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000&

Private Declare Function FormatMessageA Lib "Kernel32.dll" (ByVal dwFlags As Long, lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

'Defines the constants used.
Private Const MAX_STRING As Long = 65535   'Defines the maximum length allowed for a string buffer.

'This procedure checks whether an error has occurred during the most recent Windows API call.
Public Function CheckForError(ReturnValue As Long) As Long
Dim Description As String
Dim ErrorCode As Long
Dim Length As Long
Dim Message As String

   ErrorCode = Err.LastDllError
   Err.Clear
   
   If Not ErrorCode = ERROR_SUCCESS Then
      Description = String$(MAX_STRING, vbNullChar)
      Length = FormatMessageA(FORMAT_MESSAGE_FROM_SYSTEM, CLng(0), ErrorCode, CLng(0), Description, Len(Description), CLng(0))
      If Length = 0 Then
         Description = "No description."
      ElseIf Length > 0 Then
         Description = Left$(Description, Length - 1)
      End If
     
      Message = "API error code: " & CStr(ErrorCode) & " - " & Description
      Message = Message & "Return value: " & CStr(ReturnValue)
      MsgBox Message, vbExclamation
   End If
   
   CheckForError = ReturnValue
End Function


