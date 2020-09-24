Attribute VB_Name = "CompilerModule"
Option Explicit
Private intPos1 As Integer
Private intPos2 As Integer
Private strCmd As String
Private strPath As String
Private oFS As New Scripting.FileSystemObject
Private ts As TextStream
Private AdvCompile As Boolean
Private AdvTxt As String
Private ObjName As String
Private ILName As String
Private ILPath As String
Private Const LineSep = "------------------------------------------------------------"
Public Const OrigCompiler = "C3.exe"
'<include Microsoft Scripting Runtime (scrrun.dll)>
'Advanced Compiler 1.0


Public Sub DoCompile()

  Dim a As Long
  Dim b As Long
  Dim RemStr As String
  Dim CompRet As String
  Dim SaveFiles As Boolean

   On Error Resume Next
   If InStr(1, AdvTxt, "*") <> 0 And AdvCompile = True Then
      'Remove items from Command Line
      a = InStr(1, AdvTxt, "*")

      ts.WriteLine "Items Desired To Remove:"
      ts.WriteLine LineSep

      Do While a <> 0
         b = (InStr(a, AdvTxt, " ") - a)
         If b < 1 Then b = Len(AdvTxt)
         RemStr = Mid(AdvTxt, a + 1, b - 1)
         ts.Write RemStr & " "

         If LCase(RemStr) = "-il" Or LCase(RemStr) = "-f" Or LCase(RemStr) = "-fo" Then
            b = InStr(1, strCmd, RemStr, vbTextCompare)

            If b <> 0 Then
               strCmd = Replace(strCmd, Mid(strCmd, b, InStr(b + 5, strCmd, Chr(34)) - b + 1), "", _
                  1, 1, vbTextCompare)
            End If

          Else
            If LCase(RemStr) = "-save" Then SaveFiles = True
            strCmd = Replace(strCmd, RemStr, "", 1, 1, vbTextCompare)

         End If
         AdvTxt = Replace(AdvTxt, "*" & RemStr, "", 1, 1, vbTextCompare)
         If LCase(RemStr) = "all" Then strCmd = ""
         a = InStr(1, AdvTxt, "*")
      Loop

      ts.WriteLine " "
      ts.WriteLine LineSep
   End If

   If AdvCompile Then
      If AdvTxt <> "" Then strCmd = strCmd & " " & AdvTxt
      ts.WriteBlankLines 1
      ts.WriteLine "COMPFILE.TXT Found, New Options Used:"
      ts.WriteLine LineSep
      ts.WriteLine AdvTxt
      ts.WriteLine LineSep
      ts.WriteBlankLines 1
      ts.WriteLine "Command line arguments after modification:"
      ts.WriteLine LineSep
      ts.WriteLine strCmd
      ts.WriteLine LineSep
      ts.WriteBlankLines 1
   End If

   If SaveFiles = True Then
      'Save all the files that C2 uses to make .Obj files
      MkDir strPath & "\Obj Files"
      Kill strPath & "\Obj Files\" & ObjName & "DB"
      Kill strPath & "\Obj Files\" & ObjName & "EX"
      Kill strPath & "\Obj Files\" & ObjName & "GL"
      Kill strPath & "\Obj Files\" & ObjName & "IN"
      Kill strPath & "\Obj Files\" & ObjName & "SY"
      FileCopy ILPath & "\" & ILName & "DB", strPath & "\Obj Files\" & ObjName & "DB"
      FileCopy ILPath & "\" & ILName & "EX", strPath & "\Obj Files\" & ObjName & "EX"
      FileCopy ILPath & "\" & ILName & "GL", strPath & "\Obj Files\" & ObjName & "GL"
      FileCopy ILPath & "\" & ILName & "IN", strPath & "\Obj Files\" & ObjName & "IN"
      FileCopy ILPath & "\" & ILName & "SY", strPath & "\Obj Files\" & ObjName & "SY"
   End If

   ts.WriteLine "Calling " & OrigCompiler & " (Compiler):"
   CompRet = GetCommandOutput(OrigCompiler & " " & strCmd, True, True)
   
   If SaveFiles = True Then
      'Save the obj files
      Kill strPath & "\Obj Files\" & ObjName & ".obj"
      FileCopy strPath & "\" & ObjName & ".obj", strPath & "\Obj Files\" & ObjName & ".obj"
   End If

   'CmdOutput a wonderful module indeed
   ts.WriteBlankLines 1
   ts.WriteLine "Compiler Output:"
   ts.WriteLine LineSep
   ts.WriteLine CompRet
   ts.WriteLine LineSep
   ts.WriteBlankLines 1
   ts.WriteLine "Returned from Compiler call at " & Date & " " & Time()
   ts.Close
   Unload Form1
   
End Sub

Public Sub Main()

  Dim TmpTxt As String

   On Error GoTo CompErr
   strCmd = Command
   'fix that odd ms error
   strCmd = Replace(strCmd, "W 3", "W3", 1, 1, vbTextCompare)
RetryArgs:

   If strCmd = "" Then
      'Show Command Arguments Window
      Load Form1
      Exit Sub
   End If

   intPos1 = InStr(1, strCmd, "-il", vbTextCompare)
   ILPath = Mid(strCmd, intPos1 + 5, InStr(intPos1 + 5, strCmd, Chr(34)) - intPos1 - 5)
   intPos2 = InStrRev(ILPath, "\")
   ILName = Mid(ILPath, intPos2 + 1)
   ILPath = Mid(ILPath, 1, intPos2 - 1)
   intPos1 = InStr(1, strCmd, ".OBJ", vbTextCompare)
   'Extract The Name From the command arguments
   strPath = Mid(strCmd, 1, intPos1)
   intPos2 = InStrRev(strPath, "\")
   ObjName = Mid(strCmd, intPos2 + 1, intPos1 - intPos2 - 1)
   ' Extract path from first .obj argument
   intPos2 = InStrRev(strCmd, Chr(34), intPos1)
   strPath = Mid(strCmd, intPos2 + 1, intPos1 - intPos2)
   strPath = Left(strPath, InStrRev(strPath, "\") - 1)

   If LCase(Dir(strPath & "\COMPFILE.TXT")) = LCase("COMPFILE.TXT") Then

      Open strPath & "\COMPFILE.TXT" For Input As 1
      'Get Advanced Compiler Options if they exist

      If EOF(1) = False Then

         Do While EOF(1) = False
            Input #1, TmpTxt
            AdvTxt = AdvTxt & " " & TmpTxt
         Loop

         'Lets Get All Commands That want added or removed
         If InStr(1, AdvTxt, "-") + InStr(1, AdvTxt, "*") + InStr(1, AdvTxt, "/") <> 0 Then _
            AdvCompile = True
         AdvTxt = Replace(AdvTxt, "/", "-") 'This Compiler uses '-' instead of '/' so we'll replace
         '   any / with -
      End If

      Close #1
    Else
      Shell OrigCompiler & " " & Command$
      'if there isn't a compfile then no logs will be made so we can just
      'shell the compiler and not worry about getting a output from it
      Exit Sub
   End If

   Set ts = oFS.CreateTextFile(strPath & "\" & ObjName & "CompileLog.txt")
   'Start Log File with object name in front
   ts.WriteLine "Compile Helper 1.0"
   ts.WriteBlankLines 1
   ts.WriteLine "Beginning execution at " & Date & " " & Time()
   ts.WriteBlankLines 2
   ts.WriteLine "Command line arguments to C2.exe call:"
   ts.WriteLine LineSep
   ts.WriteLine strCmd
   ts.WriteLine LineSep
   ts.WriteBlankLines 1

   DoCompile 'call actual Compiler

   Exit Sub
CompErr:

   strCmd = InputBox("Compile Error #" & Err.Number & " Check Arguments" & vbCrLf & Err.Description _
      & vbCrLf & "Change arguments here then press OK to try again" & vbCrLf & "Press Cancel to" & _
      " exit.", "Compile Helper 1.0 Error", Command$)

   If strCmd <> Command$ Then GoTo RetryArgs
   Exit Sub

End Sub

