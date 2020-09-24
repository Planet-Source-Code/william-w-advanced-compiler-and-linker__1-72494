Attribute VB_Name = "LinkerModule"
Option Explicit
Public SpecialLink As Boolean
Private intPos As Integer
Private fCPL As Boolean
Private fResource As Boolean
Private strCmd As String
Private strPath As String
Private strFileContents As String
Private strDefFile As String
Private strResFile As String
Private oFS As New Scripting.FileSystemObject
Private fld As Folder
Private fil As File
Private ts As TextStream
Private tsDef As TextStream
Private AdvLink As Boolean
Private AdvTxt As String
Private Const LineSep = "------------------------------------------------------------"
Public Const OrigLinker = "LinkLnk.exe"
'<include Microsoft Scripting Runtime (scrrun.dll)>
'Advanced Linker 1.3


Public Sub DoLink()
Dim LinkRet As String
   On Error Resume Next

   If SpecialLink Then
      ' Determine contents of .DEF file
      Set tsDef = oFS.OpenTextFile(strDefFile)

      strFileContents = tsDef.ReadAll

      If InStr(1, strFileContents, "CplApplet", vbTextCompare) > 0 Then
         fCPL = True
      End If

      ' Add module definition before /DLL switch
      intPos = InStr(1, strCmd, "/DLL", vbTextCompare)

      If intPos > 0 Then
         strCmd = Left(strCmd, intPos - 1) & " /DEF:" & Chr(34) & strDefFile & Chr(34) & " " & _
            Mid(strCmd, intPos)
      End If

      ' Include .RES file if one exists

      If fResource Then
         intPos = InStr(1, strCmd, "/ENTRY", vbTextCompare)
         strCmd = Left(strCmd, intPos - 1) & Chr(34) & strResFile & Chr(34) & " " & Mid(strCmd, _
            intPos)
      End If

      ' If Control Panel applet, change "DLL" extension to "CPL"

      If fCPL Then
         strCmd = Replace(strCmd, ".dll", ".cpl", 1, , vbTextCompare)
      End If

      ' Write linker options to output file
      ts.WriteLine "Command line arguments after modification:"
      ts.WriteLine LineSep
      ts.WriteLine strCmd
      ts.WriteLine LineSep
      ts.WriteBlankLines 1
   End If

   If AdvLink Then
      ts.WriteLine "CMDFILE.TXT Found, New Options Used:"
      ts.WriteLine LineSep
      ts.WriteLine AdvTxt
      ts.WriteLine LineSep
      ts.WriteBlankLines 1

   End If

   ts.WriteLine "Calling " & OrigLinker & " (linker):"
   
   
   
   
   
   LinkRet = GetCommandOutput(OrigLinker & " " & strCmd, True, True)
   'CmdOutput a wonderful module indeed
   ts.WriteBlankLines 1
   ts.WriteLine "Linker Output:"
   ts.WriteLine LineSep
   ts.WriteLine LinkRet
   ts.WriteLine LineSep
   ts.WriteBlankLines 1
   ts.WriteLine "File Type:"
   ts.WriteLine LineSep

   If InStr(1, strCmd, ".DLL" & Chr(34), vbTextCompare) Then
      If SpecialLink = True Then
         ts.WriteLine "Windows DLL File"
       Else
         ts.WriteLine "VB Active X DLL File"
      End If

   End If

   If InStr(1, strCmd, ".CPL" & Chr(34), vbTextCompare) Then
      If SpecialLink = True Then
         ts.WriteLine "Windows Control Panel Applet File"
       Else
         ts.WriteLine "Worthless Control Panel Applet File"
      End If

   End If

   If InStr(1, strCmd, ".EXE" & Chr(34), vbTextCompare) Then ts.WriteLine "Standard Windows EXE" & _
      " file"
   ts.WriteLine LineSep
   ts.WriteBlankLines 2
   ts.WriteLine "Returned from linker call at " & Date & " " & Time()
   ts.Close
   Unload Form1

End Sub

Public Sub Main()
Dim TmpTxt As String
Dim IntPos2 As Integer
Dim ProjName As String
   On Error GoTo LinkErr

   strCmd = Command
RetryArgs:

   If strCmd = "" Then
      'Show Command Arguments Window
      Load Form2
      Exit Sub
   End If

   ' Determine if .DEF file exists
   '
   ' Extract path from first .obj argument
    
   intPos = InStr(1, strCmd, ".OBJ", vbTextCompare)
   strPath = Mid(strCmd, 2, intPos + 2)
   intPos = InStrRev(strPath, "\")
   strPath = Left(strPath, intPos - 1)
   'get the exe name without .exe
   intPos = InStr(1, strCmd, "/out:", vbTextCompare)
   ProjName = Mid(strCmd, intPos + 6, InStr(intPos + 6, strCmd, Chr(34)) - intPos - 6)
   IntPos2 = InStrRev(ProjName, "\")
   ProjName = Mid(ProjName, IntPos2 + 1, Len(ProjName) - IntPos2 - 4)
   
   If LCase(Dir(strPath & "\COMPFILE.TXT")) = LCase("COMPFILE.TXT") Then
   'Save The Main .Obj file if user wants them saved
   'vb makes this file after the compiler runs so i had to put this here
   'I didn't put all the save code here because i would miss the time for intermediate files
   On Error Resume Next
      Open strPath & "\COMPFILE.TXT" For Input As 1
      'Get Advanced Compiler Options if they exist

      If EOF(1) = False Then

         Do While EOF(1) = False
            Input #1, TmpTxt
            AdvTxt = AdvTxt & " " & TmpTxt
         Loop
     
         'Lets Get All Comands That want added or removed
         If InStr(1, AdvTxt, "*-save", vbTextCompare) <> 0 Then
         MkDir strPath & "\Obj Files"
         Kill strPath & "\Obj Files\" & ProjName & ".obj" 'get rid of the old one
         FileCopy strPath & "\" & ProjName & ".obj", strPath & "\Obj Files\" & ProjName & ".obj"
         End If
         End If

      Close #1
      
      AdvTxt = vbNullString
      On Error GoTo LinkErr
      End If
   
   
   If LCase(Dir(strPath & "\CMDFILE.TXT")) = LCase("CMDFILE.TXT") Then

      strCmd = strCmd & " @" & Chr(34) & strPath & "\CMDFILE.TXT" & Chr(34) ' @"PATH\CMDFILE.TXT"
      Open strPath & "\CMDFILE.TXT" For Input As 1
      'Get Advanced Compiler Options if they exist

      Do While EOF(1) = False
         Input #1, TmpTxt
         AdvTxt = AdvTxt & " " & TmpTxt
      Loop
      If InStr(1, AdvTxt, "/") <> 0 Then AdvLink = True
      Close #1

   End If

   Set ts = oFS.CreateTextFile(strPath & "\LinkLog.txt")
   'Start Log File
   ts.WriteLine "Advanced Linker 1.3"
   ts.WriteBlankLines 1
   ts.WriteLine "Beginning execution at " & Date & " " & Time()
   ts.WriteBlankLines 2
   ts.WriteLine "Command line arguments to LINK call:"
   ts.WriteLine LineSep
   ts.WriteLine strCmd
   ts.WriteLine LineSep
   ts.WriteBlankLines 1

   ' Open folder
   Set fld = oFS.GetFolder(strPath)

   ' Get files in folder

   For Each fil In fld.Files

      If UCase(oFS.GetExtensionName(fil)) = "DEF" Then
         strDefFile = fil
         SpecialLink = True
      End If

      If UCase(oFS.GetExtensionName(fil)) = "RES" Then
         strResFile = fil
         fResource = True
      End If

      If SpecialLink And fResource Then Exit For
   Next

   ' Change command line arguments if flag set

   If SpecialLink = True Then
      If InStr(1, strCmd, ".DLL" & Chr(34), vbTextCompare) Or InStr(1, strCmd, ".CPL" & Chr(34), _
         vbTextCompare) Then
         Load Form1
       Else
         DoLink ' call actual linker
      End If

    Else
      DoLink 'call actual Linker
   End If

   Exit Sub
LinkErr:

   strCmd = InputBox("Linker Error #" & Err.Number & " Check Arguments" & vbCrLf & Err.Description _
      & vbCrLf & "Change arguments here then press OK to try again" & vbCrLf & "Press Cancel to" & _
      " exit.", "Linker Helper 1.3 Error", Command$)

   If strCmd <> Command$ Then GoTo RetryArgs
   
   Exit Sub
   
End Sub

