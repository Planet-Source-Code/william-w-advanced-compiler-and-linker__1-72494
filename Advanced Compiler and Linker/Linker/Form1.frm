VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Advanced Linker 1.3"
   ClientHeight    =   900
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6765
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   900
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3120
      Top             =   480
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "Form1.frx":08CA
      Top             =   0
      Visible         =   0   'False
      Width           =   6615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "No (10)"
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Yes"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Gives option to make a .dll or .cpl file when .def exists


Private Sub Command1_Click()

   Timer1.Enabled = False
   SpecialLink = True
   DoLink

End Sub

Private Sub Command2_Click()

   Timer1.Enabled = False
   SpecialLink = False
   DoLink

End Sub

Private Sub Form_Load()

   If InStr(1, Command, ".DLL" & Chr(34), vbTextCompare) Then Label1.Caption = "This is an Active X" & _
      " DLL would you like to use the .Obj and .Def files to make a windows DLL?"
   If InStr(1, Command, ".CPL" & Chr(34), vbTextCompare) Then Label1.Caption = "This is a VB CPL" & _
      " would you like to use the .Obj and .Def files to make a windows CPL?"
   Me.Show

End Sub

Private Sub Timer1_Timer()

  Static cd As Integer

   Command2.Caption = "No (" & 10 - cd & ")"
   cd = cd + 1

   If cd > 10 Then

      Timer1.Enabled = False
      SpecialLink = False
      DoLink

   End If

End Sub

