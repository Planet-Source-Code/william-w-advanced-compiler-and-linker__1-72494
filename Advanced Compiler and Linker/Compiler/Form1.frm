VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Advanced Compiler 1.0"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6585
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   6585
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   2685
      TabIndex        =   1
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   4215
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form1.frx":08CA
      Top             =   0
      Width           =   6615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Just Shows Usage Options


Private Sub Command1_Click()

   Unload Me

End Sub

Private Sub Form_Load()

   Text1.Text = "Rename Original C2.exe to " & OrigCompiler & vbCrLf & Text1.Text
   Me.Show

End Sub

