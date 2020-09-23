VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "AboutCAL"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4620
   Icon            =   "aboutsmallcal.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   $"aboutsmallcal.frx":030A
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "www.question.chatbook.com"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   1440
      Width           =   4695
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
Form1.Show
End Sub

Private Sub Label2_Click()
Shell "Explorer http://www.question.chatbook.com"
End Sub
